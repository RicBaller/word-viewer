// Editing a document in the browser, like a word processor.
//
// Typing and text formatting change the HTML of the document directly. After a pause the
// HTML is written back into the XML (commit in docx-edit.js) and a step is added to the undo
// history. Paragraph styles, lists, alignment and tables change the XML; the document is then
// rendered again, so it looks exactly as the styles of the document make it look.

const editor = {
  active: false,
  doc: null, // the open document, as in app.js
  undo: [], // states after each change: { body, selection }
  redo: [],
  typing: false, // HTML changed since the last commit
  composing: false,
  timer: 0,
  lastSelection: null,
  // Set by app.js.
  onRender: () => {}, // the document was rendered again: (element)
  onChange: () => {}, // the content changed (after a commit)
  onInput: () => {}, // the user typed
};

const EDIT_PAUSE = 800; // ms without typing before a commit
const MAX_UNDO = 200;
const BLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6, li, div, td';
const PARAGRAPH_STYLES = ['Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Quote'];
const FONTS = ['Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Garamond', 'Georgia', 'Helvetica', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const HIGHLIGHT_COLORS = [
  ['colorYellow', '#ffff00'],
  ['colorGreen', '#00ff00'],
  ['colorCyan', '#00ffff'],
  ['colorPink', '#ff00ff'],
  ['colorRed', '#ff0000'],
  ['colorGray', '#c0c0c0'],
];

const root = () => editor.doc?.element;
const model = () => editor.doc.model;

// ---------- Starting and stopping ----------

function startEditing(doc) {
  editor.doc = doc;
  editor.active = true;
  prepareEditable(doc.element);
  snapshot(doc.model, doc.element);
  editor.undo = [captureState()];
  editor.redo = [];
  editor.typing = false;
  // The state that matches the file as opened or downloaded: undoing back to it is not a change.
  editor.cleanState = doc.dirty ? null : editor.undo[0];
  document.execCommand('defaultParagraphSeparator', false, 'p');
  updateToolbar();
}

function stopEditing() {
  if (!editor.active) return;
  flushEdits();
  const el = root();
  el.removeAttribute('contenteditable');
  el.classList.remove('editing');
  editor.active = false;
}

function prepareEditable(el) {
  el.contentEditable = 'true';
  el.spellcheck = true;
  el.classList.add('editing');
}

// Writes pending typing into the XML and adds an undo step. Safe to call any time.
function flushEdits() {
  clearTimeout(editor.timer);
  if (!editor.active || !editor.typing) return;
  commit(model(), root());
  editor.typing = false;
  pushState();
  editor.onChange();
}

function captureState() {
  return { body: model().body.cloneNode(true), selection: saveSelection() };
}

function pushState() {
  editor.undo.push(captureState());
  if (editor.undo.length > MAX_UNDO) editor.undo.shift();
  editor.redo = [];
  updateToolbar();
}

function rerender(selection) {
  const { element } = model().render();
  prepareEditable(element);
  snapshot(model(), element);
  editor.onRender(element);
  if (selection) restoreSelection(selection);
  else element.focus({ preventScroll: true });
}

// Runs a command that changes the XML. fn returns nothing, or { focus: xmlNode } to put the
// caret at the start of that node instead of where it was.
function command(fn) {
  if (!editor.active) return;
  flushEdits();
  const selection = saveSelection();
  commit(model(), root());
  const result = fn();
  rerender(result?.focus ? null : selection);
  if (result?.focus) caretAtNode(result.focus);
  pushState();
  editor.onChange();
}

function undo() {
  flushEdits();
  if (editor.undo.length < 2) return;
  editor.redo.push(editor.undo.pop());
  restoreState(editor.undo[editor.undo.length - 1]);
}

function redo() {
  flushEdits();
  const state = editor.redo.pop();
  if (!state) return;
  editor.undo.push(state);
  restoreState(state);
}

function restoreState(state) {
  // States saved before any selection was made have none; then the caret stays where it is.
  const current = saveSelection();
  model().body.replaceWith(state.body.cloneNode(true));
  model().changedParts.add(model().main);
  rerender(state.selection || current);
  updateToolbar();
  editor.onChange();
}

// ---------- Selection ----------

// Blocks that hold text: paragraphs, headings, list items, and cells without paragraphs.
function editableBlocks(el = root()) {
  return [...el.querySelectorAll(BLOCK_SELECTOR)].filter((b) => {
    if (b.closest('.notes, .textbox, [contenteditable="false"]')) return false;
    if (b.tagName === 'DIV' || b.tagName === 'TD') return !b.querySelector(':scope > :is(p, h1, h2, h3, h4, h5, h6, ul, ol, table, div)');
    return true;
  });
}

// Text nodes of a block that count for caret positions: not in nested lists or non-editable parts.
function textNodes(block) {
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
    acceptNode(n) {
      if (n.nodeType === Node.TEXT_NODE) return NodeFilter.FILTER_ACCEPT;
      if (n !== block && (n.matches('ul, ol, table, .heading-number') || n.getAttribute('contenteditable') === 'false')) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_SKIP;
    },
  });
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  return nodes;
}

function blockOf(node) {
  const el = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  const blocks = editableBlocks();
  for (let n = el; n && n !== root(); n = n.parentElement) if (blocks.includes(n)) return n;
  return null;
}

function offsetIn(block, container, offset) {
  const point = document.createRange();
  point.setStart(container, offset);
  let total = 0;
  for (const t of textNodes(block)) {
    if (t === container) return total + offset;
    if (point.comparePoint(t, 0) === 1) break;
    total += t.length;
  }
  return total;
}

function pointAt(block, offset) {
  const nodes = textNodes(block);
  for (const t of nodes) {
    if (offset <= t.length) return { node: t, offset };
    offset -= t.length;
  }
  const last = nodes[nodes.length - 1];
  return last ? { node: last, offset: last.length } : { node: block, offset: 0 };
}

// The selection as block numbers and text offsets, which survive rendering the document again.
function saveSelection() {
  const sel = getSelection();
  if (!sel.rangeCount || !root()?.contains(sel.anchorNode)) return editor.lastSelection;
  const blocks = editableBlocks();
  const point = (node, offset) => {
    const block = blockOf(node);
    if (!block) return { b: 0, o: 0 };
    return { b: blocks.indexOf(block), o: offsetIn(block, node, offset) };
  };
  return { anchor: point(sel.anchorNode, sel.anchorOffset), focus: point(sel.focusNode, sel.focusOffset) };
}

function restoreSelection(saved) {
  if (!saved) return;
  const blocks = editableBlocks();
  if (!blocks.length) return;
  const resolve = (p) => pointAt(blocks[Math.min(Math.max(p.b, 0), blocks.length - 1)], p.o);
  const a = resolve(saved.anchor);
  const f = resolve(saved.focus);
  root().focus({ preventScroll: true });
  getSelection().setBaseAndExtent(a.node, a.offset, f.node, f.offset);
}

function domOf(node) {
  const nodes = model().renderer.nodes;
  return [...root().querySelectorAll('[data-x]')].find((el) => nodes[el.dataset.x] === node) || null;
}

function caretAtNode(node) {
  const el = domOf(node);
  const block = el && (editableBlocks(el.parentElement || el).find((b) => el.contains(b) || b === el) || el);
  if (!block) return;
  const p = pointAt(block, 0);
  root().focus({ preventScroll: true });
  getSelection().collapse(p.node, p.offset);
  block.scrollIntoView({ block: 'nearest' });
}

function currentRange() {
  const sel = getSelection();
  if (!sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  return root()?.contains(range.commonAncestorContainer) ? range : null;
}

// The part of a block that is its own text: a list item without the lists nested in it.
function ownRange(block) {
  const r = document.createRange();
  r.selectNodeContents(block);
  const nested = block.tagName === 'LI' && [...block.children].find((c) => c.tagName === 'UL' || c.tagName === 'OL');
  if (nested) r.setEndBefore(nested);
  return r;
}

function selectedBlocks() {
  const range = currentRange() || (editor.lastSelection && (restoreSelection(editor.lastSelection), currentRange()));
  if (!range) return [];
  return editableBlocks().filter((b) => {
    const own = ownRange(b);
    return range.comparePoint(own.endContainer, own.endOffset) !== -1 && range.comparePoint(own.startContainer, own.startOffset) !== 1;
  });
}

function selectedParagraphs() {
  return selectedBlocks()
    .map((b) => nodeOf(model(), b))
    .filter((n) => n?.localName === 'p');
}

function currentCell() {
  const range = currentRange();
  const node = range?.startContainer;
  const td = (node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement)?.closest('td');
  return td && root().contains(td) ? td : null;
}

// ---------- Text formatting ----------

const hasDecoration = (text, value) => {
  for (let el = text.parentElement; el && el !== root(); el = el.parentElement) {
    if (getComputedStyle(el).textDecorationLine.includes(value)) return true;
    if (el.matches(BLOCK_SELECTOR)) break;
  }
  return false;
};

const INLINE = {
  bold: {
    tags: ['STRONG', 'B'],
    style: 'fontWeight',
    read: (t) => Number(getComputedStyle(t.parentElement).fontWeight) >= 600,
    on: () => document.createElement('strong'),
    off: () => styled({ fontWeight: 'normal' }),
  },
  italic: {
    tags: ['EM', 'I'],
    style: 'fontStyle',
    read: (t) => getComputedStyle(t.parentElement).fontStyle === 'italic',
    on: () => document.createElement('em'),
    off: () => styled({ fontStyle: 'normal' }),
  },
  underline: {
    tags: ['U'],
    decoration: 'underline',
    read: (t) => hasDecoration(t, 'underline'),
    on: () => document.createElement('u'),
    off: () => styled({ textDecorationLine: 'none' }),
  },
  strike: {
    tags: ['S', 'STRIKE', 'DEL'],
    decoration: 'line-through',
    read: (t) => hasDecoration(t, 'line-through'),
    on: () => document.createElement('s'),
    off: () => styled({ textDecorationLine: 'none' }),
  },
  superscript: {
    tags: ['SUP'],
    other: 'subscript',
    read: (t) => !!t.parentElement.closest('sup'),
    on: () => document.createElement('sup'),
  },
  subscript: {
    tags: ['SUB'],
    other: 'superscript',
    read: (t) => !!t.parentElement.closest('sub'),
    on: () => document.createElement('sub'),
  },
  color: { style: 'color', make: (v) => styled({ color: v }) },
  highlight: { style: 'backgroundColor', make: (v) => styled({ backgroundColor: v }) },
  size: { style: 'fontSize', make: (v) => styled({ fontSize: `${v}pt` }) },
  font: { style: 'fontFamily', make: (v) => styled({ fontFamily: fontStack(v) }) },
};

function styled(style) {
  const span = document.createElement('span');
  Object.assign(span.style, style);
  return span;
}

// Splits text nodes at the edges of range. Returns the text nodes inside it.
function textSegments(range) {
  let { startContainer: sc, startOffset: so, endContainer: ec, endOffset: eo } = range;
  if (ec.nodeType === Node.TEXT_NODE && eo < ec.length) ec.splitText(eo);
  if (sc.nodeType === Node.TEXT_NODE && so > 0) {
    const rest = sc.splitText(so);
    if (ec === sc) {
      ec = rest;
      eo -= so;
    }
    sc = rest;
    so = 0;
  }
  const r = document.createRange();
  r.setStart(sc, so);
  r.setEnd(ec, eo);
  const scope = r.commonAncestorContainer;
  if (scope.nodeType === Node.TEXT_NODE) return scope.length ? [scope] : [];
  const out = [];
  for (const block of editableBlocks(scope.nodeType === Node.ELEMENT_NODE ? scope : scope.parentElement).concat(blockOf(scope) ? [blockOf(scope)] : [])) {
    for (const t of textNodes(block)) if (t.length && r.intersectsNode(t) && !out.includes(t)) out.push(t);
  }
  return out;
}

// Splits the inline elements around text so it is the only content of each of them, up to block.
function isolate(text, block) {
  let node = text;
  while (node.parentNode && node.parentNode !== block) {
    const parent = node.parentNode;
    if (node.previousSibling) {
      const before = parent.cloneNode(false);
      while (parent.firstChild !== node) before.appendChild(parent.firstChild);
      parent.before(before);
    }
    if (node.nextSibling) {
      const after = parent.cloneNode(false);
      while (node.nextSibling) after.appendChild(node.nextSibling);
      parent.after(after);
    }
    node = parent;
  }
}

function unwrap(el) {
  el.replaceWith(...el.childNodes);
}

// Removes a formatting from the elements between text and its block.
function clearFormat(def, text, block) {
  let el = text.parentElement;
  while (el && el !== block) {
    const up = el.parentElement;
    if (def.tags?.includes(el.tagName)) unwrap(el);
    else {
      if (def.style) el.style[def.style] = '';
      if (def.decoration && el.style.textDecorationLine) {
        el.style.textDecorationLine = el.style.textDecorationLine.replace(def.decoration, '').replace('none', '').trim();
      }
      if (el.tagName === 'SPAN' && !el.attributes.length) unwrap(el);
      else if (el.tagName === 'SPAN' && el.getAttribute('style') === '') el.removeAttribute('style');
    }
    el = up;
  }
}

// Merges neighbouring elements with the same formatting and drops empty ones.
function tidyInline(block) {
  const sameTag = (a, b) => a.nodeType === Node.ELEMENT_NODE && b?.nodeType === Node.ELEMENT_NODE && a.cloneNode(false).outerHTML.replace(/ data-r="\d+"/, '') === b.cloneNode(false).outerHTML.replace(/ data-r="\d+"/, '');
  const walk = (el) => {
    for (const c of [...el.children]) walk(c);
    for (const c of [...el.childNodes]) {
      if (c.nodeType !== Node.ELEMENT_NODE || c.getAttribute('contenteditable') === 'false') continue;
      if (/^(SPAN|STRONG|B|EM|I|U|S|STRIKE|SUP|SUB|A)$/.test(c.tagName) && !c.childNodes.length) c.remove();
      else if (c.tagName === 'SPAN' && !c.attributes.length) unwrap(c);
    }
    let c = el.firstChild;
    while (c) {
      const next = c.nextSibling;
      if (next && sameTag(c, next) && next.getAttribute('contenteditable') !== 'false' && !/^(BR|IMG|HR)$/.test(c.tagName)) {
        c.append(...next.childNodes);
        next.remove();
        continue;
      }
      c = next;
    }
  };
  walk(block);
  block.normalize();
}

// Applies formatting prop to the selection: toggles for bold and the like, sets a value for
// color, highlight, size and font ('none' or 'auto' removes it).
function formatSelection(prop, value) {
  if (!editor.active) return;
  let range = currentRange();
  if (!range) return;
  if (range.collapsed) {
    const word = wordRange(range);
    if (!word) return pendingFormat(prop, value, range);
    range = word;
  }
  const def = INLINE[prop];
  const segments = textSegments(range).filter((t) => blockOf(t));
  if (!segments.length) return;
  const firstBlock = blockOf(segments[0]);
  const lastBlock = blockOf(segments[segments.length - 1]);
  const blocks = editableBlocks();
  const saved = {
    anchor: { b: blocks.indexOf(firstBlock), o: offsetIn(firstBlock, segments[0], 0) },
    focus: { b: blocks.indexOf(lastBlock), o: offsetIn(lastBlock, segments[segments.length - 1], segments[segments.length - 1].length) },
  };
  const toggle = !!def.on;
  const desired = toggle ? !segments.every((t) => def.read(t)) : value;
  const touched = new Set();
  for (const t of segments) {
    const block = blockOf(t);
    touched.add(block);
    applyFormat(def, t, block, desired);
  }
  touched.forEach(tidyInline);
  restoreSelection(saved);
  inlineChanged();
}

function applyFormat(def, text, block, desired) {
  isolate(text, block);
  clearFormat(def, text, block);
  if (def.other) clearFormat(INLINE[def.other], text, block);
  const wrap = (el) => {
    text.before(el);
    el.append(text);
  };
  if (def.on) {
    if (desired !== def.read(text)) {
      if (desired) wrap(def.on());
      else if (def.off) wrap(def.off());
    }
  } else if (desired && desired !== 'none' && desired !== 'auto') {
    wrap(def.make(desired));
  }
}

// With the caret inside a word, formatting applies to the whole word, as in Word.
function wordRange(range) {
  const node = range.startContainer;
  if (node.nodeType !== Node.TEXT_NODE) return null;
  const word = /[\p{L}\p{N}_'’-]/u;
  let start = range.startOffset;
  let end = start;
  if (!word.test(node.data[start - 1] || '') || !word.test(node.data[start] || '')) return null;
  while (start > 0 && word.test(node.data[start - 1])) start--;
  while (end < node.length && word.test(node.data[end])) end++;
  const r = document.createRange();
  r.setStart(node, start);
  r.setEnd(node, end);
  return r;
}

// Formatting with only a caret: what is typed next gets it. A zero-width space holds the
// formatting until then; it is left out when saving.
function pendingFormat(prop, value, range) {
  const block = blockOf(range.startContainer);
  if (!block) return;
  const def = INLINE[prop];
  const holder = document.createTextNode('​');
  range.insertNode(holder);
  const desired = def.on ? !def.read(holder) : value;
  applyFormat(def, holder, block, desired);
  getSelection().collapse(holder, 1);
  editor.typing = true;
  scheduleFlush();
  updateToolbar();
}

function inlineChanged() {
  editor.typing = true;
  flushEdits();
}

// ---------- Links, page breaks, text ----------

function editLink() {
  if (!editor.active) return;
  const range = currentRange();
  if (!range) return;
  const start = range.startContainer.nodeType === Node.ELEMENT_NODE ? range.startContainer : range.startContainer.parentElement;
  const existing = start.closest('a[href]');
  const answer = prompt(t('linkPrompt'), existing?.getAttribute('href') || 'https://');
  if (answer === null) return;
  const url = answer.trim();
  if (!url) {
    if (existing) unwrap(existing);
    inlineChanged();
    return;
  }
  const href = /^(https?:|mailto:|tel:)/i.test(url) ? url : `https://${url}`;
  const rid = addHyperlink(model(), href);
  if (existing && range.collapsed) {
    existing.href = href;
    existing.dataset.rid = rid;
    delete existing.dataset.x;
    inlineChanged();
    return;
  }
  const target = range.collapsed ? wordRange(range) : range;
  if (!target) return;
  const segments = textSegments(target).filter((t) => blockOf(t));
  const touched = new Set();
  for (const text of segments) {
    const block = blockOf(text);
    touched.add(block);
    isolate(text, block);
    const inner = text.parentElement.closest('a');
    if (inner && block.contains(inner)) unwrap(inner);
    const a = document.createElement('a');
    a.href = href;
    a.target = '_blank';
    a.rel = 'noopener';
    a.dataset.rid = rid;
    text.before(a);
    a.append(text);
  }
  touched.forEach(tidyInline);
  inlineChanged();
}

function insertPageBreak() {
  const range = currentRange();
  if (!range || !blockOf(range.startContainer)) return;
  if (!range.collapsed && crossesTable(range)) return;
  range.deleteContents();
  const hr = document.createElement('hr');
  hr.className = 'page-break';
  hr.contentEditable = 'false';
  range.insertNode(hr);
  const after = document.createTextNode('');
  hr.after(after);
  getSelection().collapse(after, 0);
  inlineChanged();
}

const hasContent = (el) => !!el.textContent.replace(/\u200b/g, '') || !!el.querySelector('img, br, hr, .textbox, .noteref, .math, .image-missing, .field');

// Enter: splits the paragraph at the caret. Browsers do this differently from each other and
// from Word, so it is done here. At the end of a paragraph the new one gets the next style
// (after a heading: normal text). Enter in an empty list item ends the list.
function splitParagraph() {
  let range = currentRange();
  if (!range) return;
  if (!range.collapsed) {
    document.execCommand('delete');
    range = currentRange();
  }
  const block = range && blockOf(range.startContainer);
  if (!block || block.tagName === 'TD') {
    document.execCommand('insertLineBreak');
    return;
  }
  if (block.tagName === 'LI' && !ownRange(block).toString().replace(/\u200b/g, '').trim() && !ownRange(block).cloneContents().querySelector('img')) {
    leaveList(block);
    return;
  }
  const tail = document.createRange();
  tail.setStart(range.startContainer, range.startOffset);
  tail.setEnd(block, block.childNodes.length);
  const frag = tail.extractContents();
  const probe = document.createElement('div');
  probe.append(frag.cloneNode(true));
  let next;
  if (!hasContent(probe) && block.tagName !== 'LI') {
    next = document.createElement('p');
    if (block.dataset.x != null) next.dataset.next = block.dataset.x;
    // It looks right at once, not only after the next render: the same paragraph style, or
    // normal text after a heading.
    const style = /^H[1-6]$/.test(block.tagName) ? model().renderer.normalStyle() : block.getAttribute('style');
    if (style) next.setAttribute('style', style);
  } else {
    next = block.cloneNode(false);
    next.removeAttribute('id');
  }
  next.append(frag);
  // Each part needs something to put the caret in, before any list nested in it.
  const own = (el) => {
    const r = ownRange(el);
    const div = document.createElement('div');
    div.append(r.cloneContents());
    return hasContent(div);
  };
  if (!own(block)) block.insertBefore(document.createElement('br'), [...block.children].find((c) => c.tagName === 'UL' || c.tagName === 'OL') || null);
  if (!own(next)) next.prepend(document.createElement('br'));
  block.after(next);
  const p = pointAt(next, 0);
  getSelection().collapse(p.node, p.offset);
  next.scrollIntoView({ block: 'nearest' });
  onInput();
}

// Ends a list item: one level up, or out of the list at the top level.
function leaveList(li) {
  command(() => {
    const p = nodeOf(model(), li);
    if (p?.localName !== 'p') return;
    const ilvl = num(val(child(child(child(p, 'pPr'), 'numPr'), 'ilvl'))) || 0;
    if (ilvl > 0) indent(model(), [p], -1);
    else toggleList(model(), [p], listKind(model(), p) || 'bullet');
  });
}

function insertPlainText(text) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  lines.forEach((line, i) => {
    if (i) document.execCommand('insertParagraph');
    if (line) document.execCommand('insertText', false, line);
  });
}

// ---------- Tables ----------

function tableCommand(action) {
  const td = currentCell();
  const tc = td && nodeOf(model(), td);
  if (action === 'insert') {
    const [rows, cols] = [3, 3];
    command(() => {
      const para = selectedParagraphs()[0] || [...model().body.children].filter((c) => c.localName === 'p').pop();
      // Inside a table the new table goes after that table.
      let after = para;
      while (after.parentNode && after.parentNode.localName !== 'body') after = after.parentNode;
      const tbl = insertTable(model(), after, rows, cols);
      return { focus: kids(kids(tbl, 'tr')[0], 'tc')[0] };
    });
    return;
  }
  if (!tc || tc.localName !== 'tc') return;
  const tr = tc.parentNode;
  command(() => {
    switch (action) {
      case 'rowAbove':
        return { focus: kids(insertRow(model(), tr, false), 'tc')[0] };
      case 'rowBelow':
        return { focus: kids(insertRow(model(), tr, true), 'tc')[0] };
      case 'colLeft':
        insertColumn(model(), tc, false);
        return { focus: tc.previousElementSibling || tc };
      case 'colRight':
        insertColumn(model(), tc, true);
        return { focus: tc.nextElementSibling || tc };
      case 'deleteRow': {
        const next = tr.nextElementSibling || tr.previousElementSibling;
        deleteRow(model(), tr);
        return next?.parentNode ? { focus: kids(next, 'tc')[0] } : null;
      }
      case 'deleteCol': {
        const next = tc.nextElementSibling || tc.previousElementSibling;
        deleteColumn(model(), tc);
        return next?.parentNode ? { focus: next } : null;
      }
      case 'deleteTable': {
        const tbl = tr.parentNode;
        const next = tbl.nextElementSibling;
        deleteTable(model(), tbl);
        return next?.localName === 'p' ? { focus: next } : null;
      }
    }
    return null;
  });
}

// Tab moves between cells; in the last cell it adds a row, as in Word.
function moveInTable(td, back) {
  const cells = [...td.closest('table').querySelectorAll(':scope > tbody > tr > td')];
  const i = cells.indexOf(td);
  const next = cells[i + (back ? -1 : 1)];
  if (next) {
    const block = editableBlocks(next)[0] || next;
    const range = document.createRange();
    range.selectNodeContents(block);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    return;
  }
  if (!back) tableCommand('rowBelow');
}

// ---------- Events ----------

function scheduleFlush() {
  clearTimeout(editor.timer);
  editor.timer = setTimeout(() => {
    if (!editor.composing) flushEdits();
  }, EDIT_PAUSE);
}

function crossesTable(range) {
  const cellOf = (n) => (n.nodeType === Node.ELEMENT_NODE ? n : n.parentElement)?.closest('td, table');
  return cellOf(range.startContainer) !== cellOf(range.endContainer);
}

// Deleting must not break tables apart: no deleting across cells, and no merging a paragraph
// with a table next to it.
function guardDelete(e) {
  const range = currentRange();
  if (!range) return;
  if (!range.collapsed) {
    if (crossesTable(range)) e.preventDefault();
    return;
  }
  const block = blockOf(range.startContainer);
  if (!block) return;
  const offset = offsetIn(block, range.startContainer, range.startOffset);
  const backward = e.inputType.includes('Backward');
  const atEdge = backward ? offset === 0 : offset === textNodes(block).reduce((n, t) => n + t.length, 0);
  if (!atEdge) return;
  let el = block;
  while (el.parentElement && el.parentElement !== root() && !el.parentElement.matches('td') && !(backward ? el.previousElementSibling : el.nextElementSibling)) el = el.parentElement;
  const neighbour = backward ? el.previousElementSibling : el.nextElementSibling;
  if (!neighbour && el.parentElement?.matches('td')) e.preventDefault();
  if (neighbour?.tagName === 'TABLE') e.preventDefault();
}

function onBeforeInput(e) {
  if (!editor.active) return;
  switch (e.inputType) {
    case 'historyUndo':
      e.preventDefault();
      undo();
      return;
    case 'historyRedo':
      e.preventDefault();
      redo();
      return;
    case 'formatBold':
    case 'formatItalic':
    case 'formatUnderline':
    case 'formatStrikeThrough':
    case 'formatSuperscript':
    case 'formatSubscript': {
      e.preventDefault();
      const prop = { formatBold: 'bold', formatItalic: 'italic', formatUnderline: 'underline', formatStrikeThrough: 'strike', formatSuperscript: 'superscript', formatSubscript: 'subscript' }[e.inputType];
      formatSelection(prop);
      return;
    }
    case 'insertFromDrop':
    case 'insertFromPasteAsQuotation':
      e.preventDefault();
      return;
  }
  if (e.inputType === 'deleteContentBackward') {
    // Backspace at the start of a list item removes the bullet first, as in Word.
    const range = currentRange();
    const block = range?.collapsed && blockOf(range.startContainer);
    if (block?.tagName === 'LI' && offsetIn(block, range.startContainer, range.startOffset) === 0) {
      e.preventDefault();
      leaveList(block);
      return;
    }
  }
  if (e.inputType.startsWith('delete')) guardDelete(e);
  if (e.inputType === 'insertParagraph') {
    e.preventDefault();
    const range = currentRange();
    if (range && !range.collapsed && crossesTable(range)) return;
    splitParagraph();
    return;
  }
  if (e.inputType === 'insertText') {
    const range = currentRange();
    if (range && !range.collapsed && crossesTable(range)) e.preventDefault();
  }
}

function onKeyDown(e) {
  if (!editor.active || e.isComposing) return;
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  const run = (fn) => {
    e.preventDefault();
    fn();
  };
  if (mod && !e.altKey) {
    if (key === 'b') return run(() => formatSelection('bold'));
    if (key === 'i') return run(() => formatSelection('italic'));
    if (key === 'u') return run(() => formatSelection('underline'));
    if (key === 'k') return run(editLink);
    if (key === 'z' && !e.shiftKey) return run(undo);
    if ((key === 'z' && e.shiftKey) || key === 'y') return run(redo);
    if (key === 'enter') return run(insertPageBreak);
    if (key === 'l' && e.shiftKey) return run(() => command(() => toggleList(model(), selectedParagraphs(), 'bullet')));
  }
  // Ctrl+Alt+1–4 for headings, Ctrl+Alt+0 for normal text. The code, not the key: on a Mac
  // Alt changes the character.
  const digit = e.code.match(/^Digit([0-4])$/);
  // AltGr (Ctrl+Alt on Windows) types characters such as { and } on many keyboards: not a shortcut.
  if (mod && e.altKey && digit && !e.getModifierState('AltGraph')) {
    return run(() => setStyle(digit[1] === '0' ? 'Normal' : `Heading${digit[1]}`));
  }
  if (e.key === 'Tab' && !mod && !e.altKey) {
    const range = currentRange();
    const block = range && blockOf(range.startContainer);
    if (!block) return;
    const td = currentCell();
    if (block.tagName === 'LI') return run(() => command(() => indent(model(), selectedParagraphs(), e.shiftKey ? -1 : 1)));
    if (td) return run(() => moveInTable(td, e.shiftKey));
    if (!e.shiftKey) return run(() => document.execCommand('insertText', false, '\t'));
  }
}

function onPaste(e) {
  if (!editor.active) return;
  e.preventDefault();
  const range = currentRange();
  if (range && !range.collapsed && crossesTable(range)) return;
  insertPlainText(e.clipboardData.getData('text/plain'));
}

function onInput() {
  if (!editor.active) return;
  editor.typing = true;
  editor.onInput();
  scheduleFlush();
}

function onSelectionChange() {
  if (!editor.active) return;
  const range = currentRange();
  if (!range) return;
  editor.lastSelection = saveSelection();
  updateToolbar();
}

// Listeners live on the host, so they keep working when the document is rendered again.
function attachEditor(host) {
  host.addEventListener('beforeinput', onBeforeInput);
  host.addEventListener('keydown', onKeyDown);
  host.addEventListener('paste', onPaste);
  host.addEventListener('input', onInput);
  host.addEventListener('compositionstart', () => (editor.composing = true));
  host.addEventListener('compositionend', () => {
    editor.composing = false;
    scheduleFlush();
  });
  host.addEventListener('drop', (e) => {
    if (editor.active && !e.dataTransfer?.files.length) e.preventDefault();
  });
  document.addEventListener('selectionchange', onSelectionChange);
  // Leaving the editor, for example for the toolbar, is a good moment to save the typing.
  host.addEventListener('focusout', () => {
    if (editor.active && editor.typing) scheduleFlush();
  });
}

// ---------- Toolbar ----------

const tools = {};

function setStyle(key) {
  command(() => setParagraphStyle(model(), selectedParagraphs(), key));
}

// Builds the toolbar controls that need data, and wires all of them.
function initToolbar(bar) {
  tools.bar = bar;
  tools.style = bar.querySelector('#styleSelect');
  tools.font = bar.querySelector('#fontSelect');
  tools.size = bar.querySelector('#sizeSelect');
  tools.color = bar.querySelector('#colorInput');
  for (const f of FONTS) tools.font.add(new Option(f, f));
  for (const s of FONT_SIZES) tools.size.add(new Option(String(s), String(s)));
  fillStyleOptions();

  // Buttons keep the selection in the document instead of taking focus.
  bar.addEventListener('mousedown', (e) => {
    if (e.target.closest('button')) e.preventDefault();
  });
  // Selects and the color picker take focus; remember the selection before they do.
  bar.addEventListener('focusin', () => {
    if (currentRange()) editor.lastSelection = saveSelection();
  });
  // Puts the selection back in the document when a control took it away.
  const restore = () => {
    if (!currentRange()) restoreSelection(editor.lastSelection);
  };

  bar.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-cmd]');
    if (!btn || !editor.active) return;
    const cmd = btn.dataset.cmd;
    restore();
    switch (cmd) {
      case 'undo': return undo();
      case 'redo': return redo();
      case 'bold':
      case 'italic':
      case 'underline':
      case 'strike':
      case 'superscript':
      case 'subscript':
        return formatSelection(cmd);
      case 'alignLeft':
      case 'alignCenter':
      case 'alignRight':
      case 'justify': {
        const value = { alignLeft: 'left', alignCenter: 'center', alignRight: 'right', justify: 'both' }[cmd];
        return command(() => setAlignment(model(), selectedParagraphs(), value));
      }
      case 'bullets':
        return command(() => toggleList(model(), selectedParagraphs(), 'bullet'));
      case 'numbering':
        return command(() => toggleList(model(), selectedParagraphs(), 'number'));
      case 'outdent':
        return command(() => indent(model(), selectedParagraphs(), -1));
      case 'indent':
        return command(() => indent(model(), selectedParagraphs(), 1));
      case 'link':
        return editLink();
      case 'pageBreak':
        return insertPageBreak();
      case 'highlight':
        return toggleMenu(btn, [
          ['noHighlight', () => (restore(), formatSelection('highlight', 'none'))],
          ...HIGHLIGHT_COLORS.map(([key, color]) => [key, () => (restore(), formatSelection('highlight', color))]),
        ]);
      case 'table': {
        const inCell = !!currentCell();
        return toggleMenu(btn, [
          ['insertTable', () => (restore(), tableCommand('insert'))],
          ...(inCell
            ? [
                ['insertRowAbove', () => (restore(), tableCommand('rowAbove'))],
                ['insertRowBelow', () => (restore(), tableCommand('rowBelow'))],
                ['insertColumnLeft', () => (restore(), tableCommand('colLeft'))],
                ['insertColumnRight', () => (restore(), tableCommand('colRight'))],
                ['deleteRow', () => (restore(), tableCommand('deleteRow'))],
                ['deleteColumn', () => (restore(), tableCommand('deleteCol'))],
                ['deleteTable', () => (restore(), tableCommand('deleteTable'))],
              ]
            : []),
        ]);
      }
    }
  });

  tools.style.addEventListener('change', () => {
    restore();
    setStyle(tools.style.value);
  });
  tools.font.addEventListener('change', () => {
    restore();
    formatSelection('font', tools.font.value);
  });
  tools.size.addEventListener('change', () => {
    restore();
    formatSelection('size', Number(tools.size.value));
  });
  tools.color.addEventListener('change', () => {
    restore();
    formatSelection('color', tools.color.value);
  });
}

function fillStyleOptions() {
  const current = tools.style.value;
  tools.style.replaceChildren();
  for (const key of PARAGRAPH_STYLES) {
    const label = key.startsWith('Heading') ? t('styleHeading', { n: key.slice(7) }) : t(`style${key}`);
    tools.style.add(new Option(label, key));
  }
  tools.style.value = current || 'Normal';
}

// Shows the formatting at the caret in the toolbar.
function updateToolbar() {
  if (!tools.bar || !editor.active) return;
  const bar = tools.bar;
  bar.querySelector('[data-cmd="undo"]').disabled = editor.undo.length < 2 && !editor.typing;
  bar.querySelector('[data-cmd="redo"]').disabled = !editor.redo.length;
  const range = currentRange();
  if (!range) return;
  const node = range.startContainer;
  const el = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  const block = blockOf(node);
  const text = node.nodeType === Node.TEXT_NODE ? node : null;
  for (const prop of ['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript']) {
    const on = text ? INLINE[prop].read(text) : false;
    bar.querySelector(`[data-cmd="${prop}"]`).setAttribute('aria-pressed', String(on));
  }
  const cs = getComputedStyle(el);
  const size = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
  if (![...tools.size.options].some((o) => o.value === String(size))) tools.size.add(new Option(String(size), String(size)));
  tools.size.value = String(size);
  const font = cs.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
  if (![...tools.font.options].some((o) => o.value === font)) tools.font.add(new Option(font, font));
  tools.font.value = font;
  tools.color.value = `#${toHex(cs.color) || '000000'}`.toLowerCase();
  tools.color.parentElement.style.setProperty('--color', tools.color.value);

  const p = block && nodeOf(model(), block);
  const name = p?.localName === 'p' ? styleNameOf(model(), p) : 'normal';
  const key = PARAGRAPH_STYLES.find((k) => BUILTIN_STYLES[k].name.toLowerCase() === name) || '';
  let other = tools.style.querySelector('option[data-other]');
  if (!key) {
    if (!other) {
      other = new Option('', '');
      other.dataset.other = '';
      tools.style.add(other);
    }
    other.textContent = name;
    other.value = '';
  } else other?.remove();
  tools.style.value = key;

  const align = block ? getComputedStyle(block).textAlign : 'start';
  const alignKey = { center: 'alignCenter', right: 'alignRight', end: 'alignRight', justify: 'justify' }[align] || 'alignLeft';
  for (const k of ['alignLeft', 'alignCenter', 'alignRight', 'justify']) bar.querySelector(`[data-cmd="${k}"]`).setAttribute('aria-pressed', String(k === alignKey));
  const list = block?.tagName === 'LI' ? block.parentElement.tagName : null;
  bar.querySelector('[data-cmd="bullets"]').setAttribute('aria-pressed', String(list === 'UL'));
  bar.querySelector('[data-cmd="numbering"]').setAttribute('aria-pressed', String(list === 'OL'));
}
