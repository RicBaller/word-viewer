// Editing the WordprocessingML behind a rendered document, and writing the .docx again.
//
// The XML of the package (DocxModel in docx.js) is the document; the HTML is a view of it.
// Every rendered block and run points back to its XML node (data-x, data-r). commit() writes
// the edited HTML back into the XML: a paragraph whose HTML did not change keeps its original
// XML untouched, a changed one is rebuilt from the HTML with the original properties as a
// starting point. Commands for styles, lists and tables change the XML directly; the editor
// then renders the document again. saveDocx() zips the package with only changed parts replaced.

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Child element order required by the schema. Word refuses files with properties out of order.
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'];
// Run properties the editor controls; all others are kept from the original run.
const EDITED_RPR = ['b', 'bCs', 'i', 'iCs', 'u', 'strike', 'dstrike', 'vertAlign', 'color', 'highlight', 'shd', 'sz', 'szCs'];
const MARKER_STARTS = ['bookmarkStart', 'commentRangeStart', 'permStart'];
const MARKER_ENDS = ['bookmarkEnd', 'commentRangeEnd', 'permEnd'];

// ---------- XML helpers ----------

function W(doc, name, attrs = {}, ...content) {
  const el = doc.createElementNS(NS.w, `w:${name}`);
  for (const [key, value] of Object.entries(attrs)) el.setAttributeNS(NS.w, `w:${key}`, String(value));
  el.append(...content);
  return el;
}

function removeChildren(parent, ...names) {
  for (const c of [...(parent?.children || [])]) if (c.namespaceURI === NS.w && names.includes(c.localName)) c.remove();
}

// Puts el into parent at the place the schema order asks for, replacing a child of the same name.
function putChild(parent, el, order) {
  removeChildren(parent, el.localName);
  const rank = order.indexOf(el.localName);
  const after = [...parent.children].find((c) => order.indexOf(c.localName) > rank);
  parent.insertBefore(el, after || null);
  return el;
}

function ensurePPr(p) {
  let pPr = child(p, 'pPr');
  if (!pPr) {
    pPr = W(p.ownerDocument, 'pPr');
    p.prepend(pPr);
  }
  return pPr;
}

function dropIfEmpty(el) {
  if (el && !el.children.length && !el.attributes.length) el.remove();
}

function insideTextbox(node) {
  for (let n = node.parentNode; n; n = n.parentNode) if (n.localName === 'txbxContent') return true;
  return false;
}

const isParagraphTag = (tag) => /^(P|H[1-6]|DIV|BLOCKQUOTE|PRE)$/.test(tag);

// ---------- HTML → XML ----------

// Signature of a rendered paragraph, to see whether it changed. A list item leaves out the
// lists nested in it: those are paragraphs of their own.
function signature(el) {
  if (el.tagName !== 'LI') return el.outerHTML;
  const clone = el.cloneNode(true);
  clone.querySelectorAll(':scope > ul, :scope > ol').forEach((l) => l.remove());
  return clone.outerHTML;
}

// Remembers what each rendered paragraph looks like, right after rendering.
function snapshot(model, root) {
  model.signatures = new WeakMap();
  for (const el of root.querySelectorAll('[data-x]')) {
    const node = model.renderer.nodes[el.dataset.x];
    if (node?.localName === 'p') model.signatures.set(node, signature(el));
  }
}

const nodeOf = (model, el) => (el?.dataset?.x != null ? model.renderer.nodes[el.dataset.x] : null);

function register(model, el, node) {
  el.dataset.x = model.renderer.nodes.push(node) - 1;
  if (node.localName === 'p') model.signatures.set(node, signature(el));
}

// Writes the edited HTML (root, the rendered .doc element) back into the XML of the model.
function commit(model, root) {
  const used = new Set();
  commitBlocks(model, root, model.body, null, used);
  // Blocks that were rendered but are gone from the HTML were deleted.
  for (const node of model.renderer.blocksRendered) {
    if (!used.has(node) && model.body.contains(node)) node.remove();
  }
  model.renderer.blocksRendered = [...used];
  tidy(model);
  model.changedParts.add(model.main);
}

// The blocks of an HTML container in order: paragraphs (list items included) and tables.
// Loose text and inline elements, which editing can leave behind, get a paragraph of their own.
function* flatten(parent, depth = 0) {
  let loose = [];
  const wrapLoose = function* () {
    if (!loose.some((n) => n.nodeType === Node.ELEMENT_NODE || n.data.trim())) {
      loose = [];
      return;
    }
    const p = document.createElement('p');
    loose[0].before(p);
    p.append(...loose);
    loose = [];
    yield { el: p };
  };
  for (const n of [...parent.childNodes]) {
    if (n.nodeType !== Node.ELEMENT_NODE) {
      if (n.nodeType === Node.TEXT_NODE) loose.push(n);
      continue;
    }
    const tag = n.tagName;
    if (tag === 'UL' || tag === 'OL') {
      yield* wrapLoose();
      for (const li of n.children) {
        if (li.tagName !== 'LI') continue;
        yield { el: li, list: { ilvl: depth, ordered: tag === 'OL' } };
        for (const nested of li.children) if (nested.tagName === 'UL' || nested.tagName === 'OL') yield* flatten(wrapList(nested), depth + 1);
      }
    } else if (tag === 'TABLE') {
      yield* wrapLoose();
      yield { el: n, table: true };
    } else if (tag === 'DIV' && [...n.children].some((c) => isParagraphTag(c.tagName) || c.tagName === 'TABLE' || c.tagName === 'UL' || c.tagName === 'OL')) {
      yield* wrapLoose();
      yield* flatten(n, depth);
    } else if (isParagraphTag(tag)) {
      yield* wrapLoose();
      yield { el: n };
    } else if (tag === 'SECTION' || (tag === 'HR' && n.classList.contains('page-break'))) {
      yield* wrapLoose();
    } else {
      loose.push(n);
    }
  }
  yield* wrapLoose();
}

// flatten() walks containers; a nested list is walked as the only child of a stand-in.
function wrapList(list) {
  return { childNodes: [list] };
}

function commitBlocks(model, domParent, xmlParent, anchor, used) {
  let prev = anchor;
  for (const item of flatten(domParent)) {
    let node;
    if (item.table) {
      node = nodeOf(model, item.el);
      if (!node || node.localName !== 'tbl' || used.has(node)) continue;
      commitTable(model, item.el, node, used);
    } else {
      node = commitParagraph(model, item, used);
    }
    place(node, prev, xmlParent);
    used.add(node);
    prev = node;
  }
  return prev;
}

// Keeps node after prev in the XML. New nodes go right after prev, or first in parent.
function place(node, prev, parent) {
  if (node.parentNode) {
    if (!prev || prev.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING) return;
    prev.after(node);
    return;
  }
  if (prev) {
    prev.after(node);
    return;
  }
  const first = [...parent.children].find((c) => !c.localName.endsWith('Pr') || c.localName === 'sectPr');
  parent.insertBefore(node, first || null);
}

function commitTable(model, tableEl, tbl, used) {
  for (const tr of tableEl.rows) {
    if (!nodeOf(model, tr)) continue;
    for (const td of tr.cells) {
      const tc = nodeOf(model, td);
      if (tc?.localName === 'tc') commitBlocks(model, td, tc, null, used);
    }
  }
}

function commitParagraph(model, item, used) {
  const el = item.el;
  const mapped = nodeOf(model, el);
  const isP = mapped?.localName === 'p';
  // Editing can copy an element, data-x included (Enter splits a paragraph in two). The first
  // copy is the original; the others are new paragraphs that start from its properties.
  const original = isP && !used.has(mapped) && mapped.parentNode ? mapped : null;
  if (original && model.signatures.get(original) === signature(el)) return original;
  const follows = el.dataset.next != null ? model.renderer.nodes[el.dataset.next] : null;
  const template = isP ? mapped : follows || previousTemplate(model, el);
  const p = buildParagraph(model, el, template, original, item);
  // A paragraph made by Enter at the end of another gets the style that follows that one.
  if (follows && !isP) useNextStyle(model, p);
  delete el.dataset.next;
  if (original) original.replaceWith(p);
  register(model, el, p);
  return p;
}

function useNextStyle(model, p) {
  const pPr = child(p, 'pPr');
  if (!pPr) return;
  removeChildren(pPr, 'numPr', 'pageBreakBefore');
  const id = val(child(pPr, 'pStyle'));
  const styles = model.parts.styles?.doc ? [...model.parts.styles.doc.getElementsByTagNameNS(NS.w, 'style')] : [];
  const style = id && styles.find((s) => attr(s, 'styleId') === id);
  const next = style && val(child(style, 'next'));
  const isDefault = (sid) => styles.some((s) => attr(s, 'styleId') === sid && ['1', 'true'].includes(attr(s, 'default')));
  if (!style || !next || isDefault(next)) {
    // Without a next style Word keeps the style; headings and titles go on with normal text.
    if (!style || next || /^(heading \d|title|subtitle)$/.test((val(child(style, 'name')) || '').toLowerCase())) removeChildren(pPr, 'pStyle');
  } else putChild(pPr, W(model.doc, 'pStyle', { val: next }), PPR_ORDER);
  dropIfEmpty(pPr);
}

// A paragraph without a source takes its properties from the paragraph before it.
function previousTemplate(model, el) {
  for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling) {
    const node = nodeOf(model, prev);
    if (node?.localName === 'p') return node;
  }
  return null;
}

function buildParagraph(model, el, template, original, item) {
  const doc = model.doc;
  const p = W(doc, 'p');
  if (original) for (const a of original.attributes) p.setAttributeNS(a.namespaceURI, a.name, a.value);
  let pPr = child(template, 'pPr')?.cloneNode(true) || null;
  // A section break belongs to one paragraph only.
  if (pPr && !original) removeChildren(pPr, 'sectPr');

  const numPr = child(pPr, 'numPr');
  if (item.list) {
    pPr ||= W(doc, 'pPr');
    const numId = val(child(numPr, 'numId')) || siblingNumId(model, el) || listNumId(model, null, item.list.ordered ? 'number' : 'bullet');
    putChild(pPr, W(doc, 'numPr', {}, W(doc, 'ilvl', { val: item.list.ilvl }), W(doc, 'numId', { val: numId })), PPR_ORDER);
  } else if (numPr && !/^H[1-6]$/.test(el.tagName)) {
    // No longer in a list. Numbered headings keep their numbering.
    numPr.remove();
  }
  if (pPr?.children.length || pPr?.attributes.length) p.append(pPr);

  const markers = original ? keptMarkers(original) : { start: [], end: [] };
  p.append(...markers.start.map((n) => n.cloneNode(true)));
  p.append(...buildRuns(model, el));
  p.append(...markers.end.map((n) => n.cloneNode(true)));
  return p;
}

function siblingNumId(model, li) {
  for (const other of li.parentElement?.children || []) {
    const node = nodeOf(model, other);
    const id = node && val(child(child(child(node, 'pPr'), 'numPr'), 'numId'));
    if (id && id !== '0') return id;
  }
  return null;
}

// Parts of the original paragraph that its HTML does not show but that must survive a
// rewrite: bookmark and comment ranges (they can span paragraphs), comment marks, and the
// ends of fields that span paragraphs (like a table of contents).
function keptMarkers(p) {
  const start = [];
  const end = [];
  const all = [...p.getElementsByTagNameNS(NS.w, '*')].filter((n) => !insideTextbox(n));
  for (const n of all) {
    if (MARKER_STARTS.includes(n.localName)) start.push(n);
    else if (MARKER_ENDS.includes(n.localName)) end.push(n);
    else if (n.localName === 'r' && child(n, 'commentReference')) end.push(n);
  }
  // Fields: runs of fields still open at the end of the paragraph, and ends of fields opened before it.
  const open = [];
  for (const r of all.filter((n) => n.localName === 'r')) {
    const fc = child(r, 'fldChar');
    const type = attr(fc, 'fldCharType');
    if (type === 'begin') open.push({ runs: [r], code: true });
    else if (type === 'separate' && open.length) {
      open[open.length - 1].runs.push(r);
      open[open.length - 1].code = false;
    } else if (type === 'end') {
      if (open.length) open.pop();
      else end.push(r);
    } else if (open.length && open[open.length - 1].code && child(r, 'instrText')) open[open.length - 1].runs.push(r);
  }
  start.push(...open.flatMap((f) => f.runs));
  return { start, end };
}

// Runs of a paragraph from its HTML.
function buildRuns(model, block) {
  const out = [];
  const keys = new WeakMap();
  const doc = model.doc;

  const pushRun = (ctx, content) => {
    const key = JSON.stringify([ctx.props, ctx.tpl ? model.renderer.nodes.indexOf(ctx.tpl) : -1, ctx.link]);
    const last = ctx.out[ctx.out.length - 1];
    if (last && keys.get(last) === key) {
      last.append(content);
      return;
    }
    const r = W(doc, 'r');
    const rPr = buildRPr(model, ctx);
    if (rPr) r.append(rPr);
    r.append(content);
    keys.set(r, key);
    ctx.out.push(r);
  };

  const text = (ctx, s) => {
    s = s.replace(/​/g, '');
    for (const part of s.split(/(\t|\n)/)) {
      if (!part) continue;
      if (part === '\t') pushRun(ctx, W(doc, 'tab'));
      else if (part === '\n') pushRun(ctx, W(doc, 'br'));
      else {
        const t = W(doc, 't', {}, part);
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        pushRun(ctx, t);
      }
    }
  };

  const walk = (node, ctx) => {
    for (const n of node.childNodes) {
      if (n.nodeType === Node.TEXT_NODE) {
        text(ctx, n.data);
        continue;
      }
      if (n.nodeType !== Node.ELEMENT_NODE) continue;
      const tag = n.tagName;
      if (tag === 'UL' || tag === 'OL' || n.matches('.heading-number, a.bookmark, .noteback')) continue;
      const source = nodeOf(model, n);
      // Images, note references, formulas, text boxes and fields are copied from the original.
      if (Array.isArray(source)) {
        ctx.out.push(...source.map((x) => x.cloneNode(true)));
        continue;
      }
      if (source && (source.localName === 'r' || source.localName === 'fldSimple' || source.namespaceURI === NS.m)) {
        ctx.out.push(source.cloneNode(true));
        continue;
      }
      if (tag === 'IMG' || n.matches('.image-missing, .textbox, .noteref, .math, .field')) continue;
      if (tag === 'BR') {
        if (!isLastInBlock(n, block)) pushRun(ctx, W(doc, 'br'));
        continue;
      }
      if (tag === 'HR') {
        pushRun(ctx, W(doc, 'br', { type: 'page' }));
        continue;
      }
      if (tag === 'A' && !n.dataset.field) {
        const link = hyperlinkFor(model, n, source);
        if (link) {
          const inner = { ...ctx, out: [], link: source?.localName === 'hyperlink' ? 'old' : 'new' };
          walk(n, inner);
          if (inner.out.length) {
            link.append(...inner.out);
            ctx.out.push(link);
          }
          continue;
        }
      }
      const tpl = n.dataset.r != null ? model.renderer.nodes[n.dataset.r] : ctx.tpl;
      walk(n, { ...ctx, props: domFormat(ctx.props, n), tpl });
    }
  };

  walk(block, { out, props: {}, tpl: null, link: false });
  return out;
}

function hyperlinkFor(model, a, source) {
  const doc = model.doc;
  if (source?.localName === 'hyperlink') {
    const link = W(doc, 'hyperlink');
    for (const at of source.attributes) link.setAttributeNS(at.namespaceURI, at.name, at.value);
    return link;
  }
  if (a.dataset.rid) {
    const link = W(doc, 'hyperlink');
    link.setAttributeNS(NS.r, 'r:id', a.dataset.rid);
    return link;
  }
  const href = a.getAttribute('href') || '';
  if (href.startsWith('#bm-')) return W(doc, 'hyperlink', { anchor: href.slice(4) });
  return null;
}

// True when nothing visible follows a line break in its paragraph: browsers keep such a
// <br> in empty paragraphs, it is no line break of the text.
function isLastInBlock(br, block) {
  const range = document.createRange();
  range.setStartAfter(br);
  const nested = block.tagName === 'LI' && [...block.children].find((c) => c.tagName === 'UL' || c.tagName === 'OL');
  if (nested) range.setEndBefore(nested);
  else range.setEnd(block, block.childNodes.length);
  if (range.toString().replace(/​/g, '')) return false;
  return !range.cloneContents().querySelector('img, br, hr, .noteref, .math, .textbox, .image-missing, .field');
}

// Run formatting an element adds, as { b, i, u, s, va, color, bg, size, font }. Only what the
// HTML inside the paragraph says; the paragraph itself carries the style's formatting.
function domFormat(props, el) {
  const p = { ...props };
  const tag = el.tagName;
  const st = el.style;
  if (tag === 'STRONG' || tag === 'B') p.b = true;
  if (tag === 'EM' || tag === 'I') p.i = true;
  if (tag === 'U') p.u = true;
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') p.s = true;
  if (tag === 'SUP') p.va = 'superscript';
  if (tag === 'SUB') p.va = 'subscript';
  if (st.fontWeight) p.b = st.fontWeight === 'bold' || st.fontWeight === 'bolder' || Number(st.fontWeight) >= 600;
  if (st.fontStyle) p.i = st.fontStyle === 'italic' || st.fontStyle === 'oblique';
  const deco = st.textDecorationLine || '';
  if (deco) {
    if (deco === 'none') p.u = p.s = false;
    if (deco.includes('underline')) p.u = true;
    if (deco.includes('line-through')) p.s = true;
  }
  const color = toHex(st.color);
  if (color) p.color = color;
  const bg = toHex(st.backgroundColor);
  if (bg) p.bg = bg;
  else if (st.backgroundColor === 'transparent') p.bg = '';
  const size = toPoints(st.fontSize);
  if (size) p.size = size;
  if (st.fontFamily) p.font = st.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
  return p;
}

function toHex(color) {
  if (!color) return '';
  const hex = color.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (hex) return (hex[1].length === 3 ? [...hex[1]].map((c) => c + c).join('') : hex[1]).toUpperCase();
  const rgb = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/);
  if (!rgb || rgb[4] === '0') return '';
  return rgb.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function toPoints(size) {
  const m = (size || '').match(/^([\d.]+)(pt|px)$/);
  if (!m) return 0;
  return m[2] === 'pt' ? Number(m[1]) : Number(m[1]) * 0.75;
}

const HIGHLIGHT_NAMES = Object.fromEntries(Object.entries(HIGHLIGHTS).map(([name, css]) => [css.slice(1).toUpperCase(), name]));

function buildRPr(model, ctx) {
  const doc = model.doc;
  const tpl = child(ctx.tpl, 'rPr');
  const rPr = tpl ? tpl.cloneNode(true) : W(doc, 'rPr');
  const kept = { u: child(tpl, 'u'), dstrike: child(tpl, 'dstrike') };
  removeChildren(rPr, ...EDITED_RPR);
  const p = ctx.props;
  const put = (name, attrs = {}) => putChild(rPr, W(doc, name, attrs), RPR_ORDER);
  const toggle = (name, value) => {
    if (value === undefined) return;
    put(name, value ? {} : { val: '0' });
    put(`${name}Cs`, value ? {} : { val: '0' });
  };
  if (p.font) put('rFonts', { ascii: p.font, hAnsi: p.font, cs: p.font, eastAsia: p.font });
  toggle('b', p.b);
  toggle('i', p.i);
  if (p.u !== undefined) {
    if (p.u && kept.u && val(kept.u) !== 'none') putChild(rPr, kept.u.cloneNode(true), RPR_ORDER);
    else put('u', { val: p.u ? 'single' : 'none' });
  }
  if (p.s !== undefined) {
    if (p.s && kept.dstrike) putChild(rPr, kept.dstrike.cloneNode(true), RPR_ORDER);
    else put('strike', p.s ? {} : { val: '0' });
  }
  if (p.va) put('vertAlign', { val: p.va });
  if (p.color) put('color', { val: p.color });
  if (p.bg) {
    if (HIGHLIGHT_NAMES[p.bg]) put('highlight', { val: HIGHLIGHT_NAMES[p.bg] });
    else put('shd', { val: 'clear', color: 'auto', fill: p.bg });
  }
  if (p.size) {
    put('sz', { val: Math.round(p.size * 2) });
    put('szCs', { val: Math.round(p.size * 2) });
  }
  // Text in a new link looks like a link.
  if ((ctx.link === 'new' || (ctx.link && !ctx.tpl)) && !child(rPr, 'rStyle')) put('rStyle', { val: ensureStyle(model, 'Hyperlink') });
  return rPr.children.length ? rPr : null;
}

// Keeps the XML valid for Word: every table cell ends with a paragraph, as does the body.
function tidy(model) {
  const doc = model.doc;
  for (const tc of model.body.getElementsByTagNameNS(NS.w, 'tc')) {
    const blocks = [...tc.children].filter((c) => c.localName === 'p' || c.localName === 'tbl' || c.localName === 'sdt');
    if (!blocks.length || blocks[blocks.length - 1].localName === 'tbl') tc.append(W(doc, 'p'));
  }
  const body = model.body;
  const sect = [...body.children].find((c) => c.localName === 'sectPr');
  const blocks = [...body.children].filter((c) => c.localName !== 'sectPr');
  if (!blocks.length || blocks[blocks.length - 1].localName === 'tbl') body.insertBefore(W(doc, 'p'), sect || null);
}

// ---------- Styles ----------

// Built-in styles by the name Word gives them. Documents store the id in their own language
// (Kop1, Überschrift1), so styles are looked up by name and only added when missing.
const BUILTIN_STYLES = {
  Normal: { id: 'Normal', type: 'paragraph', name: 'Normal', xml: '<w:qFormat/>' },
  Title: { id: 'Title', type: 'paragraph', name: 'Title', xml: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>' },
  Subtitle: { id: 'Subtitle', type: 'paragraph', name: 'Subtitle', xml: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:color w:val="595959"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>' },
  ListParagraph: { id: 'ListParagraph', type: 'paragraph', name: 'List Paragraph', xml: '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>' },
  Quote: { id: 'Quote', type: 'paragraph', name: 'Quote', xml: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>' },
  TableGrid: { id: 'TableGrid', type: 'table', name: 'Table Grid', xml: '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>' },
  Hyperlink: { id: 'Hyperlink', type: 'character', name: 'Hyperlink', xml: '<w:rPr><w:color w:val="467886"/><w:u w:val="single"/></w:rPr>' },
};
const HEADING_SIZES = [40, 32, 28, 24, 22, 22];
HEADING_SIZES.forEach((size, i) => {
  const n = i + 1;
  BUILTIN_STYLES[`Heading${n}`] = {
    id: `Heading${n}`,
    type: 'paragraph',
    name: `heading ${n}`,
    xml: `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${n === 1 ? 360 : 160}" w:after="80"/><w:outlineLvl w:val="${i}"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>${n >= 4 ? '<w:i/><w:iCs/>' : ''}<w:color w:val="0F4761"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr>`,
  };
});

function findStyle(model, name) {
  const doc = model.parts.styles?.doc;
  if (!doc) return null;
  for (const s of doc.getElementsByTagNameNS(NS.w, 'style')) {
    if ((val(child(s, 'name')) || '').toLowerCase() === name.toLowerCase()) return s;
  }
  return null;
}

// The id of built-in style key (a key of BUILTIN_STYLES) in this document, added when missing.
function ensureStyle(model, key) {
  const def = BUILTIN_STYLES[key];
  const found = findStyle(model, def.name);
  if (found) return attr(found, 'styleId');
  const part = ensurePart(model, 'styles', 'word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', 'styles');
  const doc = part.doc;
  let id = def.id;
  while ([...doc.getElementsByTagNameNS(NS.w, 'style')].some((s) => attr(s, 'styleId') === id)) id += 'X';
  let xml = def.xml;
  if (!model.parts.theme) xml = xml.replace(/w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/g, 'w:ascii="Aptos Display" w:hAnsi="Aptos Display" w:eastAsia="Aptos Display" w:cs="Aptos Display"');
  // Styles are based on Normal, whatever it is called here.
  if (xml.includes('w:val="Normal"')) {
    const normal = key === 'Normal' ? null : ensureStyle(model, 'Normal');
    xml = xml.replace(/w:val="Normal"/g, `w:val="${normal}"`);
  }
  const parsed = new DOMParser().parseFromString(
    `<w:style xmlns:w="${NS.w}" w:type="${def.type}" w:styleId="${id}"><w:name w:val="${def.name}"/>${xml}</w:style>`,
    'application/xml',
  );
  doc.documentElement.appendChild(doc.importNode(parsed.documentElement, true));
  model.changedParts.add(part.path);
  return id;
}

// Name of the style of paragraph p (lower case, as Word stores built-in names).
function styleNameOf(model, p) {
  const id = val(child(child(p, 'pPr'), 'pStyle'));
  const doc = model.parts.styles?.doc;
  if (!doc) return id ? id.toLowerCase() : 'normal';
  const styles = [...doc.getElementsByTagNameNS(NS.w, 'style')];
  const style = id ? styles.find((s) => attr(s, 'styleId') === id) : styles.find((s) => attr(s, 'type') === 'paragraph' && ['1', 'true'].includes(attr(s, 'default')));
  return (val(child(style, 'name')) || id || 'normal').toLowerCase();
}

// Adds a part to the package (styles.xml, numbering.xml) with its relationship and content type.
function ensurePart(model, type, path, contentType, root) {
  if (model.parts[type]) return model.parts[type];
  const doc = new DOMParser().parseFromString(`<w:${root} xmlns:w="${NS.w}" xmlns:r="${NS.r}"/>`, 'application/xml');
  model.parts[type] = { path, doc };
  model.changedParts.add(path);
  addRelationship(model, `${REL_TYPE}/${type}`, path.startsWith('word/') ? path.slice(5) : `/${path}`);
  const types = model.contentTypes;
  const override = types.createElementNS(CT_NS, 'Override');
  override.setAttribute('PartName', `/${path}`);
  override.setAttribute('ContentType', contentType);
  types.documentElement.appendChild(override);
  model.changedParts.add('[Content_Types].xml');
  return model.parts[type];
}

// Adds a relationship of the main document part. Returns its id.
function addRelationship(model, type, target, external = false) {
  if (!model.relsDoc) {
    model.relsDoc = new DOMParser().parseFromString(`<Relationships xmlns="${NS.rel}"/>`, 'application/xml');
  }
  const ids = new Set([...model.relsDoc.getElementsByTagNameNS(NS.rel, 'Relationship')].map((r) => r.getAttribute('Id')));
  let n = ids.size + 1;
  while (ids.has(`rId${n}`)) n++;
  const rel = model.relsDoc.createElementNS(NS.rel, 'Relationship');
  rel.setAttribute('Id', `rId${n}`);
  rel.setAttribute('Type', type);
  rel.setAttribute('Target', target);
  if (external) rel.setAttribute('TargetMode', 'External');
  model.relsDoc.documentElement.appendChild(rel);
  model.rels = parseRels(model.relsDoc, model.main);
  model.changedParts.add(relsPath(model.main));
  return `rId${n}`;
}

function addHyperlink(model, url) {
  return addRelationship(model, `${REL_TYPE}/hyperlink`, url, true);
}

// ---------- Paragraph commands ----------

function setParagraphStyle(model, paragraphs, key) {
  const id = ensureStyle(model, key);
  const isDefault = key === 'Normal';
  for (const p of paragraphs) {
    const pPr = ensurePPr(p);
    if (isDefault) removeChildren(pPr, 'pStyle');
    else putChild(pPr, W(model.doc, 'pStyle', { val: id }), PPR_ORDER);
    dropIfEmpty(pPr);
  }
}

function setAlignment(model, paragraphs, value) {
  for (const p of paragraphs) {
    const pPr = ensurePPr(p);
    if (value === 'left') removeChildren(pPr, 'jc');
    else putChild(pPr, W(model.doc, 'jc', { val: value }), PPR_ORDER);
    dropIfEmpty(pPr);
  }
}

function listKind(model, p) {
  const numPr = child(child(p, 'pPr'), 'numPr');
  const id = val(child(numPr, 'numId'));
  if (!id || id === '0') return null;
  const level = parseNumbering(model.parts.numbering?.doc).get(id)?.get(num(val(child(numPr, 'ilvl'))) || 0);
  if (!level) return null;
  return level.format === 'bullet' ? 'bullet' : 'number';
}

// Bullets or numbering on or off, like the buttons in Word.
function toggleList(model, paragraphs, kind) {
  const listStyle = findStyle(model, 'List Paragraph');
  const listStyleId = listStyle && attr(listStyle, 'styleId');
  if (paragraphs.every((p) => listKind(model, p) === kind)) {
    for (const p of paragraphs) {
      const pPr = ensurePPr(p);
      removeChildren(pPr, 'numPr');
      if (listStyleId && val(child(pPr, 'pStyle')) === listStyleId) removeChildren(pPr, 'pStyle');
      dropIfEmpty(pPr);
    }
    return;
  }
  const numId = listNumId(model, paragraphs[0], kind);
  for (const p of paragraphs) {
    const pPr = ensurePPr(p);
    const ilvl = val(child(child(pPr, 'numPr'), 'ilvl')) || '0';
    putChild(pPr, W(model.doc, 'numPr', {}, W(model.doc, 'ilvl', { val: ilvl }), W(model.doc, 'numId', { val: numId })), PPR_ORDER);
    // The list sets the indent, as in Word.
    removeChildren(pPr, 'ind');
    if (!child(pPr, 'pStyle')) putChild(pPr, W(model.doc, 'pStyle', { val: ensureStyle(model, 'ListParagraph') }), PPR_ORDER);
  }
}

// Number and bullet formats offered in the toolbar, as [label, numFmt, lvlText]; %n is the level.
const NUMBER_STYLES = [
  ['1. 2. 3.', 'decimal', '%n.'],
  ['1) 2) 3)', 'decimal', '%n)'],
  ['a. b. c.', 'lowerLetter', '%n.'],
  ['a) b) c)', 'lowerLetter', '%n)'],
  ['A. B. C.', 'upperLetter', '%n.'],
  ['i. ii. iii.', 'lowerRoman', '%n.'],
  ['I. II. III.', 'upperRoman', '%n.'],
];
const BULLET_STYLES = [
  ['\u2022', 'bullet', '\u2022'],
  ['\u25e6', 'bullet', '\u25e6'],
  ['\u25aa', 'bullet', '\u25aa'],
  ['\u2013', 'bullet', '\u2013'],
  ['\u27a2', 'bullet', '\u27a2'],
  ['\u2713', 'bullet', '\u2713'],
];

// Gives the level each paragraph is on a number or bullet format, as the numbering library in
// Word does: only that level of that list changes. Paragraphs not in such a list become one.
function setListFormat(model, paragraphs, kind, format, text) {
  if (!paragraphs.length) return;
  if (!paragraphs.every((p) => listKind(model, p) === kind)) {
    toggleList(model, paragraphs.filter((p) => listKind(model, p) !== kind), kind);
  }
  const doc = model.parts.numbering.doc;
  const done = new Set();
  for (const p of paragraphs) {
    const numPr = child(child(p, 'pPr'), 'numPr');
    const numId = val(child(numPr, 'numId'));
    const ilvl = num(val(child(numPr, 'ilvl'))) || 0;
    if (done.has(`${numId}:${ilvl}`)) continue;
    done.add(`${numId}:${ilvl}`);
    ensureLevel(model, numId, ilvl);
    const numEl = kids(doc.documentElement, 'num').find((n) => attr(n, 'numId') === numId);
    const abstract = kids(doc.documentElement, 'abstractNum').find((a) => attr(a, 'abstractNumId') === val(child(numEl, 'abstractNumId')));
    const override = kids(numEl, 'lvlOverride').find((o) => (num(attr(o, 'ilvl')) || 0) === ilvl);
    const base = child(override, 'lvl') || kids(abstract, 'lvl').find((l) => (num(attr(l, 'ilvl')) || 0) === ilvl);
    const lvl = base ? base.cloneNode(true) : doc.importNode(new DOMParser().parseFromString(levelXml(kind, ilvl), 'application/xml').documentElement, true);
    const order = ['start', 'numFmt', 'lvlRestart', 'pStyle', 'isLgl', 'suff', 'lvlText', 'lvlPicBulletId', 'legacy', 'lvlJc', 'pPr', 'rPr'];
    putChild(lvl, W(doc, 'numFmt', { val: format }), order);
    putChild(lvl, W(doc, 'lvlText', { val: text.replace('%n', `%${ilvl + 1}`) }), order);
    if (!child(lvl, 'start')) putChild(lvl, W(doc, 'start', { val: 1 }), order);
    // The Symbol font of Word's own bullets would turn a plain character into another glyph.
    removeChildren(lvl, 'rPr', 'lvlPicBulletId');
    const next = W(doc, 'lvlOverride', { ilvl });
    const start = child(override, 'startOverride');
    if (start) next.append(start.cloneNode(true));
    next.append(lvl);
    if (override) override.replaceWith(next);
    else numEl.append(next);
    model.changedParts.add(model.parts.numbering.path);
  }
}

// Indent: list items move a level, other paragraphs move by half an inch.
function indent(model, paragraphs, delta) {
  for (const p of paragraphs) {
    const pPr = ensurePPr(p);
    const numPr = child(pPr, 'numPr');
    if (numPr && listKind(model, p)) {
      const ilvl = Math.max(0, Math.min(8, (num(val(child(numPr, 'ilvl'))) || 0) + delta));
      ensureLevel(model, val(child(numPr, 'numId')), ilvl);
      putChild(numPr, W(model.doc, 'ilvl', { val: ilvl }), ['ilvl', 'numId']);
      continue;
    }
    const ind = child(pPr, 'ind');
    const left = Math.max(0, (num(attr(ind, 'left') ?? attr(ind, 'start')) || 0) + delta * 720);
    const next = ind ? ind.cloneNode(true) : W(model.doc, 'ind');
    next.removeAttributeNS(NS.w, 'start');
    if (left) next.setAttributeNS(NS.w, 'w:left', String(left));
    else next.removeAttributeNS(NS.w, 'left');
    if (next.attributes.length) putChild(pPr, next, PPR_ORDER);
    else removeChildren(pPr, 'ind');
    dropIfEmpty(pPr);
  }
}

const BULLETS = ['•', 'o', '▪'];
const NUMBER_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman'];

function levelXml(kind, i) {
  const format = kind === 'bullet' ? 'bullet' : NUMBER_FORMATS[i % 3];
  const text = kind === 'bullet' ? BULLETS[i % 3] : `%${i + 1}.`;
  return `<w:lvl xmlns:w="${NS.w}" w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
}

// Word defines nine levels for every list; other programs sometimes fewer. A list item moved to
// a level its list does not define would lose its bullet or number, so the level is added.
function ensureLevel(model, numId, ilvl) {
  const doc = model.parts.numbering?.doc;
  const numEl = doc && kids(doc.documentElement, 'num').find((n) => attr(n, 'numId') === numId);
  const abstractId = val(child(numEl, 'abstractNumId'));
  const abstract = kids(doc?.documentElement, 'abstractNum').find((a) => attr(a, 'abstractNumId') === abstractId);
  if (!abstract) return;
  const has = (i) => kids(abstract, 'lvl').some((l) => (num(attr(l, 'ilvl')) || 0) === i);
  if (has(ilvl)) return;
  const first = kids(abstract, 'lvl').find((l) => attr(l, 'ilvl') === '0');
  const kind = val(child(first, 'numFmt')) === 'bullet' ? 'bullet' : 'number';
  for (let i = 0; i <= ilvl; i++) {
    if (has(i)) continue;
    const lvl = doc.importNode(new DOMParser().parseFromString(levelXml(kind, i), 'application/xml').documentElement, true);
    const after = kids(abstract, 'lvl').find((l) => (num(attr(l, 'ilvl')) || 0) > i);
    abstract.insertBefore(lvl, after || null);
  }
  model.changedParts.add(model.parts.numbering.path);
}

// A numbering definition for a new list, continuing the list right before paragraph p if it
// has the same kind. Returns the numId.
function listNumId(model, p, kind) {
  let prev = p?.previousElementSibling;
  while (prev && prev.localName !== 'p') prev = prev.previousElementSibling;
  if (prev && listKind(model, prev) === kind) return val(child(child(child(prev, 'pPr'), 'numPr'), 'numId'));

  const part = ensurePart(model, 'numbering', 'word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', 'numbering');
  const doc = part.doc;
  const root = doc.documentElement;
  const abstracts = kids(root, 'abstractNum');
  let abstract = abstracts.find((a) => {
    const format = val(child(kids(a, 'lvl').find((l) => attr(l, 'ilvl') === '0'), 'numFmt'));
    return kind === 'bullet' ? format === 'bullet' : format === 'decimal';
  });
  if (!abstract) {
    const id = Math.max(-1, ...abstracts.map((a) => num(attr(a, 'abstractNumId')) || 0)) + 1;
    const levels = Array.from({ length: 9 }, (_, i) => {
      const format = kind === 'bullet' ? 'bullet' : NUMBER_FORMATS[i % 3];
      const text = kind === 'bullet' ? BULLETS[i % 3] : `%${i + 1}.`;
      return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    }).join('');
    const parsed = new DOMParser().parseFromString(`<w:abstractNum xmlns:w="${NS.w}" w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`, 'application/xml');
    abstract = doc.importNode(parsed.documentElement, true);
    // All abstract definitions come before the num elements that use them.
    root.insertBefore(abstract, kids(root, 'num')[0] || null);
  }
  const numId = Math.max(0, ...kids(root, 'num').map((n) => num(attr(n, 'numId')) || 0)) + 1;
  const numEl = W(doc, 'num', { numId }, W(doc, 'abstractNumId', { val: attr(abstract, 'abstractNumId') }));
  // A new numbered list starts at 1 again.
  if (kind === 'number') numEl.append(W(doc, 'lvlOverride', { ilvl: 0 }, W(doc, 'startOverride', { val: 1 })));
  const last = kids(root, 'num').pop() || kids(root, 'abstractNum').pop();
  last.after(numEl);
  model.changedParts.add(part.path);
  return String(numId);
}

// ---------- Tables ----------

// The width of the text area in twips, from the last section of the document.
function textWidth(model) {
  const sect = [...model.body.children].reverse().find((c) => c.localName === 'sectPr');
  const w = num(attr(child(sect, 'pgSz'), 'w')) || 11906;
  const margin = child(sect, 'pgMar');
  return w - (num(attr(margin, 'left')) ?? 1440) - (num(attr(margin, 'right')) ?? 1440);
}

function emptyCell(doc, width, template) {
  const tcPr = W(doc, 'tcPr', {}, W(doc, 'tcW', { w: Math.round(width), type: 'dxa' }));
  const p = W(doc, 'p');
  const pPr = child(template, 'pPr');
  if (pPr) p.append(pPr.cloneNode(true));
  return W(doc, 'tc', {}, tcPr, p);
}

function insertTable(model, after, rows, cols) {
  const doc = model.doc;
  const width = textWidth(model);
  const colWidth = Math.floor(width / cols);
  const tbl = W(
    doc,
    'tbl',
    {},
    W(doc, 'tblPr', {}, W(doc, 'tblStyle', { val: ensureStyle(model, 'TableGrid') }), W(doc, 'tblW', { w: 0, type: 'auto' }), W(doc, 'tblLook', { val: '04A0', firstRow: 1, lastRow: 0, firstColumn: 1, lastColumn: 0, noHBand: 0, noVBand: 1 })),
    W(doc, 'tblGrid', {}, ...Array.from({ length: cols }, () => W(doc, 'gridCol', { w: colWidth }))),
    ...Array.from({ length: rows }, () => W(doc, 'tr', {}, ...Array.from({ length: cols }, () => emptyCell(doc, colWidth)))),
  );
  after.after(tbl);
  tidy(model);
  return tbl;
}

// Rows of a table with, per cell, the first grid column it covers and how many.
function tableGrid(tbl) {
  return kids(tbl, 'tr').map((tr) => {
    let col = num(val(child(child(tr, 'trPr'), 'gridBefore'))) || 0;
    const cells = kids(tr, 'tc').map((tc) => {
      const span = num(val(child(child(tc, 'tcPr'), 'gridSpan'))) || 1;
      const cell = { tc, col, span };
      col += span;
      return cell;
    });
    return { tr, cells };
  });
}

function gridColumn(tc) {
  const row = tableGrid(tc.parentNode.parentNode).find((r) => r.tr === tc.parentNode);
  return row.cells.find((c) => c.tc === tc);
}

function insertRow(model, tr, below) {
  const doc = model.doc;
  const row = W(doc, 'tr');
  const trPr = child(tr, 'trPr')?.cloneNode(true);
  if (trPr) {
    removeChildren(trPr, 'tblHeader');
    if (trPr.children.length) row.append(trPr);
  }
  for (const tc of kids(tr, 'tc')) {
    const cell = W(doc, 'tc');
    const tcPr = child(tc, 'tcPr')?.cloneNode(true);
    if (tcPr) {
      removeChildren(tcPr, 'vMerge');
      cell.append(tcPr);
    }
    const p = W(doc, 'p');
    const pPr = child(kids(tc, 'p')[0], 'pPr');
    if (pPr) p.append(pPr.cloneNode(true));
    cell.append(p);
    row.append(cell);
  }
  if (below) tr.after(row);
  else tr.before(row);
  return row;
}

function deleteRow(model, tr) {
  const tbl = tr.parentNode;
  const grid = tableGrid(tbl);
  const index = grid.findIndex((r) => r.tr === tr);
  // A merge that starts in this row starts in the next row instead.
  const next = grid[index + 1];
  if (next) {
    for (const cell of grid[index].cells) {
      if (val(child(child(cell.tc, 'tcPr'), 'vMerge')) !== 'restart') continue;
      const below = next.cells.find((c) => c.col === cell.col);
      const merge = child(child(below?.tc, 'tcPr'), 'vMerge');
      if (merge) merge.setAttributeNS(NS.w, 'w:val', 'restart');
    }
  }
  tr.remove();
  if (!kids(tbl, 'tr').length) tbl.remove();
  tidy(model);
}

function insertColumn(model, tc, right) {
  const doc = model.doc;
  const tbl = tc.parentNode.parentNode;
  const here = gridColumn(tc);
  const at = right ? here.col + here.span : here.col;
  const split = right ? here.col + here.span - 1 : here.col;
  const gridEl = child(tbl, 'tblGrid');
  const gridCols = kids(gridEl, 'gridCol');
  const width = Math.max(1, Math.floor((num(attr(gridCols[split], 'w')) || 1440) / 2));
  gridCols[split]?.setAttributeNS(NS.w, 'w:w', String(width));
  const col = W(doc, 'gridCol', { w: width });
  gridEl.insertBefore(col, gridCols[at] || null);

  for (const row of tableGrid(tbl)) {
    const inside = row.cells.find((c) => c.col < at && at < c.col + c.span);
    if (inside) {
      const tcPr = child(inside.tc, 'tcPr') || inside.tc.insertBefore(W(doc, 'tcPr'), inside.tc.firstChild);
      putChild(tcPr, W(doc, 'gridSpan', { val: inside.span + 1 }), TCPR_ORDER);
      continue;
    }
    const neighbour = row.cells.find((c) => c.col + c.span - 1 === split || c.col === split);
    if (neighbour?.span === 1) {
      const tcW = child(child(neighbour.tc, 'tcPr'), 'tcW');
      if (tcW && attr(tcW, 'type') === 'dxa') tcW.setAttributeNS(NS.w, 'w:w', String(width));
    }
    const cell = emptyCell(doc, width, kids(neighbour?.tc, 'p')[0]);
    const before = row.cells.find((c) => c.col >= at);
    row.tr.insertBefore(cell, before?.tc || null);
  }
}

function deleteColumn(model, tc) {
  const tbl = tc.parentNode.parentNode;
  const { col } = gridColumn(tc);
  for (const row of tableGrid(tbl)) {
    const cell = row.cells.find((c) => c.col <= col && col < c.col + c.span);
    if (!cell) continue;
    if (cell.span > 1) putChild(child(cell.tc, 'tcPr'), W(model.doc, 'gridSpan', { val: cell.span - 1 }), TCPR_ORDER);
    else cell.tc.remove();
    if (!kids(row.tr, 'tc').length) row.tr.remove();
  }
  kids(child(tbl, 'tblGrid'), 'gridCol')[col]?.remove();
  if (!kids(tbl, 'tr').length || !kids(child(tbl, 'tblGrid'), 'gridCol').length) tbl.remove();
  tidy(model);
}

function deleteTable(model, tbl) {
  tbl.remove();
  tidy(model);
}

// ---------- New document ----------

// Paper sizes in twips. Letter where Word uses it by default, A4 elsewhere.
const PAPER = { a4: [11906, 16838], letter: [12240, 15840] };
const LETTER_REGIONS = ['US', 'CA', 'MX', 'PH', 'PR', 'CL', 'CO', 'VE', 'GT', 'PA'];

function defaultPaper(locale) {
  try {
    return LETTER_REGIONS.includes(new Intl.Locale(locale).maximize().region) ? 'letter' : 'a4';
  } catch {
    return 'a4';
  }
}

// An empty .docx with one paragraph, the standard font and margins of Word. Styles such as
// headings are added when they are first used.
function newDocx(paper = 'a4') {
  const [w, h] = PAPER[paper];
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  const W_NS = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const files = {
    '[Content_Types].xml':
      `${head}<Types xmlns="${CT_NS}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '</Types>',
    '_rels/.rels':
      `${head}<Relationships xmlns="${NS.rel}">` +
      `<Relationship Id="rId1" Type="${REL_TYPE}/officeDocument" Target="word/document.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '</Relationships>',
    'docProps/core.xml':
      `${head}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      '</cp:coreProperties>',
    'word/_rels/document.xml.rels':
      `${head}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${REL_TYPE}/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml':
      `${head}<w:document ${W_NS}><w:body><w:p/>` +
      `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>` +
      '</w:body></w:document>',
    'word/styles.xml':
      `${head}<w:styles ${W_NS}>` +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Aptos" w:cs="Aptos"/><w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '</w:styles>',
  };
  const encoder = new TextEncoder();
  return writeZip(
    Object.entries(files).map(([path, xml]) => {
      const data = encoder.encode(xml);
      return { path, method: 0, crc: crc32(data), originalSize: data.length, data };
    }),
  );
}

// ---------- Saving ----------

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

async function deflate(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function serializeXml(doc) {
  const xml = new XMLSerializer().serializeToString(doc);
  return xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${xml}`;
}

// The edited document as a .docx Blob. Unchanged parts are copied from the original byte for byte.
async function saveDocx(model) {
  const docs = new Map([
    [model.main, model.doc],
    [relsPath(model.main), model.relsDoc],
    ['[Content_Types].xml', model.contentTypes],
  ]);
  for (const part of Object.values(model.parts)) if (part) docs.set(part.path, part.doc);

  const encoder = new TextEncoder();
  const files = [];
  const paths = [...model.pkg.entries.keys()];
  for (const path of model.changedParts) if (!paths.includes(path)) paths.push(path);
  for (const path of paths) {
    if (model.changedParts.has(path) && docs.get(path)) {
      const raw = encoder.encode(serializeXml(docs.get(path)));
      files.push({ path, method: 8, crc: crc32(raw), originalSize: raw.length, data: await deflate(raw) });
    } else {
      files.push({ path, ...model.pkg.entries.get(path) });
    }
  }
  return new Blob([writeZip(files)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

function writeZip(files) {
  const encoder = new TextEncoder();
  const DOS_DATE = 0x21; // 1980-01-01
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.path);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true); // UTF-8 names
    local.setUint16(8, file.method, true);
    local.setUint16(12, DOS_DATE, true);
    local.setUint32(14, file.crc, true);
    local.setUint32(18, file.data.length, true);
    local.setUint32(22, file.originalSize, true);
    local.setUint16(26, name.length, true);
    locals.push(new Uint8Array(local.buffer), name, file.data);

    const central = new DataView(new ArrayBuffer(46));
    central.setUint32(0, 0x02014b50, true);
    central.setUint16(4, 20, true);
    central.setUint16(6, 20, true);
    central.setUint16(8, 0x0800, true);
    central.setUint16(10, file.method, true);
    central.setUint16(14, DOS_DATE, true);
    central.setUint32(16, file.crc, true);
    central.setUint32(20, file.data.length, true);
    central.setUint32(24, file.originalSize, true);
    central.setUint16(28, name.length, true);
    central.setUint32(42, offset, true);
    centrals.push(new Uint8Array(central.buffer), name);
    offset += 30 + name.length + file.data.length;
  }
  const centralSize = centrals.reduce((sum, part) => sum + part.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, centralSize, true);
  end.setUint32(16, offset, true);

  const parts = [...locals, ...centrals, new Uint8Array(end.buffer)];
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let pos = 0;
  for (const part of parts) {
    out.set(part, pos);
    pos += part.length;
  }
  return out;
}
