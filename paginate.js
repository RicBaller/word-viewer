// Splitting the rendered document into pages for page view.
// Word stores no page layout in a .docx, so pages are laid out here by measuring the rendered
// content: blocks fill a page until the text area is full, paragraphs and lists are split
// between lines, tables between rows. The source element stays untouched; exports use it.

const HEADING = /^H[1-6]$/;
const MIN_LINES = 2; // widow and orphan control: lines a split paragraph keeps on each page

// Lays out source (the .doc element from loadDocx) as pages inside host, which must be in the
// document and visible, so the content can be measured. Returns the pages container.
async function paginate(source, host, page) {
  await document.fonts?.ready;
  await Promise.all([...source.querySelectorAll('img')].map((img) => img.decode().catch(() => {})));

  const container = document.createElement('div');
  container.className = 'pages';
  host.replaceChildren(container);

  let body;
  const newPage = () => {
    const sheet = source.cloneNode(false);
    sheet.removeAttribute('id');
    sheet.classList.add('page');
    body = sheet.appendChild(document.createElement('div'));
    body.className = 'page-body';
    container.appendChild(sheet);
  };
  // Lowest y (in viewport pixels) that content on the current page may reach.
  const limit = () => body.getBoundingClientRect().top + page.height - page.top - page.bottom;

  newPage();
  const queue = [...source.children].map((el) => el.cloneNode(true));
  while (queue.length) {
    const block = queue.shift();
    if (isForcedBreak(block)) {
      if (body.children.length) newPage();
      continue;
    }
    // A page break inside a paragraph ends the paragraph on this page.
    const afterBreak = splitAtPageBreak(block);
    if (afterBreak !== null) queue.unshift(PAGE_BREAK, ...(afterBreak ? [afterBreak] : []));
    if (isEmptyFragment(block)) continue;

    body.appendChild(block);
    if (bottom(block) <= limit()) continue;

    const alone = body.children.length === 1;
    const rest = splitBlock(block, limit(), alone);
    if (rest === null) continue;
    if (rest === false) {
      // Nothing of the block fits. Alone on a page it stays there and the page grows.
      if (alone) continue;
      block.remove();
      queue.unshift(...takeTrailingHeadings(body), block);
    } else {
      queue.unshift(rest);
    }
    newPage();
  }
  return container;
}

const PAGE_BREAK = Object.assign(document.createElement('hr'), { className: 'page-break' });

function isForcedBreak(el) {
  return el.tagName === 'HR' && el.classList.contains('page-break');
}

// Parts left over by a split at a page break, with no text or images, are dropped.
function isEmptyFragment(el) {
  return el.dataset.fragment !== undefined && !el.textContent.trim() && !el.querySelector('img, .image-missing, .textbox');
}

const bottom = (el) => el.getBoundingClientRect().bottom;

// Headings at the end of a page go to the next page with the text that follows them.
function takeTrailingHeadings(body) {
  const moved = [];
  while (body.children.length > 1 && HEADING.test(body.lastElementChild.tagName)) moved.unshift(body.lastElementChild);
  moved.forEach((el) => el.remove());
  return moved;
}

// Copy of el without content, for the part that continues on the next page.
function continuation(el) {
  const rest = el.cloneNode(false);
  rest.removeAttribute('id');
  rest.dataset.fragment = '';
  return rest;
}

// Moves everything after the first page break inside block (outside lists and tables) into a
// new element. Returns that element, undefined when the break was at the end, or null without a break.
function splitAtPageBreak(block) {
  const hr = [...block.querySelectorAll('hr.page-break')].find((el) => !el.closest('li, td'));
  if (!hr) return null;
  const range = document.createRange();
  range.setStartAfter(hr);
  range.setEnd(block, block.childNodes.length);
  const rest = continuation(block);
  rest.append(range.extractContents());
  hr.remove();
  block.dataset.fragment = '';
  rest.style.marginTop = '0';
  rest.style.textIndent = '0';
  return rest.textContent.trim() || rest.querySelector('img, .image-missing, .textbox') ? rest : undefined;
}

// Splits block so the part left in place ends above limit. Returns the remainder for the next
// page, null when block fits, or false when nothing of it can stay. With force, at least a
// piece stays wherever possible (the block is alone on its page).
function splitBlock(block, limit, force) {
  if (bottom(block) <= limit) return null;
  const tag = block.tagName;
  if (HEADING.test(tag) || tag === 'IMG' || tag === 'HR') return false;
  if (tag === 'TABLE') return splitTable(block, limit, force);
  if (tag === 'UL' || tag === 'OL') return splitList(block, limit, force);
  if (tag === 'LI') return splitListItem(block, limit, force);
  if (tag === 'SECTION' || tag === 'DIV') return splitChildren(block, limit, force);
  return splitLines(block, limit, force);
}

// Containers (notes, text boxes): split between or inside their child blocks.
function splitChildren(el, limit, force) {
  const kids = [...el.children];
  const k = kids.findIndex((c) => bottom(c) > limit);
  if (k < 0) return null;
  const tail = splitBlock(kids[k], limit, force && k === 0);
  if (tail === false && k === 0) return false;
  const rest = continuation(el);
  if (tail) rest.append(tail);
  rest.append(...kids.slice(tail === false ? k : k + 1));
  return rest;
}

// Lists split between items, or inside an item. The remainder keeps the numbering going.
function splitList(list, limit, force) {
  const rest = splitChildren(list, limit, force);
  if (rest && list.tagName === 'OL') {
    const kept = [...list.children].length;
    const first = rest.firstElementChild;
    // A continued item uses up its own number without showing it again.
    const start = (list.start || 1) + kept - (first?.classList.contains('continued') ? 1 : 0);
    rest.start = start;
  }
  return rest;
}

function splitListItem(li, limit, force) {
  const nested = [...li.children].find((c) => (c.tagName === 'UL' || c.tagName === 'OL') && bottom(c) > limit && c.getBoundingClientRect().top < limit);
  let rest;
  if (nested) {
    // The nested list continues on the next page; if none of it fits, all of it moves.
    const tail = splitList(nested, limit, false);
    rest = continuation(li);
    if (tail) rest.append(tail);
    let next = tail === false ? nested : nested.nextSibling;
    while (next) {
      const move = next;
      next = next.nextSibling;
      rest.appendChild(move);
    }
  } else {
    rest = splitLines(li, limit, force);
  }
  if (rest) rest.classList.add('continued');
  return rest;
}

// Tables split between rows. Header rows repeat on the next page; rows joined by a merged
// cell stay together.
function splitTable(table, limit, force) {
  const rows = [...table.rows];
  const headers = rows.findIndex((r) => !r.classList.contains('header-row'));
  const headerCount = headers < 0 ? 0 : headers;
  let k = rows.findIndex((r) => bottom(r) > limit);
  if (k < 0) return null;
  const crosses = (at) => rows.slice(0, at).some((r, i) => [...r.cells].some((c) => i + c.rowSpan > at));
  while (k > headerCount && crosses(k)) k--;
  if (k <= headerCount) {
    if (!force) return false;
    k = headerCount + 1;
    while (k < rows.length && crosses(k)) k++;
    if (k >= rows.length) return null;
  }
  const rest = continuation(table);
  const cols = table.querySelector(':scope > colgroup');
  if (cols) rest.appendChild(cols.cloneNode(true));
  const tbody = rest.appendChild(document.createElement('tbody'));
  tbody.append(...rows.slice(0, headerCount).map((r) => r.cloneNode(true)), ...rows.slice(k));
  return rest;
}

// Splits a paragraph between two lines. Keeps MIN_LINES lines on both pages when it can.
function splitLines(el, limit, force) {
  const texts = [];
  let length = 0;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    if (!walker.currentNode.data.length) continue;
    texts.push({ node: walker.currentNode, start: length });
    length += walker.currentNode.data.length;
  }
  if (!length) return false;

  const locate = (i) => {
    let lo = 0;
    let hi = texts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (texts[mid].start <= i) lo = mid;
      else hi = mid - 1;
    }
    return { node: texts[lo].node, offset: i - texts[lo].start };
  };
  const range = document.createRange();
  const rectAt = (i) => {
    const { node, offset } = locate(i);
    range.setStart(node, offset);
    range.setEnd(node, offset + 1);
    return range.getBoundingClientRect();
  };
  // First character whose line reaches below y.
  const firstBelow = (y) => {
    let lo = 0;
    let hi = length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (rectAt(mid).bottom > y) hi = mid;
      else lo = mid + 1;
    }
    return lo;
  };

  const box = el.getBoundingClientRect();
  let at = firstBelow(limit);
  if (at >= length) return null;
  let line = rectAt(at);
  const lineHeight = line.height || 16;
  const linesAbove = (i) => (rectAt(i).top - box.top) / lineHeight;
  // Too few lines would move to the next page: move more of them.
  if ((box.bottom - line.top) / lineHeight < MIN_LINES - 0.5) {
    const earlier = firstBelow(box.bottom - MIN_LINES * lineHeight);
    if (earlier > 0 && earlier < at) at = earlier;
  }
  // Too few lines would stay: move the whole paragraph, unless it must stay.
  if (at === 0 || linesAbove(at) < MIN_LINES - 0.5) {
    if (!force) return false;
    at = firstBelow(limit);
    if (at === 0) return false;
  }

  const { node, offset } = locate(at);
  range.setStart(node, offset);
  range.setEnd(el, el.childNodes.length);
  const rest = continuation(el);
  rest.append(range.extractContents());
  // The first line indent and the space above belong to the start of the paragraph only.
  rest.style.textIndent = '0';
  rest.style.marginTop = '0';
  el.style.marginBottom = '0';
  return rest;
}

if (typeof module !== 'undefined') {
  module.exports = { paginate };
}
