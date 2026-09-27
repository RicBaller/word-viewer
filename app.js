const $ = (id) => document.getElementById(id);

const els = {
  dropzone: $('dropzone'),
  about: $('about'),
  fileInput: $('fileInput'),
  sampleBtn: $('sampleBtn'),
  dropStatus: $('dropStatus'),
  viewer: $('viewer'),
  docHost: $('docHost'),
  actions: $('actions'),
  fileInfo: $('fileInfo'),
  pageViewBtn: $('pageViewBtn'),
  textViewBtn: $('textViewBtn'),
  printBtn: $('printBtn'),
  exportMenuBtn: $('exportMenuBtn'),
  closeBtn: $('closeBtn'),
  contentsBtn: $('contentsBtn'),
  contents: $('contents'),
  contentsList: $('contentsList'),
  contentsEmpty: $('contentsEmpty'),
  findInput: $('findInput'),
  findCount: $('findCount'),
  findPrevBtn: $('findPrevBtn'),
  findNextBtn: $('findNextBtn'),
  changelogLink: $('changelogLink'),
};

// Open document: { name, element, page: size and margins in px, images: blob URLs, words }.
let doc = null;

// Active view: 'page' (sheets of paper) or 'text' (reflowed, for reading).
let view = 'page';
// Contents panel starts open on wide screens.
let showContents = window.matchMedia('(min-width: 1100px)').matches;

const SAMPLE_URL = 'sample.docx';
const WIDE = window.matchMedia('(min-width: 900px)');

function show(section) {
  els.dropzone.hidden = section !== 'dropzone';
  els.about.hidden = section !== 'dropzone';
  els.viewer.hidden = section !== 'viewer';
  els.actions.hidden = section !== 'viewer';
  // With a document open, the changelog opens in a new tab so the document stays.
  els.changelogLink.target = section === 'dropzone' ? '' : '_blank';
}

// ---------- Loading a file ----------

// Shows a message in the drop zone: a status while opening, or an error.
function setStatus(key, isError = false) {
  els.dropStatus.hidden = !key;
  els.dropStatus.classList.toggle('error', isError);
  if (key) els.dropStatus.dataset.i18n = key;
  else delete els.dropStatus.dataset.i18n;
  els.dropStatus.textContent = key ? t(key) : '';
}

async function openBuffer(name, getBuffer) {
  setStatus('opening');
  try {
    const loaded = await loadDocx(await getBuffer());
    setStatus(null);
    openDoc({ name, ...loaded });
  } catch (err) {
    console.error(err);
    setStatus(['invalidFile', 'legacyFile'].includes(err.message) ? err.message : 'readError', true);
  }
}

function readFile(file) {
  if (!file) return;
  openBuffer(file.name, () => file.arrayBuffer());
}

function openSample() {
  openBuffer(SAMPLE_URL, async () => {
    const response = await fetch(SAMPLE_URL);
    if (!response.ok) throw new Error('readError');
    return response.arrayBuffer();
  });
}

function openDoc(newDoc) {
  closeDoc();
  doc = newDoc;
  doc.words = countWords(doc.element.textContent);
  doc.element.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach((el, i) => {
    el.id ||= `heading-${i + 1}`;
  });
  for (const [key, value] of Object.entries(doc.page)) doc.element.style.setProperty(`--page-${key}`, `${value}px`);
  els.findInput.value = '';
  renderContents();
  show('viewer');
  render();
  window.scrollTo(0, 0);
}

function closeDoc() {
  if (!doc) return;
  for (const image of doc.images) URL.revokeObjectURL(image.url);
  els.docHost.replaceChildren();
  doc = null;
  clearFind();
}

function countWords(text) {
  if (window.Intl?.Segmenter) {
    let n = 0;
    for (const s of new Intl.Segmenter(lang, { granularity: 'word' }).segment(text)) if (s.isWordLike) n++;
    return n;
  }
  return text.split(/\s+/).filter(Boolean).length;
}

// Page size and margins of the document, used when printing.
const printStyle = document.createElement('style');
document.head.appendChild(printStyle);

// Shows the document in the active view. Page view is laid out once per document, the first
// time it is shown; it needs the document on screen to measure it.
async function showDocument() {
  if (view === 'page' && !doc.pages) {
    const current = doc;
    current.paging ??= paginate(current.element, els.docHost, current.page);
    const pages = await current.paging;
    if (doc !== current) return;
    doc.pages = pages;
  }
  const root = currentRoot();
  if (doc.shown !== root) {
    doc.shown = root;
    els.docHost.replaceChildren(root);
    // Printed pages match the pages on screen; the text view is paged by the browser.
    const { width, height, top, right, bottom, left } = doc.page;
    const margin = view === 'page' ? '0' : `${top}px ${right}px ${bottom}px ${left}px`;
    printStyle.textContent = `@page { size: ${width}px ${height}px; margin: ${margin}; }`;
    // Matches are ranges in the element on screen, so they are found again.
    runFind();
  }
  fitPages();
  updateInfo();
}

// The element on screen: the pages in page view, the document itself in text view.
function currentRoot() {
  return view === 'page' ? doc.pages : doc.element;
}

// Pages keep the paper size of the document; on narrow screens they are scaled down to fit.
function fitPages() {
  if (!doc?.pages) return;
  const scale = Math.min(1, els.docHost.clientWidth / doc.page.width);
  doc.pages.style.zoom = scale < 1 ? String(scale) : '';
}

new ResizeObserver(fitPages).observe(els.docHost);

function updateInfo() {
  const parts = [doc.name];
  if (view === 'page' && doc.pages) parts.push(t('pages', { n: doc.pages.children.length }));
  parts.push(t('words', { n: doc.words }));
  els.fileInfo.textContent = parts.join(' · ');
}

// ---------- Rendering ----------

function render() {
  closeMenu();
  const isText = view === 'text';
  els.pageViewBtn.classList.toggle('active', !isText);
  els.textViewBtn.classList.toggle('active', isText);
  els.contents.hidden = !showContents;
  els.contentsBtn.classList.toggle('active', showContents);
  els.contentsBtn.setAttribute('aria-expanded', String(showContents));
  updateInfo();
  updateFindCount();
  showDocument();
}

function renderContents() {
  const headings = [...doc.element.querySelectorAll('h1, h2, h3, h4, h5, h6')].filter((el) => el.textContent.trim());
  const top = Math.min(...headings.map((el) => Number(el.tagName[1])));
  els.contentsList.replaceChildren(
    ...headings.map((el) => {
      const li = document.createElement('li');
      li.style.setProperty('--level', Number(el.tagName[1]) - top);
      const a = document.createElement('a');
      a.href = `#${el.id}`;
      a.textContent = el.textContent.trim();
      a.dir = 'auto';
      li.appendChild(a);
      return li;
    }),
  );
  els.contentsEmpty.hidden = headings.length > 0;
}

els.contentsList.addEventListener('click', (e) => {
  const a = e.target.closest('a');
  if (!a) return;
  e.preventDefault();
  document.getElementById(a.hash.slice(1))?.scrollIntoView({ block: 'start' });
  // On small screens the panel covers the document, so close it after jumping.
  if (!WIDE.matches) {
    showContents = false;
    render();
  }
});

// Links inside the document to bookmarks and notes scroll there without changing the URL.
els.docHost.addEventListener('click', (e) => {
  const a = e.target.closest('a[href^="#"]');
  if (!a) return;
  e.preventDefault();
  document.getElementById(decodeURIComponent(a.hash.slice(1)))?.scrollIntoView({ block: 'center' });
});

function setView(v) {
  view = v;
  render();
}

els.pageViewBtn.addEventListener('click', () => setView('page'));
els.textViewBtn.addEventListener('click', () => setView('text'));
els.contentsBtn.addEventListener('click', () => {
  showContents = !showContents;
  render();
});

// ---------- Find ----------

// Matches are DOM ranges, drawn with the CSS Custom Highlight API so the document is not changed.
let matches = [];
let currentMatch = -1;
const MAX_MATCHES = 5000;
const canHighlight = !!(window.CSS?.highlights && window.Highlight);

function clearFind() {
  matches = [];
  currentMatch = -1;
  if (canHighlight) {
    CSS.highlights.delete('find');
    CSS.highlights.delete('find-current');
  }
}

function runFind() {
  clearFind();
  const query = els.findInput.value;
  const root = doc && currentRoot();
  if (root && query.trim()) {
    const nodes = [];
    let text = '';
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      nodes.push({ node: walker.currentNode, start: text.length });
      text += walker.currentNode.data;
    }
    // Finds the text node holding character offset i of the joined text.
    const locate = (i, end) => {
      let lo = 0;
      let hi = nodes.length - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (nodes[mid].start < i || (!end && nodes[mid].start === i)) lo = mid;
        else hi = mid - 1;
      }
      return { node: nodes[lo].node, offset: i - nodes[lo].start };
    };
    const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu');
    for (const m of text.matchAll(pattern)) {
      const from = locate(m.index, false);
      const to = locate(m.index + m[0].length, true);
      const range = document.createRange();
      range.setStart(from.node, from.offset);
      range.setEnd(to.node, to.offset);
      matches.push(range);
      if (matches.length >= MAX_MATCHES) break;
    }
    if (canHighlight && matches.length) CSS.highlights.set('find', new Highlight(...matches));
  }
  if (matches.length) goToMatch(0);
  else updateFindCount();
}

function goToMatch(i) {
  if (!matches.length) return;
  currentMatch = (i + matches.length) % matches.length;
  const range = matches[currentMatch];
  if (canHighlight) CSS.highlights.set('find-current', new Highlight(range));
  const rect = range.getBoundingClientRect();
  const toolsBottom = document.querySelector('.doc-tools').getBoundingClientRect().bottom;
  if (rect.top < toolsBottom + 8 || rect.bottom > window.innerHeight - 8) {
    window.scrollBy({ top: rect.top - window.innerHeight / 3 });
  }
  updateFindCount();
}

function updateFindCount() {
  const hasQuery = !!els.findInput.value.trim();
  els.findCount.textContent = !hasQuery ? '' : matches.length ? t('matchCount', { i: currentMatch + 1, n: matches.length }) : t('noMatches');
  els.findPrevBtn.disabled = els.findNextBtn.disabled = matches.length === 0;
}

els.findInput.addEventListener('input', runFind);
els.findInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    goToMatch(currentMatch + (e.shiftKey ? -1 : 1));
  }
  if (e.key === 'Escape') {
    els.findInput.value = '';
    runFind();
    els.findInput.blur();
  }
});
els.findPrevBtn.addEventListener('click', () => goToMatch(currentMatch - 1));
els.findNextBtn.addEventListener('click', () => goToMatch(currentMatch + 1));

// Ctrl+F / Cmd+F searches the document while one is open.
document.addEventListener('keydown', (e) => {
  if (!doc || els.viewer.hidden || e.key.toLowerCase() !== 'f' || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
  e.preventDefault();
  els.findInput.focus();
  els.findInput.select();
});

// ---------- Menu ----------

const menu = document.createElement('div');
menu.className = 'menu';
menu.setAttribute('role', 'menu');
menu.hidden = true;
document.body.appendChild(menu);
let menuButton = null; // button the open menu belongs to
let menuActions = [];

// Items are [translation key, action] pairs, or [translation key, items] for a labelled group.
// anchor is the element the menu lines up with; defaults to the button itself.
function toggleMenu(btn, items, anchor = btn) {
  const wasOpen = menuButton === btn;
  closeMenu();
  if (wasOpen) return;
  menuButton = btn;
  menuActions = [];
  btn.setAttribute('aria-expanded', 'true');
  const build = ([key, action]) => {
    if (Array.isArray(action)) {
      const group = document.createElement('div');
      group.className = 'menu-group';
      group.setAttribute('role', 'group');
      const label = document.createElement('div');
      label.className = 'menu-group-label';
      label.id = `menu-group-${menuActions.length}`;
      label.textContent = t(key);
      group.setAttribute('aria-labelledby', label.id);
      group.append(label, ...action.map(build));
      return group;
    }
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'secondary';
    item.setAttribute('role', 'menuitem');
    item.textContent = t(key);
    item.dataset.index = menuActions.push(action) - 1;
    return item;
  };
  menu.replaceChildren(...items.map(build));
  menu.hidden = false;
  const rect = anchor.getBoundingClientRect();
  const rtl = document.documentElement.dir === 'rtl';
  const left = rtl ? rect.left : rect.right - menu.offsetWidth;
  const maxLeft = document.documentElement.clientWidth - menu.offsetWidth - 8;
  menu.style.top = `${rect.bottom + 4}px`;
  menu.style.left = `${Math.max(8, Math.min(left, maxLeft))}px`;
  menu.querySelector('button').focus();
}

function closeMenu() {
  if (!menuButton) return;
  menuButton.removeAttribute('aria-expanded');
  menuButton = null;
  menu.hidden = true;
}

menu.addEventListener('click', (e) => {
  const item = e.target.closest('button');
  if (!item) return;
  const action = menuActions[Number(item.dataset.index)];
  closeMenu();
  action();
});

document.addEventListener('pointerdown', (e) => {
  if (menuButton && !menu.contains(e.target) && !menuButton.contains(e.target)) closeMenu();
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && menuButton) {
    menuButton.focus();
    closeMenu();
  }
});
window.addEventListener('scroll', closeMenu, true);

// ---------- Export ----------

const baseName = () => doc.name.replace(/\.[^.]*$/, '');

function download(content, type, ext) {
  const blob = new Blob([content], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${baseName()}.${ext}`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// Short message at the bottom of the screen, e.g. after copying.
const toast = document.createElement('div');
toast.className = 'toast';
toast.setAttribute('role', 'status');
toast.hidden = true;
document.body.appendChild(toast);
let toastTimer = 0;

function showToast(message) {
  toast.textContent = message;
  toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (toast.hidden = true), 2500);
}

async function copy(text) {
  try {
    await navigator.clipboard.writeText(text);
    showToast(t('copied'));
  } catch {
    showToast(t('copyFailed'));
  }
}

const EXPORTS = {
  print: () => window.print(),
  html: () => download(toHTML(doc.element, baseName(), doc.images, lang), 'text/html;charset=utf-8', 'html'),
  md: () => download(toMarkdown(doc.element), 'text/markdown;charset=utf-8', 'md'),
  copyMd: () => copy(toMarkdown(doc.element)),
  txt: () => download(toText(doc.element), 'text/plain;charset=utf-8', 'txt'),
  copyTxt: () => copy(toText(doc.element)),
};

els.printBtn.addEventListener('click', EXPORTS.print);
els.exportMenuBtn.addEventListener('click', () => {
  toggleMenu(els.exportMenuBtn, [
    ['exportHtml', EXPORTS.html],
    ['exportMd', [['downloadFile', EXPORTS.md], ['copyClipboard', EXPORTS.copyMd]]],
    ['exportTxt', [['downloadFile', EXPORTS.txt], ['copyClipboard', EXPORTS.copyTxt]]],
  ], els.printBtn.parentElement);
});

// ---------- Wiring ----------

els.fileInput.addEventListener('change', () => {
  readFile(els.fileInput.files[0]);
  els.fileInput.value = '';
});

els.sampleBtn.addEventListener('click', openSample);

els.closeBtn.addEventListener('click', () => {
  closeDoc();
  setStatus(null);
  show('dropzone');
  window.scrollTo(0, 0);
});

// Drag & drop anywhere on the page.
let dragDepth = 0;
window.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dragDepth++;
  document.body.classList.add('dragging');
});
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) document.body.classList.remove('dragging');
});
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove('dragging');
  readFile(e.dataTransfer.files[0]);
});

// Installed as an app, .docx files can be opened with it from the file manager.
if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async ({ files }) => {
    if (files?.length) readFile(await files[0].getFile());
  });
}

// ---------- Language ----------

initSite(() => {
  if (doc) render();
});

show('dropzone');
