// Reading .docx files: unzipping, parsing the WordprocessingML and rendering it as HTML.
// Uses only the browser: DecompressionStream for zip entries, DOMParser for the XML.
// Errors carry an i18n key as message: 'invalidFile' or 'legacyFile'.

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  v: 'urn:schemas-microsoft-com:vml',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

const TWIPS_PER_PX = 15; // 1440 twips per inch, 96 px per inch
const EMU_PER_PX = 9525; // 914400 EMU per inch

const IMAGE_TYPES = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};

const HIGHLIGHTS = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
  red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
  darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
  lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

const LIST_STYLES = {
  decimal: 'decimal', decimalZero: 'decimal-leading-zero', lowerLetter: 'lower-alpha',
  upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman',
};

// Word writes most bullets as characters of the Symbol and Wingdings fonts, in the Unicode
// private use area. Browsers do not have those glyphs, so they are shown as their look-alikes.
const SYMBOL_BULLETS = { '\uf0b7': '\u2022', '\uf0a7': '\u25aa', '\uf0d8': '\u27a2', '\uf076': '\u2756', '\uf0fc': '\u2713', '\uf0e8': '\u27a2', o: '\u25e6' };

function bulletMarker(text) {
  const c = SYMBOL_BULLETS[text] || text;
  return c && !/[\ue000-\uf8ff]/.test(c) ? `"${c.replace(/["\\]/g, '')}  "` : null;
}

const SERIF_FONTS = /times|cambria|georgia|garamond|palatino|book antiqua|baskerville|minion|constantia|serif/i;

// ---------- Zip ----------

// Reads the central directory. Returns a map of path to { method, crc, originalSize, data }
// (data still compressed, so unchanged entries can be written back as they are).
function readZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  if (bytes.length >= 4 && view.getUint32(0) === 0xd0cf11e0) throw new Error('legacyFile');
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 0xffff); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error('invalidFile');

  const decoder = new TextDecoder();
  const entries = new Map();
  const count = view.getUint16(end + 10, true);
  let pos = view.getUint32(end + 16, true);
  for (let n = 0; n < count; n++) {
    if (view.getUint32(pos, true) !== 0x02014b50) throw new Error('invalidFile');
    const method = view.getUint16(pos + 10, true);
    const crc = view.getUint32(pos + 16, true);
    const size = view.getUint32(pos + 20, true);
    const originalSize = view.getUint32(pos + 24, true);
    const nameLength = view.getUint16(pos + 28, true);
    const extraLength = view.getUint16(pos + 30, true);
    const commentLength = view.getUint16(pos + 32, true);
    const local = view.getUint32(pos + 42, true);
    const name = decoder.decode(bytes.subarray(pos + 46, pos + 46 + nameLength));
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    entries.set(name, { method, crc, originalSize, data: bytes.subarray(start, start + size) });
    pos += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

async function unzipEntry(entry) {
  if (entry.method === 0) return entry.data;
  if (entry.method !== 8) throw new Error('invalidFile');
  const stream = new Blob([entry.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// ---------- XML helpers ----------

const child = (el, name, ns = NS.w) => {
  if (!el) return null;
  for (const c of el.children) if (c.localName === name && c.namespaceURI === ns) return c;
  return null;
};
const kids = (el, name, ns = NS.w) => (el ? [...el.children].filter((c) => c.localName === name && c.namespaceURI === ns) : []);
const attr = (el, name, ns = NS.w) => (el ? el.getAttributeNS(ns, name) : null);
const val = (el) => attr(el, 'val');
const num = (value) => (value == null || value === '' ? undefined : Number(value));
// Toggle property such as <w:b/>: present means on, unless its value says otherwise.
const on = (el) => !['0', 'false', 'off', 'none'].includes(val(el));
const descendant = (el, name, ns) => el.getElementsByTagNameNS(ns, name)[0] || null;

// Resolves a relationship target against the folder of the part that owns it.
function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const piece of target.split('/')) {
    if (piece === '..') parts.pop();
    else if (piece !== '.') parts.push(piece);
  }
  return parts.join('/');
}

// ---------- Package ----------

class DocxPackage {
  constructor(entries) {
    this.entries = entries;
  }

  async bytes(path) {
    const entry = this.entries.get(path);
    return entry ? unzipEntry(entry) : null;
  }

  async xml(path) {
    const bytes = await this.bytes(path);
    if (!bytes) return null;
    const doc = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('invalidFile');
    return doc;
  }

  // Relationships of a part: id → { type, target (package path or external URL), external }.
  async rels(part) {
    return parseRels((await this.xml(relsPath(part))), part);
  }
}

const relsPath = (part) => {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};

function parseRels(doc, part) {
  const rels = new Map();
  if (!doc) return rels;
  for (const r of doc.getElementsByTagNameNS(NS.rel, 'Relationship')) {
    const external = r.getAttribute('TargetMode') === 'External';
    const target = r.getAttribute('Target') || '';
    rels.set(r.getAttribute('Id'), {
      type: (r.getAttribute('Type') || '').split('/').pop(),
      target: external ? target : resolvePath(part, target),
      external,
    });
  }
  return rels;
}

// ---------- Styles and numbering ----------

function parseStyles(doc) {
  const styles = new Map();
  let defaultParagraph = null;
  let defaultRPr = null;
  let defaultPPr = null;
  if (doc) {
    const defaults = descendant(doc, 'docDefaults', NS.w);
    defaultRPr = child(child(defaults, 'rPrDefault'), 'rPr');
    defaultPPr = child(child(defaults, 'pPrDefault'), 'pPr');
    for (const s of doc.getElementsByTagNameNS(NS.w, 'style')) {
      const id = attr(s, 'styleId');
      const style = {
        type: attr(s, 'type'),
        name: (val(child(s, 'name')) || '').toLowerCase(),
        basedOn: val(child(s, 'basedOn')),
        pPr: child(s, 'pPr'),
        rPr: child(s, 'rPr'),
      };
      styles.set(id, style);
      if (style.type === 'paragraph' && ['1', 'true'].includes(attr(s, 'default'))) defaultParagraph = id;
    }
  }
  return { styles, defaultParagraph, defaultRPr, defaultPPr };
}

// The style and the styles it is based on, base first.
function styleChain(styles, id) {
  const chain = [];
  const seen = new Set();
  while (id && styles.has(id) && !seen.has(id)) {
    seen.add(id);
    const style = styles.get(id);
    chain.unshift(style);
    id = style.basedOn;
  }
  return chain;
}

// Numbering definitions: numId → ilvl → { format, start, text, indent }.
function parseNumbering(doc) {
  const abstracts = new Map();
  const numbering = new Map();
  if (!doc) return numbering;
  const levels = (el) => {
    const map = new Map();
    for (const lvl of kids(el, 'lvl')) {
      map.set(num(attr(lvl, 'ilvl')) || 0, {
        format: val(child(lvl, 'numFmt')) || 'decimal',
        text: val(child(lvl, 'lvlText')) ?? '',
        // Where the text of this level starts, in twips from the margin.
        indent: num(attr(child(child(lvl, 'pPr'), 'ind'), 'left') ?? attr(child(child(lvl, 'pPr'), 'ind'), 'start')),
        start: num(val(child(lvl, 'start'))) ?? 1,
      });
    }
    return map;
  };
  for (const a of kids(doc.documentElement, 'abstractNum')) abstracts.set(attr(a, 'abstractNumId'), levels(a));
  for (const n of kids(doc.documentElement, 'num')) {
    const map = new Map(abstracts.get(val(child(n, 'abstractNumId'))) || []);
    for (const override of kids(n, 'lvlOverride')) {
      const ilvl = num(attr(override, 'ilvl')) || 0;
      const start = num(val(child(override, 'startOverride')));
      const own = levels(override).get(ilvl);
      const base = own || map.get(ilvl) || { format: 'decimal', start: 1 };
      map.set(ilvl, { ...base, start: start ?? base.start });
    }
    numbering.set(attr(n, 'numId'), map);
  }
  return numbering;
}

// Theme fonts, used when styles refer to "minor" (body) or "major" (heading) fonts.
function parseThemeFonts(doc) {
  const font = (name) => {
    const el = doc && descendant(doc, name, NS.a);
    return (el && child(el, 'latin', NS.a)?.getAttribute('typeface')) || '';
  };
  return { minor: font('minorFont'), major: font('majorFont') };
}

// ---------- Properties ----------

// Run properties from an rPr element on top of inherited ones.
function applyRPr(props, rPr, theme) {
  if (!rPr) return props;
  const p = { ...props };
  for (const el of rPr.children) {
    if (el.namespaceURI !== NS.w) continue;
    switch (el.localName) {
      case 'b': p.b = on(el); break;
      case 'i': p.i = on(el); break;
      case 'u': p.u = on(el); break;
      case 'strike':
      case 'dstrike': p.s = on(el); break;
      case 'caps': p.caps = on(el); break;
      case 'smallCaps': p.smallCaps = on(el); break;
      case 'vanish': p.hidden = on(el); break;
      case 'vertAlign': {
        const v = val(el);
        p.va = v === 'superscript' ? 'sup' : v === 'subscript' ? 'sub' : '';
        break;
      }
      case 'color': {
        const v = val(el);
        p.color = /^[0-9a-f]{6}$/i.test(v) ? `#${v.toLowerCase()}` : '';
        break;
      }
      case 'highlight': p.highlight = HIGHLIGHTS[val(el)] || ''; break;
      case 'shd': {
        const fill = attr(el, 'fill');
        if (/^[0-9a-f]{6}$/i.test(fill)) p.shade = `#${fill.toLowerCase()}`;
        break;
      }
      case 'sz': p.size = num(val(el)) / 2; break;
      case 'rFonts': {
        const themed = attr(el, 'asciiTheme') || attr(el, 'hAnsiTheme') || '';
        const font = attr(el, 'ascii') || attr(el, 'hAnsi') || (themed.startsWith('major') ? theme.major : themed ? theme.minor : '');
        if (font) p.font = font;
        break;
      }
    }
  }
  return p;
}

// Paragraph properties from a pPr element on top of inherited ones.
function applyPPr(props, pPr) {
  if (!pPr) return props;
  const p = { ...props };
  for (const el of pPr.children) {
    if (el.namespaceURI !== NS.w) continue;
    switch (el.localName) {
      case 'jc': p.align = { left: 'start', start: 'start', center: 'center', right: 'end', end: 'end', both: 'justify', distribute: 'justify' }[val(el)] ?? p.align; break;
      case 'ind': {
        p.left = num(attr(el, 'start') ?? attr(el, 'left')) ?? p.left;
        p.right = num(attr(el, 'end') ?? attr(el, 'right')) ?? p.right;
        const hanging = num(attr(el, 'hanging'));
        const first = num(attr(el, 'firstLine'));
        if (hanging !== undefined) p.first = -hanging;
        else if (first !== undefined) p.first = first;
        break;
      }
      case 'spacing': {
        p.before = num(attr(el, 'before')) ?? p.before;
        p.after = num(attr(el, 'after')) ?? p.after;
        const line = num(attr(el, 'line'));
        if (line !== undefined) p.line = (attr(el, 'lineRule') || 'auto') === 'auto' ? line / 240 : undefined;
        break;
      }
      case 'numPr': {
        const id = val(child(el, 'numId'));
        const ilvl = num(val(child(el, 'ilvl')));
        if (id !== null) p.numId = id;
        if (ilvl !== undefined) p.ilvl = ilvl;
        break;
      }
      case 'outlineLvl': p.outline = num(val(el)); break;
      case 'pageBreakBefore': p.pageBreakBefore = on(el); break;
      case 'contextualSpacing': p.contextual = on(el); break;
      case 'shd': {
        const fill = attr(el, 'fill');
        p.shade = /^[0-9a-f]{6}$/i.test(fill) ? `#${fill.toLowerCase()}` : '';
        break;
      }
      case 'bidi': p.rtl = on(el); break;
    }
  }
  return p;
}

const fontStack = (font) => {
  const name = font.replace(/["\\;{}]/g, '');
  return `"${name}", ${SERIF_FONTS.test(name) ? 'serif' : 'sans-serif'}`;
};

// Inline styles for run properties that differ from the surrounding block.
function runStyle(props, base) {
  const style = {};
  if (props.color !== base.color) style.color = props.color || '#000000';
  if (props.highlight || props.shade) style.background = props.highlight || props.shade;
  if (props.size && props.size !== base.size) style.fontSize = `${props.size}pt`;
  if (props.font && props.font !== base.font) style.fontFamily = fontStack(props.font);
  if (!!props.caps !== !!base.caps) style.textTransform = props.caps ? 'uppercase' : 'none';
  if (!!props.smallCaps !== !!base.smallCaps) style.fontVariant = props.smallCaps ? 'small-caps' : 'normal';
  if (base.b && !props.b) style.fontWeight = 'normal';
  if (base.i && !props.i) style.fontStyle = 'normal';
  return style;
}

// ---------- Rendering ----------

function h(tag, props = {}, ...content) {
  const el = document.createElement(tag);
  for (const [key, value] of Object.entries(props)) {
    if (key === 'style') Object.assign(el.style, value);
    else if (key === 'class') el.className = value;
    else el.setAttribute(key, value);
  }
  el.append(...content);
  return el;
}

// Only links that cannot run script: web, mail, phone and anchors within the document.
function safeHref(href) {
  return /^(https?:|mailto:|tel:|#)/i.test(href.trim()) ? href.trim() : null;
}

class Renderer {
  constructor({ styles, numbering, theme, images }) {
    this.styles = styles;
    this.numbering = numbering;
    this.theme = theme;
    this.images = images; // package path → { url, mime } or { unsupported }
    this.counters = new Map(); // numId → level counters
    this.fields = []; // open complex fields: { instr, link }
    this.notes = { footnote: [], endnote: [] };
    this.noteContent = { footnote: new Map(), endnote: new Map() };
    // For editing: rendered elements point to their XML node through data-x (blocks, links,
    // objects) or data-r (the run whose properties a stretch of text started from).
    this.nodes = [];
    this.blocksRendered = []; // editable w:p and w:tbl nodes of the body
    this.locked = 0; // > 0 inside text boxes and notes, which are not edited

    const { defaultRPr, defaultPPr, defaultParagraph } = styles;
    this.defaultParagraph = defaultParagraph;
    this.baseRun = applyRPr({}, defaultRPr, theme);
    this.basePara = applyPPr({}, defaultPPr);
    for (const s of styleChain(styles.styles, defaultParagraph)) {
      this.baseRun = applyRPr(this.baseRun, s.rPr, theme);
      this.basePara = applyPPr(this.basePara, s.pPr);
    }
  }

  mark(el, node, key = 'x') {
    if (this.locked) return;
    el.dataset[key] = this.nodes.push(node) - 1;
    if (key === 'x' && (node.localName === 'p' || node.localName === 'tbl')) this.blocksRendered.push(node);
  }

  // Renders body-level content (paragraphs, tables) into parent.
  blocks(parent, nodes, rels) {
    let lists = []; // open lists, outermost first: { el, ilvl, numId }
    for (const node of nodes) {
      if (node.namespaceURI !== NS.w) continue;
      switch (node.localName) {
        case 'p': {
          const para = this.paragraph(node, rels);
          this.mark(para.el, node);
          if (para.breakBefore) {
            lists = [];
            parent.appendChild(pageBreak());
          }
          if (para.list) lists = this.listItem(parent, lists, para);
          else {
            lists = [];
            parent.appendChild(para.el);
          }
          if (para.breakAfter) {
            lists = [];
            parent.appendChild(pageBreak());
          }
          break;
        }
        case 'tbl':
          lists = [];
          parent.appendChild(this.table(node, rels));
          break;
        case 'sdt':
          this.blocks(parent, child(node, 'sdtContent')?.children || [], rels);
          lists = [];
          break;
        case 'customXml':
        case 'ins':
        case 'moveTo':
          this.blocks(parent, node.children, rels);
          lists = [];
          break;
      }
    }
  }

  // Adds a list paragraph to the right list, opening and closing (nested) lists as needed.
  listItem(parent, lists, para) {
    const { numId, ilvl, level } = para.list;
    const ordered = level.format !== 'bullet' && level.format !== 'none';
    const tag = ordered ? 'OL' : 'UL';
    lists = lists.filter((l) => l.ilvl <= ilvl);
    let top = lists[lists.length - 1];
    if (top && top.ilvl === ilvl && (top.numId !== numId || top.el.tagName !== tag)) {
      lists.pop();
      top = lists[lists.length - 1];
    }

    const counters = this.count(numId, ilvl, level);

    if (!top || top.ilvl < ilvl) {
      const list = h(ordered ? 'ol' : 'ul');
      // The indent of the level, as in Word, measured from the list it is nested in. A jump of
      // several levels then indents several steps.
      if (level.indent !== undefined) list.style.paddingInlineStart = `${Math.max(0, level.indent - (top?.indent || 0)) / TWIPS_PER_PX}px`;
      if (ordered) {
        list.style.listStyleType = LIST_STYLES[level.format] || 'decimal';
        // Numbers followed by something else than a period, like 1) or (a).
        const m = level.text.match(/^([^%]*)%\d([^%]*)$/);
        if (m && (m[1] || m[2] !== '.')) {
          list.style.setProperty('--num-style', LIST_STYLES[level.format] || 'decimal');
          list.style.setProperty('--num-before', `"${m[1].replace(/["\\]/g, '')}"`);
          list.style.setProperty('--num-after', `"${m[2].replace(/["\\]/g, '')}"`);
          list.classList.add('custom-number');
        }
        if (counters[ilvl] !== 1) list.start = counters[ilvl];
      } else if (level.format === 'none') list.style.listStyleType = 'none';
      else {
        const marker = bulletMarker(level.text);
        if (marker) list.style.listStyleType = marker;
      }
      const host = top ? top.el.lastElementChild || top.el : parent;
      host.appendChild(list);
      top = { el: list, ilvl, numId, indent: level.indent ?? (lists[lists.length - 1]?.indent || 0) };
      lists.push(top);
    }
    top.el.appendChild(para.el);
    return lists;
  }

  // Advances the counter of a list level; deeper levels start over. Returns the counters.
  count(numId, ilvl, level) {
    const counters = this.counters.get(numId) || [];
    this.counters.set(numId, counters);
    counters[ilvl] = (counters[ilvl] ?? level.start - 1) + 1;
    counters.length = ilvl + 1;
    return counters;
  }

  // Number text of a numbered heading, such as "2.1", built from the level's lvlText ("%1.%2").
  numberLabel(numId, ilvl, level) {
    const counters = this.count(numId, ilvl, level);
    const levels = this.numbering.get(numId);
    return level.text.replace(/%(\d)/g, (m, d) => {
      const i = Number(d) - 1;
      const n = counters[i] ?? levels.get(i)?.start ?? 1;
      const format = levels.get(i)?.format;
      if (format === 'lowerLetter' || format === 'upperLetter') {
        const letter = String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
        return format === 'upperLetter' ? letter.toUpperCase() : letter;
      }
      if (format === 'lowerRoman') return toRoman(n);
      if (format === 'upperRoman') return toRoman(n).toUpperCase();
      return String(n);
    });
  }

  paragraph(p, rels) {
    const pPr = child(p, 'pPr');
    const styleId = val(child(pPr, 'pStyle')) || this.defaultParagraph;
    const chain = styleChain(this.styles.styles, styleId);
    let props = this.basePara;
    let run = this.baseRun;
    let tag = 'p';
    let title = false;
    // The default paragraph style is already part of the base properties.
    if (styleId !== this.defaultParagraph) {
      for (const s of chain) {
        props = applyPPr(props, s.pPr);
        run = applyRPr(run, s.rPr, this.theme);
      }
    }
    props = applyPPr(props, pPr);

    for (const s of [...chain].reverse()) {
      const heading = s.name.match(/^heading (\d)$/);
      if (heading) {
        tag = `h${Math.min(Number(heading[1]), 6)}`;
        break;
      }
      if (s.name === 'title') {
        tag = 'h1';
        title = true;
        break;
      }
    }
    if (tag === 'p' && props.outline !== undefined && props.outline < 6) tag = `h${props.outline + 1}`;

    let level = props.numId && props.numId !== '0' && this.numbering.get(props.numId)?.get(props.ilvl || 0);
    // Numbered headings stay headings, with their number in front instead of a list around them.
    let label = '';
    if (level && tag !== 'p') {
      label = this.numberLabel(props.numId, props.ilvl || 0, level);
      level = null;
    }
    const el = h(level ? 'li' : tag);
    if (title) el.className = 'doc-title';
    this.blockStyle(el, props, run, !!level);
    if (label) el.append(h('span', { class: 'heading-number', contenteditable: 'false' }, `${label} `));

    const ctx = { rels, base: run, last: null };
    this.inline(el, p.children, ctx, run);

    const text = el.textContent.trim();
    if (!text && !el.querySelector('img, .image-missing, .textbox')) {
      if (tag !== 'p' && !level) {
        const empty = h('p');
        this.blockStyle(empty, props, run, false);
        empty.appendChild(h('br'));
        return { el: empty, breakBefore: props.pageBreakBefore };
      }
      el.appendChild(h('br'));
    }

    const sect = child(pPr, 'sectPr');
    const sectType = val(child(sect, 'type'));
    return {
      el,
      list: level ? { numId: props.numId, ilvl: props.ilvl || 0, level } : null,
      breakBefore: props.pageBreakBefore,
      breakAfter: !!sect && sectType !== 'continuous',
    };
  }

  // Inline style of a paragraph in the default style, for paragraphs the editor adds.
  normalStyle() {
    const el = h('p');
    this.blockStyle(el, this.basePara, this.baseRun, false);
    return el.getAttribute('style') || '';
  }

  blockStyle(el, props, run, isList) {
    const style = el.style;
    if (props.align && props.align !== 'start') style.textAlign = props.align;
    if (!isList) {
      if (props.left) style.marginInlineStart = `${props.left / TWIPS_PER_PX}px`;
      if (props.right) style.marginInlineEnd = `${props.right / TWIPS_PER_PX}px`;
      if (props.first) style.textIndent = `${props.first / TWIPS_PER_PX}px`;
    }
    // List items usually share one style with contextual spacing: no space between them.
    if (!(isList && props.contextual)) {
      if (props.before !== undefined) style.marginTop = `${props.before / TWIPS_PER_PX}px`;
      if (props.after !== undefined) style.marginBottom = `${props.after / TWIPS_PER_PX}px`;
    }
    if (props.line) style.lineHeight = String(Math.round(props.line * 100) / 100);
    if (props.shade) style.background = props.shade;
    if (props.rtl) el.dir = 'rtl';
    Object.assign(style, runStyle(run, this.baseRun));
    if (run.b && !this.baseRun.b) style.fontWeight = 'bold';
    if (run.i && !this.baseRun.i) style.fontStyle = 'italic';
  }

  // Renders runs and other inline content of a paragraph into el.
  inline(el, nodes, ctx, base, inField = false) {
    nodes = [...nodes];
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i];
      // A field that starts and ends here (page number, cross-reference, date) is shown as its
      // result and kept as one piece, so editing the text around it does not break it.
      const end = this.locked || inField ? -1 : fieldEnd(nodes, i);
      if (end > i) {
        const field = h('span', { class: 'field', contenteditable: 'false' });
        const saved = this.fields.length;
        this.inline(field, nodes.slice(i, end + 1), { ...ctx, last: null }, base, true);
        this.fields.length = saved;
        this.mark(field, nodes.slice(i, end + 1));
        this.append(el, field, ctx);
        ctx.last = null;
        i = end;
        continue;
      }
      if (node.namespaceURI === NS.m && (node.localName === 'oMath' || node.localName === 'oMathPara')) {
        const text = [...node.getElementsByTagNameNS(NS.m, 't')].map((t) => t.textContent).join('');
        const math = h('span', { class: 'math', contenteditable: 'false' }, text);
        this.mark(math, node);
        this.append(el, math, ctx);
        continue;
      }
      if (node.namespaceURI !== NS.w) continue;
      switch (node.localName) {
        case 'r':
          this.run(el, node, ctx, base);
          break;
        case 'hyperlink': {
          const rel = ctx.rels.get(attr(node, 'id', NS.r));
          const anchor = attr(node, 'anchor');
          const href = safeHref(anchor ? `#bm-${anchor}` : rel?.external ? rel.target : '');
          if (href) {
            const a = h('a', { href });
            if (!href.startsWith('#')) Object.assign(a, { target: '_blank', rel: 'noopener' });
            this.mark(a, node);
            this.append(el, a, ctx);
            this.inline(a, node.children, { ...ctx, last: null }, base);
            ctx.last = null;
          } else this.inline(el, node.children, ctx, base);
          break;
        }
        case 'bookmarkStart': {
          const name = attr(node, 'name');
          if (name && name !== '_GoBack') {
            this.append(el, h('a', { id: `bm-${name}`, class: 'bookmark', contenteditable: 'false' }), ctx);
            ctx.last = null;
          }
          break;
        }
        case 'fldSimple': {
          if (this.locked) {
            this.inline(el, node.children, ctx, base);
            break;
          }
          const field = h('span', { class: 'field', contenteditable: 'false' });
          this.inline(field, node.children, { ...ctx, last: null }, base);
          this.mark(field, node);
          this.append(el, field, ctx);
          ctx.last = null;
          break;
        }
        case 'smartTag':
        case 'customXml':
        case 'ins':
        case 'moveTo':
          this.inline(el, node.children, ctx, base);
          break;
        case 'sdt':
          this.inline(el, child(node, 'sdtContent')?.children || [], ctx, base);
          break;
      }
    }
  }

  // Adds inline content, inside the link of an open HYPERLINK field when there is one.
  append(el, node, ctx) {
    const link = this.fields.findLast((f) => f.link)?.link;
    if (link && !el.closest('a')) {
      let a = el.lastChild;
      if (!(a instanceof HTMLAnchorElement) || a.dataset.field !== link) {
        a = h('a', { href: link, 'data-field': link });
        if (!link.startsWith('#')) Object.assign(a, { target: '_blank', rel: 'noopener' });
        el.appendChild(a);
      }
      a.appendChild(node);
      return;
    }
    el.appendChild(node);
  }

  run(el, r, ctx, base) {
    const rPr = child(r, 'rPr');
    let props = base;
    for (const s of styleChain(this.styles.styles, val(child(rPr, 'rStyle')))) props = applyRPr(props, s.rPr, this.theme);
    props = applyRPr(props, rPr, this.theme);
    if (props.hidden) return;

    const text = (s) => {
      if (this.fields.some((f) => f.instr !== null)) return; // field code, not its result
      const key = [props.b, props.i, props.u, props.s, props.va, props.color, props.highlight, props.shade, props.size, props.font, props.caps, props.smallCaps].join('|');
      const last = ctx.last;
      if (last && last.key === key && last.inner.isConnected && (last.outer === el.lastChild || last.outer === el.lastChild?.lastChild)) {
        last.inner.append(s);
        return;
      }
      const { outer, inner } = this.wrap(props, ctx);
      this.mark(outer, r, 'r');
      inner.append(s);
      this.append(el, outer, ctx);
      ctx.last = { key, outer, inner };
    };

    for (const node of r.children) {
      if (node.namespaceURI !== NS.w) {
        if (node.localName === 'AlternateContent') {
          const fallback = [...node.children].find((c) => c.localName === 'Fallback');
          const choice = [...node.children].find((c) => c.localName === 'Choice');
          for (const c of (choice || fallback)?.children || []) this.object(el, c, ctx, r);
        }
        continue;
      }
      switch (node.localName) {
        case 't':
          text(node.textContent);
          break;
        case 'tab':
        case 'ptab':
          text('\t');
          break;
        case 'noBreakHyphen':
          text('\u2011');
          break;
        case 'softHyphen':
          text('\u00ad');
          break;
        case 'sym': {
          const code = parseInt(attr(node, 'char'), 16);
          if (code) text(String.fromCodePoint(code >= 0xf000 ? code - 0xf000 : code));
          break;
        }
        case 'br':
        case 'cr':
          if (attr(node, 'type') === 'page') {
            this.append(el, h('hr', { class: 'page-break', contenteditable: 'false' }), ctx);
            ctx.last = null;
          } else text(h('br'));
          break;
        case 'instrText': {
          const field = this.fields[this.fields.length - 1];
          if (field && field.instr !== null) field.instr += node.textContent;
          break;
        }
        case 'fldChar': {
          const type = attr(node, 'fldCharType');
          if (type === 'begin') this.fields.push({ instr: '', link: null });
          else if (type === 'separate') {
            const field = this.fields[this.fields.length - 1];
            if (field) {
              const m = field.instr.match(/HYPERLINK\s+(\\l\s+)?"([^"]+)"(?:\s+\\l\s+"([^"]+)")?/);
              if (m) field.link = safeHref(m[1] ? `#bm-${m[2]}` : m[3] ? `${m[2]}#${m[3]}` : m[2]);
              field.instr = null;
            }
          } else if (type === 'end') this.fields.pop();
          ctx.last = null;
          break;
        }
        case 'footnoteReference':
        case 'endnoteReference': {
          const kind = node.localName === 'footnoteReference' ? 'footnote' : 'endnote';
          const id = attr(node, 'id');
          this.notes[kind].push(id);
          const n = this.notes[kind].length;
          const label = kind === 'footnote' ? String(n) : toRoman(n);
          const prefix = kind === 'footnote' ? 'fn' : 'en';
          const ref = h('sup', { class: 'noteref', contenteditable: 'false' }, h('a', { href: `#${prefix}-${id}`, id: `${prefix}ref-${id}` }, label));
          this.mark(ref, r);
          this.append(el, ref, ctx);
          ctx.last = null;
          break;
        }
        case 'drawing':
        case 'pict':
        case 'object':
          this.object(el, node, ctx, r);
          break;
      }
    }
  }

  // Images and text boxes inside run r.
  object(el, node, ctx, r) {
    const box = descendant(node, 'txbxContent', NS.w);
    if (box) {
      const div = h('span', { class: 'textbox', contenteditable: 'false' });
      this.locked++;
      this.blocks(div, box.children, ctx.rels);
      this.locked--;
      this.mark(div, r);
      this.append(el, div, ctx);
      ctx.last = null;
      return;
    }
    const blip = descendant(node, 'blip', NS.a);
    const imagedata = descendant(node, 'imagedata', NS.v);
    const id = blip ? attr(blip, 'embed', NS.r) : imagedata ? attr(imagedata, 'id', NS.r) : null;
    const rel = id && ctx.rels.get(id);
    if (!rel || rel.external) return;
    const image = this.images.get(rel.target);
    const docPr = descendant(node, 'docPr', NS.wp);
    const alt = docPr?.getAttribute('descr') || docPr?.getAttribute('title') || '';
    const extent = descendant(node, 'extent', NS.wp);
    let out;
    if (image?.url) {
      out = h('img', { src: image.url, alt });
      if (extent) {
        out.width = Math.round(num(extent.getAttribute('cx')) / EMU_PER_PX);
        out.height = Math.round(num(extent.getAttribute('cy')) / EMU_PER_PX);
      }
    } else {
      out = h('span', { class: 'image-missing', contenteditable: 'false', 'data-i18n': 'imageUnsupported', title: rel.target.split('/').pop() }, t('imageUnsupported'));
    }
    this.mark(out, r);
    this.append(el, out, ctx);
    ctx.last = null;
  }

  // Formatting elements for run properties: outer is added to the paragraph, text goes in inner.
  wrap(props, ctx) {
    const tags = [];
    if (props.b && !ctx.base.b) tags.push('strong');
    if (props.i && !ctx.base.i) tags.push('em');
    if (props.u && !ctx.base.u) tags.push('u');
    if (props.s && !ctx.base.s) tags.push('s');
    if (props.va) tags.push(props.va);
    const style = runStyle(props, ctx.base);
    if (Object.keys(style).length || !tags.length) tags.push('span');
    const outer = h(tags[0]);
    let inner = outer;
    for (const tag of tags.slice(1)) inner = inner.appendChild(h(tag));
    if (inner.tagName === 'SPAN') Object.assign(inner.style, style);
    return { outer, inner };
  }

  table(tbl, rels) {
    const table = h('table');
    this.mark(table, tbl);
    // Column widths from the table grid, as in Word. They also keep the columns the same
    // width when a table continues on the next page.
    const cols = kids(child(tbl, 'tblGrid'), 'gridCol').map((c) => num(attr(c, 'w')) || 0);
    if (cols.length && cols.every((w) => w > 0)) {
      table.appendChild(h('colgroup', {}, ...cols.map((w) => h('col', { style: { width: `${w / TWIPS_PER_PX}px` } }))));
      table.style.tableLayout = 'fixed';
      table.style.width = `${cols.reduce((a, b) => a + b, 0) / TWIPS_PER_PX}px`;
    }
    const tbody = table.appendChild(h('tbody'));
    const grid = []; // grid[row][column] = cell covering it, for vertical merges
    kids(tbl, 'tr').forEach((tr, r) => {
      const row = tbody.insertRow();
      this.mark(row, tr);
      // Header rows repeat at the top of each page the table continues on.
      const header = child(child(tr, 'trPr'), 'tblHeader');
      if (header && on(header)) row.className = 'header-row';
      grid[r] = [];
      let col = num(val(child(child(tr, 'trPr'), 'gridBefore'))) || 0;
      const cells = [...tr.children].flatMap((c) => (c.localName === 'sdt' ? kids(child(c, 'sdtContent'), 'tc') : c.localName === 'tc' ? [c] : []));
      for (const tc of cells) {
        const tcPr = child(tc, 'tcPr');
        const span = num(val(child(tcPr, 'gridSpan'))) || 1;
        const merge = child(tcPr, 'vMerge');
        const above = grid[r - 1]?.[col];
        if (merge && val(merge) !== 'restart' && above) {
          above.rowSpan += 1;
          for (let k = 0; k < span; k++) grid[r][col + k] = above;
          col += span;
          continue;
        }
        const td = row.insertCell();
        this.mark(td, tc);
        if (span > 1) td.colSpan = span;
        const fill = attr(child(tcPr, 'shd'), 'fill');
        if (/^[0-9a-f]{6}$/i.test(fill)) td.style.background = `#${fill.toLowerCase()}`;
        this.blocks(td, tc.children, rels);
        for (let k = 0; k < span; k++) grid[r][col + k] = td;
        col += span;
      }
    });
    return table;
  }

  // Footnotes or endnotes in the order they are referenced.
  notesSection(kind, rels) {
    const ids = this.notes[kind];
    if (!ids.length) return null;
    const prefix = kind === 'footnote' ? 'fn' : 'en';
    const list = h('ol', { class: 'notes-list' });
    if (kind === 'endnote') list.style.listStyleType = 'lower-roman';
    for (const id of ids) {
      const li = h('li', { id: `${prefix}-${id}` });
      const note = this.noteContent[kind].get(id);
      this.locked++;
      if (note) this.blocks(li, note.children, rels);
      this.locked--;
      // The back link goes at the end of the note's last paragraph, not on a line of its own.
      const last = li.lastElementChild?.tagName === 'P' ? li.lastElementChild : li;
      last.querySelector(':scope > br:last-child')?.remove();
      last.append(' ', h('a', { href: `#${prefix}ref-${id}`, class: 'noteback', 'aria-label': '↩' }, '↩'));
      list.appendChild(li);
    }
    return h('section', { class: `notes ${kind}s`, contenteditable: 'false' }, h('hr'), list);
  }
}

const pageBreak = () => h('hr', { class: 'page-break', contenteditable: 'false' });

// Index of the run that ends the complex field beginning at nodes[i], or -1 when nodes[i] does
// not begin a field or the field does not end among these siblings.
function fieldEnd(nodes, i) {
  const type = (r) => (r.localName === 'r' && r.namespaceURI === NS.w ? attr(child(r, 'fldChar'), 'fldCharType') : null);
  if (type(nodes[i]) !== 'begin') return -1;
  let depth = 0;
  for (let j = i; j < nodes.length; j++) {
    const t = type(nodes[j]);
    if (t === 'begin') depth++;
    else if (t === 'end' && --depth === 0) return j;
  }
  return -1;
}

function toRoman(n) {
  const numerals = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let out = '';
  for (const [value, s] of numerals) for (; n >= value; n -= value) out += s;
  return out;
}

// ---------- Loading ----------

// Opens a .docx from an ArrayBuffer. Returns { element, page, images, model }: element is the
// rendered document, page the page size and margins in px, images the blob URLs to revoke
// later, and model the parsed package that editing changes (see DocxModel).
async function loadDocx(buffer) {
  const pkg = new DocxPackage(readZip(buffer));
  const rootRels = await pkg.rels('');
  const main = [...rootRels.values()].find((r) => r.type === 'officeDocument')?.target || 'word/document.xml';
  const doc = await pkg.xml(main);
  if (!doc || !descendant(doc, 'body', NS.w)) throw new Error('invalidFile');

  const relsDoc = await pkg.xml(relsPath(main));
  const rels = parseRels(relsDoc, main);
  const partOf = (type) => [...rels.values()].find((r) => r.type === type && !r.external)?.target;
  const parts = {};
  for (const type of ['styles', 'numbering', 'theme']) {
    const path = partOf(type);
    parts[type] = path ? { path, doc: await pkg.xml(path) } : null;
  }

  const noteParts = {};
  for (const kind of ['footnote', 'endnote']) {
    const path = partOf(`${kind}s`);
    if (path) noteParts[kind] = { xml: await pkg.xml(path), rels: await pkg.rels(path) };
  }

  // Unpack all images up front, so rendering can stay synchronous.
  const images = new Map();
  const allRels = [rels, ...Object.values(noteParts).map((p) => p.rels)];
  for (const map of allRels) {
    for (const rel of map.values()) {
      if (rel.type !== 'image' || rel.external || images.has(rel.target)) continue;
      const mime = IMAGE_TYPES[rel.target.split('.').pop().toLowerCase()];
      const bytes = mime && (await pkg.bytes(rel.target));
      images.set(rel.target, bytes ? { url: URL.createObjectURL(new Blob([bytes], { type: mime })), mime, bytes } : { unsupported: true });
    }
  }

  const contentTypes = await pkg.xml('[Content_Types].xml');
  const model = new DocxModel({ buffer, pkg, main, doc, relsDoc, rels, parts, noteParts, images, contentTypes });
  const { element } = model.render();

  const sect = [...model.body.children].reverse().find((c) => c.localName === 'sectPr');
  const size = child(sect, 'pgSz');
  const margin = child(sect, 'pgMar');
  const px = (el, name, fallback) => (num(attr(el, name)) ?? fallback) / TWIPS_PER_PX;
  const page = {
    width: px(size, 'w', 11906),
    height: px(size, 'h', 16838),
    top: Math.abs(px(margin, 'top', 1440)),
    right: px(margin, 'right', 1440),
    bottom: Math.abs(px(margin, 'bottom', 1440)),
    left: px(margin, 'left', 1440),
  };

  return { element, page, images: [...images.values()].filter((i) => i.url), model };
}

// The parsed package: the XML of the main document and the parts around it. Editing changes
// this XML (docx-edit.js); render() turns it into HTML again.
class DocxModel {
  constructor(props) {
    Object.assign(this, props);
    this.changedParts = new Set(); // package paths whose XML must be written on save
  }

  get body() {
    return descendant(this.doc, 'body', NS.w);
  }

  render() {
    const renderer = new Renderer({
      styles: parseStyles(this.parts.styles?.doc),
      numbering: parseNumbering(this.parts.numbering?.doc),
      theme: parseThemeFonts(this.parts.theme?.doc),
      images: this.images,
    });
    for (const [kind, part] of Object.entries(this.noteParts)) {
      for (const note of part.xml?.getElementsByTagNameNS(NS.w, kind) || []) {
        if (!attr(note, 'type') || attr(note, 'type') === 'normal') renderer.noteContent[kind].set(attr(note, 'id'), note);
      }
    }

    // Documents run left to right unless a paragraph says otherwise, also inside a right-to-left interface.
    const article = h('article', { class: 'doc', dir: 'ltr' });
    const base = renderer.baseRun;
    if (base.font) article.style.fontFamily = fontStack(base.font);
    article.style.fontSize = `${base.size || 11}pt`;
    if (base.color) article.style.color = base.color;
    renderer.blocks(article, this.body.children, this.rels);
    for (const kind of ['footnote', 'endnote']) {
      const section = renderer.notesSection(kind, this.noteParts[kind]?.rels || new Map());
      if (section) article.appendChild(section);
    }
    this.renderer = renderer;
    return { element: article, renderer };
  }
}

if (typeof module !== 'undefined') {
  module.exports = { readZip, resolvePath, toRoman };
}
