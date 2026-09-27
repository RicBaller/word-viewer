// Exporting the rendered document (the element from loadDocx) to other formats:
// a standalone web page, Markdown and plain text. Printing to PDF is done by the browser.

const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'TABLE', 'SECTION', 'DIV', 'ARTICLE', 'HR']);

// Children grouped into blocks: block elements on their own, runs of inline content together.
function blockParts(el) {
  const parts = [];
  let inline = [];
  const flush = () => {
    if (inline.length) parts.push(inline);
    inline = [];
  };
  for (const node of el.childNodes) {
    if (node.nodeType === Node.ELEMENT_NODE && (BLOCK_TAGS.has(node.tagName) || node.classList.contains('textbox'))) {
      flush();
      parts.push(node);
    } else inline.push(node);
  }
  flush();
  return parts;
}

const isPageBreak = (node) => node.tagName === 'HR' && node.classList.contains('page-break');

// ---------- Markdown ----------

function toMarkdown(root) {
  const escape = (s) => s.replace(/([\\`*_[\]<>|~])/g, '\\$1').replace(/\t/g, ' ').replace(/ /g, ' ');
  // Markers go around the text, not around spaces at its ends: "**bold** " instead of "**bold **".
  const mark = (m, s) => {
    const [, lead, body, trail] = s.match(/^(\s*)([\s\S]*?)(\s*)$/);
    return body ? `${lead}${m}${body}${m}${trail}` : s;
  };

  const inline = (nodes, br) =>
    nodes
      .map((node) => {
        if (node.nodeType === Node.TEXT_NODE) return escape(node.data.replace(/\u200b/g, ''));
        if (node.nodeType !== Node.ELEMENT_NODE) return '';
        const inner = () => inline([...node.childNodes], br);
        switch (node.tagName) {
          case 'STRONG': return mark('**', inner());
          case 'EM': return mark('*', inner());
          case 'S': return mark('~~', inner());
          case 'BR': return br;
          case 'IMG':
          case 'HR': return '';
          case 'SUP':
            // Note labels are the element ids (fn-2, en-1), so footnotes and endnotes stay apart.
            if (node.classList.contains('noteref')) return `[^${node.querySelector('a').getAttribute('href').slice(1)}]`;
            return `<sup>${inner()}</sup>`;
          case 'SUB': return `<sub>${inner()}</sub>`;
          case 'A': {
            const href = node.getAttribute('href');
            const text = inner();
            return href && !href.startsWith('#') && text.trim() ? `[${text}](${href.replace(/[()\s]/g, encodeURIComponent)})` : text;
          }
          default:
            if (node.classList.contains('textbox')) return blocks(node, '').replace(/\n+/g, ' ');
            return inner();
        }
      })
      .join('');

  // Text at the start of a line that Markdown would read as a heading, list or quote.
  const guard = (s) => s.replace(/^(#{1,6}\s|[-+*]\s|\d+[.)]\s|>)/, '\\$1');

  const list = (el, indent) => {
    let n = el.start || 1;
    return [...el.children]
      .map((li) => {
        const marker = el.tagName === 'OL' ? `${n++}. ` : '- ';
        const pad = ' '.repeat(indent.length + marker.length);
        const lines = blockParts(li).map((part) =>
          Array.isArray(part) ? inline(part, `  \n${pad}`).trim() : part.tagName === 'UL' || part.tagName === 'OL' ? list(part, pad) : blocks(part, pad),
        );
        return `${indent}${marker}${lines.filter(Boolean).join('\n')}`;
      })
      .join('\n');
  };

  const table = (el) => {
    const rows = [...el.rows].map((tr) =>
      [...tr.cells].flatMap((td) => {
        const text = blockParts(td)
          .map((part) => (Array.isArray(part) ? inline(part, '<br>') : blocks(part, '').replace(/\n+/g, '<br>')))
          .filter(Boolean)
          .join('<br>')
          .trim();
        return [text, ...Array(td.colSpan - 1).fill('')];
      }),
    );
    if (!rows.length) return '';
    const width = Math.max(...rows.map((r) => r.length));
    const line = (cells) => `| ${[...cells, ...Array(width - cells.length).fill('')].join(' | ')} |`;
    return [line(rows[0]), line(Array(width).fill('---')), ...rows.slice(1).map(line)].join('\n');
  };

  const block = (part, indent) => {
    if (Array.isArray(part)) return indent + guard(inline(part, `  \n${indent}`).trim());
    const tag = part.tagName;
    if (/^H[1-6]$/.test(tag)) {
      const text = inline([...part.childNodes], ' ').trim();
      return text && `${indent}${'#'.repeat(Number(tag[1]))} ${text}`;
    }
    if (tag === 'P') {
      const text = inline([...part.childNodes], `  \n${indent}`).trim();
      return text && indent + guard(text);
    }
    if (tag === 'UL' || tag === 'OL') return list(part, indent);
    if (tag === 'TABLE') return table(part);
    if (tag === 'HR') return isPageBreak(part) ? '' : `${indent}---`;
    if (tag === 'SECTION' && part.classList.contains('notes')) {
      return [...part.querySelectorAll('.notes-list > li')]
        .map((li) => {
          const clone = li.cloneNode(true);
          clone.querySelector('.noteback')?.remove();
          return `[^${li.id}]: ${blocks(clone, '    ').trim()}`;
        })
        .join('\n');
    }
    return blocks(part, indent);
  };

  const blocks = (el, indent) =>
    blockParts(el)
      .map((part) => block(part, indent))
      .filter(Boolean)
      .join('\n\n');

  return blocks(root, '') + '\n';
}

// ---------- Plain text ----------

// Paragraphs separated by blank lines, lists with markers, tables as aligned columns.
function toText(root) {
  const inline = (nodes) =>
    nodes
      .map((node) => {
        if (node.nodeType === Node.TEXT_NODE) return node.data.replace(/\u00a0/g, ' ').replace(/\u200b/g, '');
        if (node.nodeType !== Node.ELEMENT_NODE) return '';
        if (node.tagName === 'BR') return '\n';
        if (node.tagName === 'IMG' || node.tagName === 'HR') return '';
        if (node.classList.contains('noteref')) return `[${node.textContent}]`;
        if (node.classList.contains('textbox')) return blocks(node, '').replace(/\n+/g, ' ');
        return inline([...node.childNodes]);
      })
      .join('');

  const indentLines = (s, pad) => s.replace(/\n/g, `\n${pad}`);

  const list = (el, indent) => {
    let n = el.start || 1;
    return [...el.children]
      .map((li) => {
        const marker = el.tagName === 'OL' ? `${n++}. ` : '- ';
        const pad = ' '.repeat(indent.length + marker.length);
        const lines = blockParts(li).map((part) =>
          Array.isArray(part) ? indentLines(inline(part).trim(), pad) : part.tagName === 'UL' || part.tagName === 'OL' ? list(part, pad) : blocks(part, pad),
        );
        return `${indent}${marker}${lines.filter(Boolean).join('\n')}`;
      })
      .join('\n');
  };

  const table = (el, indent) => {
    const rows = [...el.rows].map((tr) =>
      [...tr.cells].flatMap((td) => [blocks(td, '').replace(/\s*\n\s*/g, ' '), ...Array(td.colSpan - 1).fill('')]),
    );
    const width = Math.max(0, ...rows.map((r) => r.length));
    const widths = Array.from({ length: width }, (_, c) => Math.max(1, ...rows.map((r) => displayWidth(r[c] || ''))));
    return rows
      .map((r) => indent + widths.map((w, c) => (r[c] || '') + ' '.repeat(w - displayWidth(r[c] || ''))).join('  ').trimEnd())
      .join('\n');
  };

  const block = (part, indent) => {
    if (Array.isArray(part)) return indent + indentLines(inline(part).trim(), indent);
    const tag = part.tagName;
    if (tag === 'P' || /^H[1-6]$/.test(tag)) {
      const text = inline([...part.childNodes]).trim();
      return text && indent + indentLines(text, indent);
    }
    if (tag === 'UL' || tag === 'OL') return list(part, indent);
    if (tag === 'TABLE') return table(part, indent);
    if (tag === 'HR') return '';
    if (tag === 'SECTION' && part.classList.contains('notes')) {
      return [...part.querySelectorAll('.notes-list > li')]
        .map((li, i) => {
          const clone = li.cloneNode(true);
          clone.querySelector('.noteback')?.remove();
          const label = part.classList.contains('endnotes') ? toRoman(i + 1) : String(i + 1);
          return `[${label}] ${blocks(clone, '').trim()}`;
        })
        .join('\n');
    }
    return blocks(part, indent);
  };

  const blocks = (el, indent) =>
    blockParts(el)
      .map((part) => block(part, indent))
      .filter(Boolean)
      .join('\n\n');

  return blocks(root, '').replace(/[ \t]+$/gm, '') + '\n';
}

function displayWidth(s) {
  let width = 0;
  for (const ch of s) {
    const code = ch.codePointAt(0);
    const wide =
      (code >= 0x1100 && code <= 0x115f) ||
      (code >= 0x2e80 && code <= 0xa4cf) ||
      (code >= 0xac00 && code <= 0xd7a3) ||
      (code >= 0xf900 && code <= 0xfaff) ||
      (code >= 0xfe30 && code <= 0xfe4f) ||
      (code >= 0xff00 && code <= 0xff60) ||
      (code >= 0xffe0 && code <= 0xffe6) ||
      (code >= 0x1f300 && code <= 0x1faff) ||
      (code >= 0x20000 && code <= 0x3fffd);
    width += wide ? 2 : 1;
  }
  return width;
}

// ---------- Web page ----------

const HTML_EXPORT_CSS = `
body { margin: 0; background: #f3f4f6; }
.doc { box-sizing: border-box; max-width: 816px; margin: 24px auto; padding: 72px; background: #fff; color: #000; line-height: 1.15; white-space: pre-wrap; tab-size: 4; overflow-wrap: break-word; }
.doc p, .doc h1, .doc h2, .doc h3, .doc h4, .doc h5, .doc h6, .doc ul, .doc ol { margin: 0; }
.doc ul, .doc ol { padding-inline-start: 2em; }
.doc img { max-width: 100%; height: auto; vertical-align: bottom; }
.doc table { border-collapse: collapse; margin: 4px 0 12px; white-space: normal; }
.doc td { border: 1px solid #999; padding: 4px 6px; vertical-align: top; }
.doc .textbox { display: block; margin: 8px 0; padding: 8px; border: 1px solid #999; }
.doc .page-break { border: 0; border-top: 1px dashed #bbb; margin: 24px 0; }
.doc .notes { margin-top: 24px; font-size: 0.85em; }
.doc .noteback { text-decoration: none; }
@media print { body { background: none; } .doc { margin: 0; padding: 0; max-width: none; } .doc .page-break { break-after: page; border: 0; margin: 0; } }
`;

function toBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

// A standalone HTML file: images are embedded as data URLs, so it works without the .docx.
function toHTML(root, title, images, lang) {
  const clone = root.cloneNode(true);
  const byUrl = new Map(images.map((i) => [i.url, i]));
  clone.querySelectorAll('img').forEach((img) => {
    const image = byUrl.get(img.getAttribute('src'));
    if (image) img.src = `data:${image.mime};base64,${toBase64(image.bytes)}`;
  });
  // Leave out what only the viewer and editor use.
  for (const el of [clone, ...clone.querySelectorAll('*')]) {
    for (const name of ['data-i18n', 'data-x', 'data-r', 'data-rid', 'data-field', 'contenteditable', 'spellcheck']) el.removeAttribute(name);
  }
  clone.classList.remove('editing');
  const walker = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) walker.currentNode.data = walker.currentNode.data.replace(/\u200b/g, '');
  const escape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return (
    `<!doctype html>\n<html lang="${lang}">\n<head>\n<meta charset="utf-8">\n` +
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
    `<title>${escape(title)}</title>\n<style>${HTML_EXPORT_CSS}</style>\n</head>\n<body>\n${clone.outerHTML}\n</body>\n</html>\n`
  );
}

if (typeof module !== 'undefined') {
  module.exports = { toMarkdown, toText, toHTML, displayWidth };
}
