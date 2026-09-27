# Word Viewer

Simple browser app to open and read Word documents (`.docx`). No build step, no dependencies.

Live at [word-viewer.app](https://word-viewer.app/).

## Usage

Run `npx wrangler dev` and open http://localhost:8787. That serves the files the same way as production, clean URLs included (`/changelog`). Any static web server works for the app itself, but links to `/changelog` then need `changelog.html`.

- Drag a `.docx` file onto the page (or pick one with the button), or open the sample document.
- The document shows headings, lists, tables (merged cells included), images, links, bookmarks, footnotes, endnotes and text boxes, with fonts, colors, alignment, indents and spacing. Headers, footers, comments and tracked deletions are not shown.
- `Page` view shows paper with the page size and margins of the document. Word stores no page layout in a .docx, so `paginate.js` lays out the pages by measuring the rendered content: paragraphs and list items split between lines (at least two lines on each page), tables between rows (header rows repeat, merged rows stay together), and headings move to the next page with the text that follows. Page breaks in the document start a new page. Page counts come close to Word but can differ, because fonts such as Calibri or Aptos are often replaced by another font with other widths. On narrow screens the pages are scaled down instead of laid out again. Printing prints these pages. `Text` view reflows the text for reading.
- `Contents` lists the headings; click one to jump to it.
- Find (`Ctrl+F` / `Cmd+F` while a document is open) highlights every match. `Enter` and `Shift+Enter` go to the next and previous match.
- `Print / PDF` opens the print dialog with the page size and margins of the document. The arrow next to it saves a web page (`.html`, images embedded), Markdown (`.md`) or plain text (`.txt`). Markdown and plain text can also be copied to the clipboard. Exporters live in `export.js`.
- Old `.doc` files and password-protected documents are not supported; the app says so.
- The interface language follows your browser settings. Pick another language in the top bar; the choice is remembered. Supported: English, Dutch, German, French, Spanish, Italian, Portuguese, Polish, Turkish, Russian, Ukrainian, Arabic (right-to-left), Hindi, Indonesian, Chinese (Simplified), Japanese and Korean. Translations live in `i18n.js`.
- The footer links to `/changelog` (`changelog.html`), the release history. Add new releases at the top of `CHANGELOG` in `changelog.js`.

## How it works

`docx.js` reads the zip with `DecompressionStream`, parses the XML parts with `DOMParser` and builds the document as plain HTML elements (`h1`–`h6`, `p`, `ul`/`ol`, `table`, `strong`, `em`, …). Images become blob URLs. The exporters in `export.js` work on that HTML. Nothing is sent to a server.

Needs a current browser: Chrome/Edge 103+, Firefox 113+, Safari 16.4+. Find highlighting uses the CSS Custom Highlight API.

## SEO and hosting

- `word-viewer.app` is the canonical domain: `canonical`, Open Graph URLs, `sitemap.xml` and `robots.txt` point there.
- Update `lastmod` in `sitemap.xml` when page content changes.
- The intro and FAQ on the start page are duplicated in the JSON-LD `FAQPage` in `index.html`. Keep both in sync.
- `404.html` uses absolute paths, so it works at any URL. The web server must serve it for missing pages (nginx: `error_page 404 /404.html;`).
- Pushing a tag runs `.github/workflows/purgeCache.yml`, which purges the Cloudflare cache. It needs the repository variable `CLOUDFLARE_ZONES` (one `site=zone_id` per line) and the secret `CLOUDFLARE_TOKEN`.
- Icons: `favicon.svg` is the source. The PNGs in `icons/`, `apple-touch-icon.png`, `favicon.ico` and `og-image.png` (1200×630) are rendered from it.
- `sample.docx` is the document behind "Open sample document".
