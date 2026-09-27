// Release history, newest first. Each version is also the element id, so /changelog#v1.2.0 links to it.
// Change texts are keyed by language; languages without their own text show English.

const CHANGELOG = [
  {
    version: '1.1.0',
    date: '2026-09-27',
    changes: {
      en: [
        'Page view now splits the document into pages of the right size, like in Word.',
        'Paragraphs, lists and tables continue on the next page; table header rows repeat at the top.',
        'Headings no longer end up alone at the bottom of a page.',
        'The top bar shows how many pages the document has.',
        'Printing and saving as PDF give the same pages as on screen.',
        'On small screens the pages are scaled down to fit.',
        'Table columns have the widths from the document and stay the same on every page.',
        'The changelog has a shorter address: /changelog.',
      ],
      nl: [
        'De paginaweergave deelt het document nu op in pagina’s van het juiste formaat, zoals in Word.',
        'Alinea’s, lijsten en tabellen lopen door op de volgende pagina; kopregels van tabellen herhalen bovenaan.',
        'Koppen blijven niet meer alleen onderaan een pagina staan.',
        'De bovenbalk toont hoeveel pagina’s het document heeft.',
        'Afdrukken en opslaan als PDF geven dezelfde pagina’s als op het scherm.',
        'Op kleine schermen worden de pagina’s verkleind zodat ze passen.',
        'Tabelkolommen krijgen de breedtes uit het document en blijven op elke pagina gelijk.',
        'De changelog heeft een korter adres: /changelog.',
      ],
    },
  },
  {
    version: '1.0.0',
    date: '2026-09-26',
    changes: {
      en: [
        'Open Word documents (.docx) by drag and drop or file picker, without uploading them.',
        'Headings, lists, tables, images, links, footnotes and endnotes are shown as in the document.',
        'Page view with the page size and margins of the document, and a text view for reading.',
        'Contents panel to jump to any heading.',
        'Find words in the whole document, with every match highlighted.',
        'Print the document or save it as PDF.',
        'Save the document as a web page, Markdown or plain text, or copy it to the clipboard.',
        'Try it out with a sample document.',
        'Interface available in 17 languages.',
      ],
      nl: [
        'Word-documenten (.docx) openen via slepen of bestandskiezer, zonder ze te uploaden.',
        'Koppen, lijsten, tabellen, afbeeldingen, links, voetnoten en eindnoten worden getoond zoals in het document.',
        'Paginaweergave met het paginaformaat en de marges van het document, en een tekstweergave om te lezen.',
        'Inhoudsopgave om naar elke kop te springen.',
        'Woorden zoeken in het hele document, met alle treffers gemarkeerd.',
        'Het document afdrukken of opslaan als PDF.',
        'Het document bewaren als webpagina, Markdown of platte tekst, of kopiëren naar het klembord.',
        'Probeer het uit met een voorbeelddocument.',
        'Interface beschikbaar in 17 talen.',
      ],
    },
  },
];
function renderChangelog() {
  const list = document.getElementById('timeline');
  const dateFormat = new Intl.DateTimeFormat(lang, { dateStyle: 'long' });
  list.replaceChildren();

  CHANGELOG.forEach((release, i) => {
    const item = document.createElement('li');
    item.className = 'release';
    item.id = `v${release.version}`;

    const head = document.createElement('div');
    head.className = 'release-head';
    const version = document.createElement('a');
    version.className = 'release-version';
    version.href = `#${item.id}`;
    version.textContent = `v${release.version}`;
    head.appendChild(version);
    if (i === 0) {
      const latest = document.createElement('span');
      latest.className = 'release-latest';
      latest.textContent = t('latest');
      head.appendChild(latest);
    }
    const date = document.createElement('time');
    date.className = 'muted';
    date.dateTime = release.date;
    date.textContent = dateFormat.format(new Date(`${release.date}T00:00:00`));
    head.appendChild(date);

    const changes = document.createElement('ul');
    for (const text of release.changes[lang] || release.changes.en) {
      changes.appendChild(document.createElement('li')).textContent = text;
    }

    item.append(head, changes);
    list.appendChild(item);
  });
}

initSite(renderChangelog);
renderChangelog();
if (location.hash) document.getElementById(location.hash.slice(1))?.scrollIntoView();
