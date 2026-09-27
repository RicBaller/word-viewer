// Release history, newest first. Each version is also the element id, so /changelog#v1.2.0 links to it.
// Change texts are keyed by language; languages without their own text show English.

const CHANGELOG = [
  {
    version: '1.3.0',
    date: '2026-09-27',
    changes: {
      en: [
        'Start a new Word document from the start page; it opens ready for typing.',
        'New documents use A4 paper, or Letter where that is the standard, with the font and margins Word uses.',
        'Links between the pages also work when you open the app from your own computer.',
        'Switching to the Edit view while pages are still being laid out no longer hides the document.',
      ],
      nl: [
        'Begin een nieuw Word-document vanaf de startpagina; het opent meteen klaar om te typen.',
        'Nieuwe documenten krijgen A4-papier, of Letter waar dat de standaard is, met het lettertype en de marges van Word.',
        'Links tussen de pagina’s werken ook als je de app vanaf je eigen computer opent.',
        'Overschakelen naar Bewerken terwijl de pagina’s nog worden opgedeeld, laat het document niet meer verdwijnen.',
      ],
    },
  },
  {
    version: '1.2.0',
    date: '2026-09-27',
    changes: {
      en: [
        'Edit Word documents in the new Edit view and download the result as a .docx file.',
        'Everything you do not change stays in the file exactly as it was, including headers, comments and fields.',
        'Format text: bold, italic, underline, strikethrough, superscript, subscript, font, size, color and highlight.',
        'Choose a paragraph style such as Heading 1 to 4, Title or Quote, and align text left, center, right or justified.',
        'Make bulleted and numbered lists, and indent them with Tab and Shift+Tab.',
        'Insert tables, add or delete rows and columns, and move between cells with Tab.',
        'Add links and page breaks.',
        'Undo and redo, with the keyboard shortcuts you know from Word.',
        'The top bar shows when changes are not downloaded yet, and closing the document asks first.',
        'Download the document as Word (.docx) from the export menu or with Ctrl+S.',
      ],
      nl: [
        'Bewerk Word-documenten in de nieuwe weergave Bewerken en download het resultaat als .docx-bestand.',
        'Alles wat je niet wijzigt, blijft precies zoals het was in het bestand, ook kopteksten, opmerkingen en velden.',
        'Maak tekst op: vet, cursief, onderstrepen, doorhalen, superscript, subscript, lettertype, grootte, kleur en markering.',
        'Kies een alineastijl zoals Kop 1 tot 4, Titel of Citaat, en lijn tekst links, gecentreerd, rechts of uitgevuld uit.',
        'Maak lijsten met opsommingstekens of nummers en laat ze inspringen met Tab en Shift+Tab.',
        'Voeg tabellen in, voeg rijen en kolommen toe of verwijder ze, en ga met Tab van cel naar cel.',
        'Voeg koppelingen en pagina-eindes toe.',
        'Maak ongedaan en opnieuw, met de sneltoetsen die je van Word kent.',
        'De bovenbalk toont wanneer wijzigingen nog niet zijn gedownload, en sluiten vraagt eerst om bevestiging.',
        'Download het document als Word (.docx) via het exportmenu of met Ctrl+S.',
      ],
    },
  },
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
