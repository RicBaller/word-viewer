// Page chrome shared by all pages: language picker and footer credits.

const AUTHOR = 'Rick Baller';

function renderFooter() {
  const credits = document.getElementById('credits');
  if (credits) credits.textContent = `${t('madeBy', { name: AUTHOR })} · © ${new Date().getFullYear()}`;
}

// Detects the language, fills #languageSelect and calls onChange after the user picks another language.
function initSite(onChange = () => {}) {
  const select = document.getElementById('languageSelect');
  for (const [code, name] of Object.entries(LANGUAGES)) {
    select.add(new Option(name, code));
  }
  select.addEventListener('change', () => {
    setLanguage(select.value, { save: true });
    renderFooter();
    onChange();
  });
  setLanguage(detectLanguage());
  select.value = lang;
  renderFooter();
}
