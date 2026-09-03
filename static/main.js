// ── Language ──────────────────────────────────────────────────────────────
let currentLang = 'ru';

function setLang(lang) {
  currentLang = lang;
  document.getElementById('btnRu').classList.toggle('active', lang === 'ru');
  document.getElementById('btnEn').classList.toggle('active', lang === 'en');
  document.documentElement.lang = lang;

  document.querySelectorAll('[data-ru][data-en]').forEach(el => {
    const val = el.dataset[lang];
    if (val !== undefined) el.textContent = val;
  });

  // Placeholder needs separate handling
  const elec = document.getElementById('electricity');
  elec.placeholder = elec.dataset[`placeholder${lang.charAt(0).toUpperCase() + lang.slice(1)}`] || '';
}

// ── File inputs ───────────────────────────────────────────────────────────
const dropZone  = document.getElementById('dropZone');
const pdfInput  = document.getElementById('pdfInput');
const xlsxZone  = document.getElementById('xlsxZone');
const xlsxInput = document.getElementById('xlsxInput');
let selectedFile = null;
let selectedXlsx = null;

pdfInput.addEventListener('change', e => { if (e.target.files[0]) setFile(e.target.files[0]); });
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('dragover');
  if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
});

xlsxInput.addEventListener('change', e => { if (e.target.files[0]) setXlsx(e.target.files[0]); });
xlsxZone.addEventListener('dragover', e => { e.preventDefault(); xlsxZone.classList.add('dragover'); });
xlsxZone.addEventListener('dragleave', () => xlsxZone.classList.remove('dragover'));
xlsxZone.addEventListener('drop', e => {
  e.preventDefault(); xlsxZone.classList.remove('dragover');
  if (e.dataTransfer.files[0]) setXlsx(e.dataTransfer.files[0]);
});

function setFile(f) {
  selectedFile = f;
  dropZone.classList.add('has-file');
  document.getElementById('dropText').textContent = '✓ ' + f.name;
}

function setXlsx(f) {
  selectedXlsx = f;
  xlsxZone.classList.add('has-file');
  document.getElementById('xlsxText').textContent = '✓ ' + f.name;
}

// ── Status ────────────────────────────────────────────────────────────────
function showStatus(type) {
  ['Loading', 'Error', 'Success'].forEach(t => {
    document.getElementById('status' + t).style.display = 'none';
  });
  if (type) {
    document.getElementById('status' + type).style.display =
      type === 'Loading' ? 'flex' : 'block';
  }
}

// ── Process ───────────────────────────────────────────────────────────────
const LOADING_STEPS = {
  ru: ['AI читает квитанцию...', 'Извлекаем данные...', 'Считаем корректировки...', 'Записываем в Excel...'],
  en: ['AI is reading the bill...', 'Extracting data...', 'Calculating corrections...', 'Writing to Excel...'],
};

async function processFile() {
  if (!selectedFile) {
    alert(currentLang === 'ru' ? 'Выберите PDF-файл' : 'Please select a PDF file');
    return;
  }
  const elec = document.getElementById('electricity').value;
  if (!elec) {
    alert(currentLang === 'ru' ? 'Введите сумму электроэнергии ИПУ' : 'Please enter the electricity amount');
    return;
  }

  const btn = document.getElementById('processBtn');
  btn.disabled = true;
  showStatus('Loading');

  const steps = LOADING_STEPS[currentLang];
  let si = 0;
  const lt = document.getElementById('loadingText');
  lt.textContent = steps[0];
  const interval = setInterval(() => { si = (si + 1) % steps.length; lt.textContent = steps[si]; }, 2000);

  const fd = new FormData();
  fd.append('pdf', selectedFile);
  fd.append('electricity', elec);
  if (selectedXlsx) fd.append('xlsx', selectedXlsx);

  try {
    const resp = await fetch('/process', { method: 'POST', body: fd });
    const data = await resp.json();
    clearInterval(interval);

    if (data.error) {
      showStatus('Error');
      document.getElementById('statusError').textContent = '❌ ' + data.error;
    } else {
      document.getElementById('resPeriod').textContent = data.period || '—';

      const fmt = v => v != null ? Math.round(v).toLocaleString('ru') + ' ₽' : '—';
      document.getElementById('resTenant').textContent   = fmt(data.tenant_total);
      document.getElementById('resLandlord').textContent = fmt(data.landlord_total);

      const fname = data.filename || 'Квитанции_updated.xlsx';
      const dlBtn = document.getElementById('downloadBtn');
      dlBtn.href = '/download?file=' + encodeURIComponent(fname);
      dlBtn.textContent = (currentLang === 'ru' ? '⬇ Скачать ' : '⬇ Download ') + fname;

      // Re-apply language to success block static strings
      setLang(currentLang);
      // But restore dynamic content overwritten by setLang
      document.getElementById('resPeriod').textContent = data.period || '—';
      document.getElementById('resTenant').textContent   = fmt(data.tenant_total);
      document.getElementById('resLandlord').textContent = fmt(data.landlord_total);
      dlBtn.textContent = (currentLang === 'ru' ? '⬇ Скачать ' : '⬇ Download ') + fname;

      showStatus('Success');
      lucide.createIcons();
    }
  } catch (e) {
    clearInterval(interval);
    showStatus('Error');
    document.getElementById('statusError').textContent = '❌ ' + (currentLang === 'ru' ? 'Ошибка соединения: ' : 'Connection error: ') + e.message;
  }
  btn.disabled = false;
}

// ── Reset ─────────────────────────────────────────────────────────────────
function resetForm() {
  selectedFile = null;
  selectedXlsx = null;
  dropZone.classList.remove('has-file');
  xlsxZone.classList.remove('has-file');
  setLang(currentLang); // restore translated placeholders
  document.getElementById('electricity').value = '';
  document.getElementById('processBtn').disabled = false;
  showStatus(null);
  pdfInput.value = '';
  xlsxInput.value = '';
}

// ── Event listeners ───────────────────────────────────────────────────────
document.getElementById('processBtn').addEventListener('click', processFile);
document.getElementById('resetBtn').addEventListener('click', resetForm);
