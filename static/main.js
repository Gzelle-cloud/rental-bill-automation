// DOM Elements
const dropZone = document.getElementById('dropZone');
const pdfInput = document.getElementById('pdfInput');
const xlsxZone = document.getElementById('xlsxZone');
const xlsxInput = document.getElementById('xlsxInput');
const processBtn = document.getElementById('processBtn');
const resetBtn = document.getElementById('resetBtn');
const electricityInput = document.getElementById('electricity');

// State
let selectedFile = null;
let selectedXlsx = null;

// Initialize event listeners
function initializeEventListeners() {
  // PDF file input
  pdfInput.addEventListener('change', handlePdfChange);
  
  // PDF drop zone
  dropZone.addEventListener('dragover', handleDragOver);
  dropZone.addEventListener('dragleave', handleDragLeave);
  dropZone.addEventListener('drop', handlePdfDrop);

  // XLSX file input
  xlsxInput.addEventListener('change', handleXlsxChange);
  
  // XLSX drop zone
  xlsxZone.addEventListener('dragover', handleDragOver);
  xlsxZone.addEventListener('dragleave', handleDragLeave);
  xlsxZone.addEventListener('drop', handleXlsxDrop);

  // Buttons
  processBtn.addEventListener('click', processFile);
  resetBtn.addEventListener('click', resetForm);
}

// Handle PDF file selection from input
function handlePdfChange(event) {
  if (event.target.files[0]) {
    setFile(event.target.files[0]);
  }
}

// Handle XLSX file selection from input
function handleXlsxChange(event) {
  if (event.target.files[0]) {
    setXlsx(event.target.files[0]);
  }
}

// Handle drag over for any drop zone
function handleDragOver(event) {
  event.preventDefault();
  event.currentTarget.classList.add('dragover');
}

// Handle drag leave for any drop zone
function handleDragLeave(event) {
  event.currentTarget.classList.remove('dragover');
}

// Handle PDF file drop
function handlePdfDrop(event) {
  event.preventDefault();
  dropZone.classList.remove('dragover');
  if (event.dataTransfer.files[0]) {
    setFile(event.dataTransfer.files[0]);
  }
}

// Handle XLSX file drop
function handleXlsxDrop(event) {
  event.preventDefault();
  xlsxZone.classList.remove('dragover');
  if (event.dataTransfer.files[0]) {
    setXlsx(event.dataTransfer.files[0]);
  }
}

// Set selected PDF file and update UI
function setFile(file) {
  selectedFile = file;
  dropZone.classList.add('has-file');
  document.getElementById('dropText').textContent = '✓ ' + file.name;
}

// Set selected XLSX file and update UI
function setXlsx(file) {
  selectedXlsx = file;
  xlsxZone.classList.add('has-file');
  document.getElementById('xlsxText').textContent = '✓ ' + file.name;
}

// Show status message (loading, success, error)
function showStatus(type) {
  ['Loading', 'Error', 'Success'].forEach(t => {
    document.getElementById('status' + t).style.display = 'none';
  });
  if (type) {
    document.getElementById('status' + type).style.display =
      type === 'Loading' ? 'flex' : 'block';
  }
}

// Create FormData from current state
function createFormData() {
  const fd = new FormData();
  fd.append('pdf', selectedFile);
  fd.append('electricity', electricityInput.value);
  if (selectedXlsx) {
    fd.append('xlsx', selectedXlsx);
  }
  return fd;
}

// Update loading progress text
function startLoadingAnimation() {
  const steps = [
    'AI читает квитанцию...',
    'Извлекаем данные об услугах...',
    'Рассчитываем корректировки...',
    'Записываем в Excel...',
  ];
  let stepIndex = 0;
  const loadingText = document.getElementById('loadingText');
  return setInterval(() => {
    stepIndex = (stepIndex + 1) % steps.length;
    loadingText.textContent = steps[stepIndex];
  }, 2000);
}

// Format number as Russian rubles rounded to whole rubles 
function formatRubles(value) {
  if (value == null) return '—';
  return Math.round(value).toLocaleString('ru') + ' ₽';
}

// Update result display with response data 
function updateResultDisplay(data) {
  document.getElementById('resPeriod').textContent = data.period || '—';
  document.getElementById('resTenant').textContent = formatRubles(data.tenant_total);
  document.getElementById('resLandlord').textContent = formatRubles(data.landlord_total);

  const filename = data.filename || 'Квитанции_updated.xlsx';
  const downloadBtn = document.getElementById('downloadBtn');
  downloadBtn.href = '/download?file=' + encodeURIComponent(filename);
  downloadBtn.textContent = '⬇ Скачать ' + filename;
}

// Show error message
function showError(message) {
  showStatus('Error');
  document.getElementById('statusError').textContent = '❌ ' + message;
}

// Main process file function 
async function processFile() {
  // Validate inputs
  if (!selectedFile) {
    alert('Выберите PDF-файл');
    return;
  }
  if (!electricityInput.value) {
    alert('Введите сумму электроэнергии ИПУ');
    return;
  }

  // Disable button and show loading
  processBtn.disabled = true;
  showStatus('Loading');
  const interval = startLoadingAnimation();

  try {
    // Send request
    const response = await fetch('/process', {
      method: 'POST',
      body: createFormData(),
    });
    const data = await response.json();

    // Handle response
    if (data.error) {
      showError(data.error);
    } else {
      updateResultDisplay(data);
      showStatus('Success');
    }
  } catch (error) {
    showError('Ошибка соединения: ' + error.message);
  } finally {
    clearInterval(interval);
    processBtn.disabled = false;
  }
}

// Reset form to initial state
function resetForm() {
  // Clear state
  selectedFile = null;
  selectedXlsx = null;

  // Reset UI
  dropZone.classList.remove('has-file');
  xlsxZone.classList.remove('has-file');
  document.getElementById('dropText').textContent =
    'Перетащите PDF сюда или нажмите для выбора';
  document.getElementById('xlsxText').textContent =
    'Загрузите актуальный файл Квитанции.xlsx';
  electricityInput.value = '';
  processBtn.disabled = false;

  // Clear inputs
  pdfInput.value = '';
  xlsxInput.value = '';

  // Hide status
  showStatus(null);
}

// Initialize on DOM ready
document.addEventListener('DOMContentLoaded', initializeEventListeners);
