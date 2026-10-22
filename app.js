// CV Match: interacciones básicas de la interfaz.
// En esta etapa no se llama a ninguna API de IA.

const fileInput = document.querySelector('#cv-file');
const uploadZone = document.querySelector('#upload-zone');
const uploadTitle = document.querySelector('#upload-title');
const fileStatus = document.querySelector('#file-status');
const jobOffer = document.querySelector('#job-offer');
const charCount = document.querySelector('#char-count');
const adaptButton = document.querySelector('#adapt-button');
const formMessage = document.querySelector('#form-message');

const MAX_FILE_SIZE = 10 * 1024 * 1024;
const ALLOWED_EXTENSIONS = ['pdf', 'docx'];
let selectedFile = null;

function showMessage(message) {
  formMessage.textContent = message;
  formMessage.classList.remove('d-none');
}

function clearMessage() {
  formMessage.textContent = '';
  formMessage.classList.add('d-none');
}

function setFile(file) {
  clearMessage();
  if (!file) return;

  const extension = file.name.split('.').pop().toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(extension)) {
    selectedFile = null;
    fileInput.value = '';
    uploadTitle.textContent = 'Arrastra tu CV aquí';
    fileStatus.textContent = 'Formato no compatible. Selecciona un PDF o DOCX.';
    showMessage('El archivo debe estar en formato PDF o DOCX.');
    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    selectedFile = null;
    fileInput.value = '';
    uploadTitle.textContent = 'Arrastra tu CV aquí';
    fileStatus.textContent = 'El archivo supera el límite de 10 MB.';
    showMessage('El archivo es demasiado grande. El tamaño máximo es 10 MB.');
    return;
  }

  selectedFile = file;
  uploadTitle.textContent = file.name;
  fileStatus.textContent = `Archivo seleccionado · ${(file.size / (1024 * 1024)).toFixed(2)} MB`;
}

fileInput.addEventListener('change', (event) => setFile(event.target.files[0]));

['dragenter', 'dragover'].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.add('dragging');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.remove('dragging');
  });
});

uploadZone.addEventListener('drop', (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;
  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    fileInput.files = transfer.files;
  } catch (_) {
    // La validación funciona aunque el navegador no permita asignar el archivo al input.
  }
  setFile(file);
});

jobOffer.addEventListener('input', () => {
  const count = jobOffer.value.length;
  charCount.textContent = `${count.toLocaleString('es-CL')} ${count === 1 ? 'carácter' : 'caracteres'}`;
  clearMessage();
});

adaptButton.addEventListener('click', () => {
  clearMessage();

  if (!selectedFile) {
    showMessage('Primero selecciona tu currículum en PDF o DOCX.');
    fileInput.focus();
    return;
  }

  if (!jobOffer.value.trim()) {
    showMessage('Pega la oferta laboral para preparar tu postulación.');
    jobOffer.focus();
    return;
  }

  showMessage('¡Perfecto! La interfaz está lista. Conectaremos la IA en la siguiente etapa.');
  formMessage.classList.remove('alert-warning');
  formMessage.classList.add('alert-success');
});
