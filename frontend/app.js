const fileInput = document.getElementById('file');
const choose = document.getElementById('choose');
const dropzone = document.getElementById('dropzone');
const status = document.getElementById('status');
const statusTitle = document.getElementById('statusTitle');
const statusText = document.getElementById('statusText');

choose.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => fileInput.files[0] && upload(fileInput.files[0]));

['dragenter','dragover'].forEach(event => dropzone.addEventListener(event, e => {
  e.preventDefault(); dropzone.classList.add('drag');
}));
['dragleave','drop'].forEach(event => dropzone.addEventListener(event, e => {
  e.preventDefault(); dropzone.classList.remove('drag');
}));
dropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) upload(file);
});

async function upload(file) {
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    alert('Please choose a PDF file.');
    return;
  }
  const data = new FormData();
  data.append('file', file);
  dropzone.classList.add('hidden');
  status.classList.remove('hidden');
  statusTitle.textContent = 'Converting…';
  statusText.textContent = 'Extracting text, analyzing layout and creating the Word document.';

  try {
    const response = await fetch('/api/convert', { method: 'POST', body: data });
    if (!response.ok) throw new Error(await response.text());
    const blob = await response.blob();
    const disposition = response.headers.get('content-disposition') || '';
    const match = disposition.match(/filename="?([^";]+)"?/i);
    const filename = match ? match[1] : file.name.replace(/\.pdf$/i, '.docx');
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
    statusTitle.textContent = 'Conversion complete';
    statusText.textContent = 'The DOCX download has started.';
  } catch (error) {
    statusTitle.textContent = 'Conversion failed';
    statusText.textContent = error.message || 'Unknown error';
    dropzone.classList.remove('hidden');
  }
}
