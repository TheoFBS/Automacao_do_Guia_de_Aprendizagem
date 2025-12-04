const form = document.getElementById('jsonForm');
const jsonPreview = document.getElementById('jsonPreview');
const jsonOutput = document.getElementById('jsonOutput');
const downloadBtn = document.getElementById('downloadBtn');
let jsonData = null;

form.addEventListener('submit', function (e) {
    e.preventDefault();

    const formData = new FormData(form);
    const data = {};

    for (let [key, value] of formData.entries()) {
        if (value.trim() !== '') {
            data[key] = value;
        }
    }

    data.dataSubmissao = new Date().toISOString();
    data.id = Date.now();

    jsonData = data;
    jsonOutput.textContent = JSON.stringify(data, null, 2);

    fetch('/processar',{
        method: 'POST',
        headers:{
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(jsonData),
    })
    .then(response => response.json())
    .then(data => {
        console.log('Resposta do servidor', data);
    })
    .catch(error => {
        console.error('Erro:', error);
    })

    jsonPreview.classList.add('show');

    jsonPreview.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
});

downloadBtn.addEventListener('click', function () {
    if (!jsonData) return;

    const jsonString = JSON.stringify(jsonData, null, 2);

    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = `formulario_${jsonData.id}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
});

form.addEventListener('reset', function () {
    jsonPreview.classList.remove('show');
    jsonData = null;
});
