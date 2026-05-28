// Configurar pdf.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// Elementos do DOM
const uploadBox = document.getElementById('uploadBox');
const fileInput = document.getElementById('fileInput');
const loadingSection = document.getElementById('loadingSection');
const successSection = document.getElementById('successSection');
const errorSection = document.getElementById('errorSection');
const downloadBtn = document.getElementById('downloadBtn');
const newFileBtn = document.getElementById('newFileBtn');
const retryBtn = document.getElementById('retryBtn');
const errorMessage = document.getElementById('errorMessage');
const loadingText = document.getElementById('loadingText');

let currentFile = null;
let docxContent = null;

// Event Listeners
uploadBox.addEventListener('click', () => fileInput.click());

uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadBox.classList.add('dragover');
});

uploadBox.addEventListener('dragleave', () => {
    uploadBox.classList.remove('dragover');
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadBox.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

downloadBtn.addEventListener('click', downloadWord);
newFileBtn.addEventListener('click', resetConverter);
retryBtn.addEventListener('click', resetConverter);

// Funções principais
function handleFile(file) {
    // Validar arquivo
    if (file.type !== 'application/pdf') {
        showError('Por favor, selecione um arquivo PDF válido.');
        return;
    }

    if (file.size > 100 * 1024 * 1024) { // 100MB
        showError('O arquivo é muito grande. Máximo 100MB.');
        return;
    }

    currentFile = file;
    convertPdfToWord(file);
}

async function convertPdfToWord(file) {
    try {
        showLoading();
        loadingText.textContent = 'Lendo PDF...';

        // Ler arquivo
        const arrayBuffer = await file.arrayBuffer();
        
        // Carregar PDF com pdf.js
        loadingText.textContent = 'Processando conteúdo...';
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        const paragraphs = [];

        // Extrair texto de cada página
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            loadingText.textContent = `Processando página ${pageNum} de ${pdf.numPages}...`;
            
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            
            // Agrupar texto por linhas
            const pageText = textContent.items
                .map(item => item.str)
                .join(' ');

            if (pageText.trim()) {
                paragraphs.push(pageText);
            }

            // Adicionar quebra de página
            if (pageNum < pdf.numPages) {
                paragraphs.push('\n--- Página ' + pageNum + ' ---\n');
            }
        }

        loadingText.textContent = 'Gerando documento Word...';

        // Criar documento Word com docx.js
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: paragraphs.map(text => 
                    new docx.Paragraph({
                        text: text,
                        spacing: {
                            line: 240,
                            lineRule: 'auto',
                        },
                    })
                ),
            }],
        });

        // Gerar arquivo DOCX
        docxContent = await docx.Packer.toBlob(doc);
        
        showSuccess();
    } catch (error) {
        console.error('Erro:', error);
        showError('Erro ao processar PDF: ' + error.message);
    }
}

function downloadWord() {
    if (!docxContent || !currentFile) {
        showError('Nenhum arquivo processado.');
        return;
    }

    // Gerar nome do arquivo
    const fileName = currentFile.name
        .replace('.pdf', '')
        .replace(/[^a-zA-Z0-9-_]/g, '_') + '.docx';

    // Criar link de download
    const url = URL.createObjectURL(docxContent);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function showLoading() {
    uploadBox.classList.add('hidden');
    successSection.classList.add('hidden');
    errorSection.classList.add('hidden');
    loadingSection.classList.remove('hidden');
}

function showSuccess() {
    uploadBox.classList.add('hidden');
    loadingSection.classList.add('hidden');
    errorSection.classList.add('hidden');
    successSection.classList.remove('hidden');
}

function showError(message) {
    uploadBox.classList.add('hidden');
    loadingSection.classList.add('hidden');
    successSection.classList.add('hidden');
    errorSection.classList.remove('hidden');
    errorMessage.textContent = message;
}

function resetConverter() {
    uploadBox.classList.remove('hidden');
    loadingSection.classList.add('hidden');
    successSection.classList.add('hidden');
    errorSection.classList.add('hidden');
    fileInput.value = '';
    currentFile = null;
    docxContent = null;
}