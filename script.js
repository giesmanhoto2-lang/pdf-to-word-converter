// ===== CONFIGURAÇÃO INICIAL =====
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// Carregar biblioteca docx.js dinamicamente
let docxLib = null;

async function loadDocxLibrary() {
    if (docxLib) return;
    
    return new Promise((resolve, reject) => {
        if (window.docx) {
            docxLib = window.docx;
            resolve();
        } else {
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.js';
            script.onload = () => {
                docxLib = window.docx;
                resolve();
            };
            script.onerror = () => reject(new Error('Falha ao carregar biblioteca de conversão'));
            document.head.appendChild(script);
        }
    });
}

// ===== ELEMENTOS DO DOM =====
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

// ===== STATE MANAGEMENT =====
let currentFile = null;
let docxContent = null;

// ===== EVENT LISTENERS =====
uploadBox.addEventListener('click', () => fileInput.click());

uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadBox.classList.add('dragover');
});

uploadBox.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadBox.classList.remove('dragover');
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
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

// ===== VALIDAÇÃO DE ARQUIVO =====
function handleFile(file) {
    // Validar tipo
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
        showError('Por favor, selecione um arquivo PDF válido.');
        return;
    }

    // Validar tamanho (100MB)
    if (file.size > 100 * 1024 * 1024) {
        showError('O arquivo é muito grande. Máximo permitido é 100MB.');
        return;
    }

    currentFile = file;
    convertPdfToWord(file);
}

// ===== CONVERSÃO PDF → WORD =====
async function convertPdfToWord(file) {
    try {
        showLoading();
        
        // Carregar biblioteca docx
        loadingText.textContent = 'Preparando ambiente...';
        await loadDocxLibrary();
        updateSubstep(1);

        // Ler arquivo
        loadingText.textContent = 'Lendo arquivo PDF...';
        const arrayBuffer = await file.arrayBuffer();
        
        // Carregar PDF com pdf.js
        loadingText.textContent = 'Processando conteúdo...';
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        updateSubstep(2);
        
        const paragraphs = [];
        let totalChars = 0;

        // Extrair texto de cada página
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            loadingText.textContent = `Processando página ${pageNum} de ${pdf.numPages}...`;
            updateProgress((pageNum / pdf.numPages) * 100);
            
            try {
                const page = await pdf.getPage(pageNum);
                const textContent = await page.getTextContent();
                
                // Agrupar texto por linhas
                const pageText = textContent.items
                    .map(item => item.str)
                    .join(' ')
                    .trim();

                if (pageText) {
                    paragraphs.push(pageText);
                    totalChars += pageText.length;
                }

                // Adicionar quebra de página (exceto na última)
                if (pageNum < pdf.numPages) {
                    paragraphs.push('\n');
                }
            } catch (pageError) {
                console.warn(`Aviso ao processar página ${pageNum}:`, pageError);
            }
        }

        if (paragraphs.length === 0) {
            throw new Error('Nenhum texto foi encontrado no PDF. Verifique se o arquivo contém texto ou se não é um PDF escaneado.');
        }

        loadingText.textContent = 'Gerando documento Word...';
        updateSubstep(3);

        // Criar documento Word
        const docContent = paragraphs.map(text => 
            new docxLib.Paragraph({
                text: text,
                spacing: {
                    line: 240,
                    lineRule: 'auto',
                    after: text === '\n' ? 0 : 200,
                },
                alignment: docxLib.AlignmentType.LEFT,
            })
        );

        const doc = new docxLib.Document({
            sections: [{
                properties: {
                    page: {
                        margins: {
                            top: 1440,    // 1 polegada
                            right: 1440,
                            bottom: 1440,
                            left: 1440,
                        },
                    },
                },
                children: docContent,
            }],
        });

        // Gerar arquivo DOCX
        docxContent = await docxLib.Packer.toBlob(doc);
        
        if (!docxContent || docxContent.size === 0) {
            throw new Error('Falha ao gerar o documento Word.');
        }

        updateProgress(100);
        showSuccess();
    } catch (error) {
        console.error('Erro detalhado:', error);
        
        let userMessage = 'Erro ao processar PDF';
        
        if (error.message.includes('PDF')) {
            userMessage = error.message;
        } else if (error.message.includes('docx')) {
            userMessage = 'Erro ao gerar documento Word. Tente novamente.';
        } else if (error.message.includes('biblioteca')) {
            userMessage = 'Erro ao carregar recursos necessários. Verifique sua conexão com a internet.';
        }
        
        showError(userMessage);
    }
}

// ===== DOWNLOAD DO ARQUIVO =====
function downloadWord() {
    if (!docxContent || !currentFile) {
        showError('Nenhum arquivo processado. Por favor, tente novamente.');
        return;
    }

    try {
        // Gerar nome do arquivo
        const fileName = currentFile.name
            .replace(/\.pdf$/i, '')
            .replace(/[^a-zA-Z0-9-_]/g, '_')
            .substring(0, 200) + '.docx';

        // Criar link de download
        const url = URL.createObjectURL(docxContent);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        
        // Cleanup
        setTimeout(() => URL.revokeObjectURL(url), 100);
    } catch (error) {
        console.error('Erro ao fazer download:', error);
        showError('Erro ao fazer download do arquivo. Por favor, tente novamente.');
    }
}

// ===== STATE DISPLAY FUNCTIONS =====
function showLoading() {
    uploadBox.classList.add('hidden');
    successSection.classList.add('hidden');
    errorSection.classList.add('hidden');
    loadingSection.classList.remove('hidden');
    updateProgress(0);
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
    updateProgress(0);
    resetSubsteps();
}

// ===== UTILIDADE: ATUALIZAR PROGRESSO =====
function updateProgress(percent) {
    const progressElem = document.getElementById('progressPercent');
    if (progressElem) {
        progressElem.textContent = Math.round(percent) + '%';
    }
}

// ===== UTILIDADE: ATUALIZAR SUBSTEPS =====
function updateSubstep(step) {
    // Marcar steps anteriores como completos
    for (let i = 1; i < step; i++) {
        const substep = document.getElementById(`substep${i}`);
        if (substep) {
            substep.classList.add('active');
            substep.querySelector('.substep-icon').textContent = '✓';
        }
    }
    
    // Marcar step atual como ativo
    const currentSubstep = document.getElementById(`substep${step}`);
    if (currentSubstep) {
        currentSubstep.classList.add('active');
    }
}

function resetSubsteps() {
    for (let i = 1; i <= 3; i++) {
        const substep = document.getElementById(`substep${i}`);
        if (substep) {
            substep.classList.remove('active');
            substep.querySelector('.substep-icon').textContent = i === 1 ? '✓' : '◌';
        }
    }
}

// ===== INICIALIZAÇÃO =====
document.addEventListener('DOMContentLoaded', () => {
    // Pré-carregar biblioteca docx em background
    loadDocxLibrary().catch(() => {
        console.warn('Biblioteca será carregada quando necessário');
    });
});