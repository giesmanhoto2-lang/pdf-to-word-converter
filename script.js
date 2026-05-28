// ===== CONFIGURAÇÃO INICIAL =====
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

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
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');

// ===== STATE MANAGEMENT =====
let currentFile = null;
let docxContent = null;

// ===== CONVERTER PDF TO DOCX USING JSZip + CUSTOM XML =====
class PDFToDocxConverter {
    constructor() {
        this.pageBreak = {
            type: 'paragraph',
            pageBreakBefore: true
        };
    }

    async convertPDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        const paragraphs = [];
        const totalPages = pdf.numPages;

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            updateProgress((pageNum / totalPages) * 100);
            loadingText.textContent = `Processando página ${pageNum} de ${totalPages}...`;
            
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            
            const pageText = textContent.items
                .map(item => item.str)
                .join(' ')
                .trim();

            if (pageText) {
                paragraphs.push(pageText);
            }

            if (pageNum < totalPages) {
                paragraphs.push('\n\n');
            }
        }

        if (paragraphs.length === 0) {
            throw new Error('Nenhum texto encontrado no PDF');
        }

        return this.createDocx(paragraphs);
    }

    createDocx(paragraphs) {
        const docxml = this.generateDocXML(paragraphs);
        const zip = new JSZip();

        // Estrutura básica do DOCX
        zip.folder('word').file('document.xml', docxml);
        zip.folder('word').file('styles.xml', this.getStylesXML());
        zip.folder('word').file('fontTable.xml', this.getFontTableXML());
        zip.folder('word').file('numbering.xml', this.getNumberingXML());
        zip.folder('_rels').file('.rels', this.getRelsXML());
        zip.folder('word').folder('_rels').file('document.xml.rels', this.getDocumentRelsXML());
        zip.file('[Content_Types].xml', this.getContentTypesXML());

        return zip.generateAsync({ type: 'blob' });
    }

    generateDocXML(paragraphs) {
        const xmlParagraphs = paragraphs
            .map(text => this.escapXml(text))
            .map(text => `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`)
            .join('');

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${xmlParagraphs}
    </w:body>
</w:document>`;
    }

    escapXml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    getStylesXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                <w:sz w:val="22"/>
            </w:rPr>
        </w:rPrDefault>
    </w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:pPr>
            <w:pStyle w:val="Normal"/>
        </w:pPr>
    </w:style>
</w:styles>`;
    }

    getFontTableXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fontTable xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:font w:name="Calibri">
        <w:panose1 w:val="020B0604020202020204"/>
        <w:charset w:val="00"/>
        <w:family w:val="swiss"/>
        <w:pitch w:val="variable"/>
    </w:font>
</w:fontTable>`;
    }

    getNumberingXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    }

    getRelsXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
    }

    getDocumentRelsXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    }

    getContentTypesXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;
    }
}

// ===== EVENT LISTENERS =====
uploadBox.addEventListener('click', () => fileInput.click());

uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    uploadBox.classList.add('dragover');
});

uploadBox.addEventListener('dragleave', () => {
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

// ===== VALIDAÇÃO E PROCESSAMENTO =====
function handleFile(file) {
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
        showError('Por favor, selecione um arquivo PDF válido.');
        return;
    }

    if (file.size > 100 * 1024 * 1024) {
        showError('Arquivo muito grande. Máximo permitido é 100MB.');
        return;
    }

    currentFile = file;
    convertPdfToWord(file);
}

async function convertPdfToWord(file) {
    try {
        showLoading();
        loadingText.textContent = 'Preparando conversão...';
        updateProgress(10);

        // Verificar se JSZip está disponível
        if (typeof JSZip === 'undefined') {
            throw new Error('Biblioteca JSZip não disponível');
        }

        const converter = new PDFToDocxConverter();
        docxContent = await converter.convertPDF(file);

        if (!docxContent || docxContent.size === 0) {
            throw new Error('Falha ao gerar documento');
        }

        updateProgress(100);
        loadingText.textContent = 'Conversão concluída!';
        setTimeout(showSuccess, 500);
    } catch (error) {
        console.error('Erro:', error);
        showError(error.message || 'Erro ao processar PDF. Tente novamente.');
    }
}

// ===== DOWNLOAD =====
function downloadWord() {
    if (!docxContent || !currentFile) {
        showError('Nenhum arquivo processado.');
        return;
    }

    try {
        const fileName = currentFile.name
            .replace(/\.pdf$/i, '')
            .replace(/[^a-zA-Z0-9-_]/g, '_')
            .substring(0, 200) + '.docx';

        const url = URL.createObjectURL(docxContent);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        
        setTimeout(() => URL.revokeObjectURL(url), 100);
    } catch (error) {
        showError('Erro ao fazer download. Tente novamente.');
    }
}

// ===== DISPLAY STATES =====
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
    updateProgress(0);
}

function updateProgress(percent) {
    progressBar.style.width = percent + '%';
    progressText.textContent = Math.round(percent) + '%';
}