```javascript
// ===== CONFIGURAÇÃO PDF.JS =====
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ===== ELEMENTOS =====
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

// ===== STATE =====
let currentFile = null;
let docxContent = null;

// ===== CONVERSOR =====
class PDFToDocxConverter {

    async convertPDF(file) {

        const arrayBuffer = await file.arrayBuffer();

        const loadingTask = pdfjsLib.getDocument({
            data: arrayBuffer
        });

        const pdf = await loadingTask.promise;

        const totalPages = pdf.numPages;

        let allParagraphs = [];

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {

            loadingText.textContent =
                `Processando página ${pageNum} de ${totalPages}...`;

            updateProgress((pageNum / totalPages) * 100);

            const page = await pdf.getPage(pageNum);

            const textContent = await page.getTextContent();

            const lines = [];

            textContent.items.forEach(item => {

                if (!item.str) return;

                const text = String(item.str).trim();

                if (text.length > 0) {
                    lines.push(text);
                }

            });

            const pageText = lines.join(' ');

            if (pageText.trim().length > 0) {
                allParagraphs.push(pageText);
            }

            // quebra de página
            if (pageNum < totalPages) {
                allParagraphs.push('[[PAGE_BREAK]]');
            }
        }

        if (allParagraphs.length === 0) {
            throw new Error('Nenhum texto encontrado neste PDF.');
        }

        return await this.createDocx(allParagraphs);
    }

    async createDocx(paragraphs) {

        const zip = new JSZip();

        // pastas
        zip.folder('_rels');
        zip.folder('docProps');
        zip.folder('word');
        zip.folder('word/_rels');

        // arquivos principais
        zip.file('[Content_Types].xml', this.getContentTypesXML());

        zip.folder('_rels')
            .file('.rels', this.getRootRelsXML());

        zip.folder('docProps')
            .file('app.xml', this.getAppXML());

        zip.folder('docProps')
            .file('core.xml', this.getCoreXML());

        zip.folder('word')
            .file('document.xml', this.generateDocumentXML(paragraphs));

        zip.folder('word')
            .file('styles.xml', this.getStylesXML());

        zip.folder('word')
            .file('fontTable.xml', this.getFontTableXML());

        zip.folder('word')
            .file('settings.xml', this.getSettingsXML());

        zip.folder('word')
            .file('webSettings.xml', this.getWebSettingsXML());

        zip.folder('word')
            .file('numbering.xml', this.getNumberingXML());

        zip.folder('word/_rels')
            .file('document.xml.rels', this.getDocumentRelsXML());

        return await zip.generateAsync({
            type: 'blob',
            mimeType:
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
    }

    generateDocumentXML(paragraphs) {

        let bodyContent = '';

        paragraphs.forEach(paragraph => {

            if (paragraph === '[[PAGE_BREAK]]') {

                bodyContent += `
<w:p>
    <w:r>
        <w:br w:type="page"/>
    </w:r>
</w:p>
`;

                return;
            }

            const safeText = this.escapeXML(paragraph);

            bodyContent += `
<w:p>
    <w:r>
        <w:rPr>
            <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
            <w:sz w:val="22"/>
        </w:rPr>
        <w:t xml:space="preserve">${safeText}</w:t>
    </w:r>
</w:p>
`;
        });

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:w10="urn:schemas-microsoft-com:office:word"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
xmlns:wne="http://schemas.microsoft.com/office/2006/wordml"
xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
mc:Ignorable="w14 wp14">

<w:body>

${bodyContent}

<w:sectPr>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar
        w:top="1440"
        w:right="1440"
        w:bottom="1440"
        w:left="1440"
        w:header="708"
        w:footer="708"
        w:gutter="0"/>
</w:sectPr>

</w:body>
</w:document>`;
    }

    escapeXML(text) {

        if (!text) return '';

        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    getContentTypesXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

<Default Extension="rels"
ContentType="application/vnd.openxmlformats-package.relationships+xml"/>

<Default Extension="xml"
ContentType="application/xml"/>

<Override PartName="/word/document.xml"
ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>

<Override PartName="/word/styles.xml"
ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>

<Override PartName="/word/settings.xml"
ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>

<Override PartName="/docProps/core.xml"
ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>

<Override PartName="/docProps/app.xml"
ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>

</Types>`;
    }

    getRootRelsXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

<Relationship
Id="rId1"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
Target="word/document.xml"/>

<Relationship
Id="rId2"
Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
Target="docProps/core.xml"/>

<Relationship
Id="rId3"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
Target="docProps/app.xml"/>

</Relationships>`;
    }

    getDocumentRelsXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    }

    getStylesXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
<w:name w:val="Normal"/>
</w:style>

</w:styles>`;
    }

    getFontTableXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fontTable xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    }

    getSettingsXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    }

    getWebSettingsXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    }

    getNumberingXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
    }

    getAppXML() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties
xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">

<Application>PDF to Word Converter</Application>

</Properties>`;
    }

    getCoreXML() {

        const now = new Date().toISOString();

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:dcmitype="http://purl.org/dc/dcmitype/"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">

<dc:title>PDF Convertido</dc:title>

<dcterms:created xsi:type="dcterms:W3CDTF">
${now}
</dcterms:created>

</cp:coreProperties>`;
    }
}

// ===== EVENTOS =====
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

// ===== PROCESSAMENTO =====
function handleFile(file) {

    if (!file) {
        showError('Arquivo inválido.');
        return;
    }

    if (
        file.type !== 'application/pdf' &&
        !file.name.toLowerCase().endsWith('.pdf')
    ) {
        showError('Selecione um PDF válido.');
        return;
    }

    if (file.size > 50 * 1024 * 1024) {
        showError('Arquivo muito grande. Máximo 50MB.');
        return;
    }

    currentFile = file;

    convertPdfToWord(file);
}

async function convertPdfToWord(file) {

    try {

        showLoading();

        updateProgress(5);

        if (typeof JSZip === 'undefined') {
            throw new Error('JSZip não carregado.');
        }

        if (typeof pdfjsLib === 'undefined') {
            throw new Error('PDF.js não carregado.');
        }

        const converter = new PDFToDocxConverter();

        docxContent = await converter.convertPDF(file);

        if (!docxContent) {
            throw new Error('Falha ao gerar DOCX.');
        }

        updateProgress(100);

        loadingText.textContent =
            'Conversão concluída com sucesso!';

        setTimeout(() => {
            showSuccess();
        }, 500);

    } catch (err) {

        console.error(err);

        showError(
            err.message ||
            'Erro ao converter PDF.'
        );
    }
}

// ===== DOWNLOAD =====
function downloadWord() {

    try {

        if (!docxContent) {
            throw new Error('Nenhum DOCX disponível.');
        }

        const url = URL.createObjectURL(docxContent);

        const a = document.createElement('a');

        const cleanName = currentFile.name
            .replace(/\.pdf$/i, '')
            .replace(/[^\w\-]/g, '_');

        a.href = url;
        a.download = cleanName + '.docx';

        document.body.appendChild(a);

        a.click();

        document.body.removeChild(a);

        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 1000);

    } catch (err) {

        console.error(err);

        showError('Erro ao baixar arquivo.');
    }
}

// ===== UI =====
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

    progressText.textContent =
        Math.round(percent) + '%';
}
```
