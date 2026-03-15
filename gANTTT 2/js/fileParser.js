// js/fileParser.js

/**
 * Innehåller funktioner för att tolka innehållet från olika filtyper
 * på klientsidan. Används av AI-assistenten för att förstå kontext från dokument.
 */

/**
 * Dynamiskt laddar ett externt skript om det inte redan finns på sidan.
 * @param {string} url - URL till skriptet som ska laddas.
 * @returns {Promise<void>} En promise som resolverar när skriptet har laddats.
 */
function loadScript(url) {
    return new Promise((resolve, reject) => {
        if (document.querySelector(`script[src="${url}"]`)) {
            return resolve();
        }
        const script = document.createElement('script');
        script.src = url;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`Kunde inte ladda skript: ${url}`));
        document.head.appendChild(script);
    });
}

/**
 * Tolkar en PDF-fil och extraherar textinnehållet.
 * @param {File} file - PDF-filen som ska tolkas.
 * @returns {Promise<string>} En promise som resolverar med det extraherade textinnehållet.
 */
async function parsePDFFile(file) {
    try {
        if (typeof pdfjsLib === 'undefined') {
            await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.min.js');
            pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js`;
        }

        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            fullText += textContent.items.map(item => item.str).join(' ') + '\n';
        }
        return fullText;
    } catch (error) {
        console.error("Fel vid tolkning av PDF:", error);
        return `Fel: Kunde inte tolka PDF-filen "${file.name}".`;
    }
}

/**
 * Tolkar .docx-filer genom att packa upp dem och läsa XML-innehållet.
 * @param {File} file - .docx-filen.
 * @returns {Promise<string>} En promise som resolverar med textinnehållet.
 */
async function parseDocxFile(file) {
    try {
        const content = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(content);
        const doc = zip.file("word/document.xml");
        
        if (doc) {
            const xmlText = await doc.async("string");
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlText, "text/xml");
            const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:t"));
            return paragraphs.map(node => node.textContent).join(' ');
        }
        return `Fel: "word/document.xml" kunde inte hittas i filen "${file.name}".`;
    } catch (error) {
        console.error("Fel vid tolkning av DOCX:", error);
        return `Fel: Kunde inte tolka DOCX-filen "${file.name}".`;
    }
}

/**
 * Tolkar Excel-filer (.xlsx, .xls) och extraherar innehållet från alla blad som text.
 * @param {File} file - Excel-filen.
 * @returns {Promise<string>} En promise som resolverar med textinnehållet från alla blad.
 */
async function parseExcelFile(file) {
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        let fullText = '';
        workbook.SheetNames.forEach(sheetName => {
            fullText += `--- Blad: ${sheetName} ---\n`;
            const worksheet = workbook.Sheets[sheetName];
            fullText += XLSX.utils.sheet_to_csv(worksheet) + '\n\n';
        });
        return fullText;
    } catch (error) {
        console.error("Fel vid tolkning av Excel-fil:", error);
        return `Fel: Kunde inte tolka Excel-filen "${file.name}".`;
    }
}