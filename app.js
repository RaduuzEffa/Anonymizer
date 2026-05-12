// 1. Service Worker & DB Init
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}
const db = new Dexie('AnonymizerDB');
db.version(1).stores({ documentMaps: '++id, docHash, fileName, createdAt' });

// 2. State
let uploadedFiles = []; // Array of File objects
let fileData = null; // Buffer for single file
let currentFileName = '';
let currentFileHash = '';
let currentFileType = ''; // 'xlsx', 'csv', 'txt', 'docx', 'pdf'
let currentSessionPrefix = ''; 
let searchedCustomWords = new Set();

// Konfigurace PDF.js workeru
const pdfjsLib = window['pdfjs-dist/build/pdf'];
if (pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

async function extractTextFromPDF(arrayBuffer) {
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';
    let charCount = 0;
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        let pageText = '';
        let lastY = -1;
        content.items.forEach(item => {
            if (lastY !== item.transform[5] && lastY !== -1) {
                pageText += '\n';
            } else if (lastY !== -1 && item.str.trim() !== '') {
                pageText += ' ';
            }
            pageText += item.str;
            lastY = item.transform[5];
        });
        fullText += pageText + '\n\n';
        charCount += pageText.replace(/\s/g, '').length;
    }
    
    if (charCount < pdf.numPages * 50) {
        if (confirm('Tento PDF dokument vypadá jako naskenovaný obrázek (neobsahuje čitelný text). Chcete spustit rozpoznávání textu pomocí umělé inteligence (OCR)? Zpracování může trvat i několik minut.')) {
            showLoader('Spouštím umělou inteligenci pro čtení obrazu. Stahuji OCR modely...');
            const worker = await Tesseract.createWorker('ces');
            fullText = '';
            for (let i = 1; i <= pdf.numPages; i++) {
                showLoader(`OCR čtení obrázku: Strana ${i} z ${pdf.numPages}...`);
                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;
                await page.render({ canvasContext: ctx, viewport: viewport }).promise;
                const ret = await worker.recognize(canvas);
                fullText += ret.data.text + '\n\n';
            }
            await worker.terminate();
        }
    }
    return fullText;
}
let foundEntities = []; 
let entityIdCounter = 0;
let excelHeaders = []; 

let currentWorkbookExcelJS = null; 
let currentWorkbookSheetJS = null; 
let currentTextContent = '';

// 3. Regex Patterns
const REGEX_PATTERNS = {
    RC: /\b\d{6}\/?\d{3,4}\b/g,
    EMAIL: /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi
};

// 4. DOM Elements
const uploadSection = document.getElementById('upload-section');
const actionSelection = document.getElementById('action-selection');
const searchSection = document.getElementById('search-section');
const deanonSection = document.getElementById('deanon-section');
const actionSingleFile = document.getElementById('action-single-file');
const actionMultiFile = document.getElementById('action-multi-file');

const resultsArea = document.getElementById('results-area');
const entitiesContainer = document.getElementById('entities-container');
const excelColumnsWrapper = document.getElementById('excel-columns-wrapper');
const excelColumnsContainer = document.getElementById('excel-columns-container');
const dbMapsSelect = document.getElementById('db-maps-select');
const manualMapSelector = document.getElementById('manual-map-selector');

const fileInput = document.getElementById('file-upload');
const uploadArea = document.querySelector('.upload-area');
const currentFilenameActionEl = document.getElementById('current-filename-action');
const detectionSummaryEl = document.getElementById('detection-summary');

const btnSearch = document.getElementById('btn-search');
const btnToggleAll = document.getElementById('btn-toggle-all');
const btnAnonymizeExport = document.getElementById('btn-anonymize-export');
const btnExecuteDeanon = document.getElementById('btn-execute-deanon');
const btnChooseMergeDeanon = document.getElementById('btn-choose-merge-deanon');
const btnChooseMergeOnly = document.getElementById('btn-choose-merge-only');

const loader = document.getElementById('loader');
const loaderText = document.getElementById('loader-text');
const loaderSpinner = document.getElementById('loader-spinner');

// Helpers
async function getFileHash(buffer) {
    const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

function generateTag(type, index) { 
    return `{{${type}_${currentSessionPrefix}_${String(index).padStart(3, '0')}}}`; 
}

function showLoader(text, colorClass = 'border-t-secondary') { 
    loaderText.textContent = text; 
    loaderSpinner.className = `w-16 h-16 border-4 border-gray-200 rounded-full animate-spin mb-4 ${colorClass}`;
    loader.style.display = 'flex'; 
}
function hideLoader() { loader.style.display = 'none'; }

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

function getCellValueString(val) {
    if (val === null || val === undefined) return '';
    if (typeof val === 'object') {
        if (val.richText) return val.richText.map(rt => rt.text).join('');
        if (val.formula || val.sharedFormula) return ''; // Vzorce ignorujeme
        if (val.text) return String(val.text); 
    }
    return String(val);
}

function safeReplaceCellValue(cell, replacerFunc) {
    if (cell.value === null || cell.value === undefined) return;
    
    if (typeof cell.value === 'object') {
        if (cell.value.richText) {
            let changed = false;
            const newRichText = cell.value.richText.map(rt => {
                const rep = replacerFunc(rt.text);
                if (rep !== rt.text) changed = true;
                return { ...rt, text: rep };
            });
            if (changed) cell.value = { richText: newRichText };
        } else if (cell.value.formula || cell.value.sharedFormula) {
            // Vzorce přísně ignorujeme! Excel si je po spuštění přepočítá sám.
            // Zápis do vzorců způsoboval "Chybu s obsahem" a poškození souboru.
            return;
        } else if (cell.value.text && cell.value.hyperlink) {
            const oldTxt = String(cell.value.text);
            const newTxt = replacerFunc(oldTxt);
            if (oldTxt !== newTxt) cell.value = { text: newTxt, hyperlink: cell.value.hyperlink };
        } else if (cell.value instanceof Date) {
            // Nemenime Date
        } else {
            const oldTxt = String(cell.value);
            const newTxt = replacerFunc(oldTxt);
            if (oldTxt !== newTxt) cell.value = newTxt;
        }
    } else {
        const oldTxt = String(cell.value);
        const newTxt = replacerFunc(oldTxt);
        if (oldTxt !== newTxt) cell.value = newTxt;
    }
}

// Upload Handlers
uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('bg-blue-50', 'border-blue-500'); });
uploadArea.addEventListener('dragleave', () => { uploadArea.classList.remove('bg-blue-50', 'border-blue-500'); });
uploadArea.addEventListener('drop', e => {
    e.preventDefault(); uploadArea.classList.remove('bg-blue-50', 'border-blue-500');
    if (e.dataTransfer.files.length) initFiles(Array.from(e.dataTransfer.files));
});
fileInput.addEventListener('change', e => {
    if (e.target.files.length) initFiles(Array.from(e.target.files));
});

document.getElementById('btn-cancel-action').addEventListener('click', resetApp);
document.getElementById('btn-back-from-anon').addEventListener('click', () => {
    searchSection.classList.add('hidden');
    actionSelection.classList.remove('hidden');
});
document.getElementById('btn-back-from-deanon').addEventListener('click', () => {
    deanonSection.classList.add('hidden');
    actionSelection.classList.remove('hidden');
});

function resetApp() {
    uploadedFiles = []; fileData = null; currentWorkbookExcelJS = null; currentWorkbookSheetJS = null; currentTextContent = '';
    if (fileInput) fileInput.value = ''; // FIX pro opakované nahrání stejného souboru
    searchedCustomWords.clear();
    if (document.getElementById('search-custom')) document.getElementById('search-custom').value = '';
    if (document.getElementById('active-search-terms')) document.getElementById('active-search-terms').innerHTML = '';
    searchSection.classList.add('hidden');
    deanonSection.classList.add('hidden');
    actionSelection.classList.add('hidden');
    resultsArea.classList.add('hidden');
    excelColumnsWrapper.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    foundEntities = []; excelHeaders = [];
    updateExportButtonState();
}

async function initFiles(files) {
    if (files.length === 0) return;
    
    const ext1 = files[0].name.split('.').pop().toLowerCase();
    const isMultiValid = Array.from(files).every(f => f.name.toLowerCase().endsWith(ext1));

    if (files.length > 1) {
        if (!isMultiValid || !['xlsx', 'xls', 'docx', 'doc', 'txt', 'csv', 'pdf'].includes(ext1)) {
            alert('Pro hromadné slučování prosím nahrajte soubory stejného typu (.xlsx, .docx, .txt, .csv nebo .pdf).');
            return;
        }
        uploadedFiles = Array.from(files);
        currentFileType = (ext1 === 'xls') ? 'csv' : ext1;
        
        currentFilenameActionEl.textContent = `${files.length} souborů připraveno ke sloučení`;
        actionSingleFile.classList.add('hidden');
        actionMultiFile.classList.remove('hidden');
        uploadSection.classList.add('hidden');
        actionSelection.classList.remove('hidden');
        renderMergeFileList();
        return;
    }

    const file = files[0];
    uploadedFiles = [file];
    currentSessionPrefix = Math.random().toString(36).substring(2, 6).toUpperCase();
    showLoader('Načítám soubor...');
    currentFileName = file.name;
    currentFilenameActionEl.textContent = currentFileName;
    
    try {
        const arrayBuffer = await file.arrayBuffer();
        currentFileHash = await getFileHash(arrayBuffer);
        fileData = arrayBuffer;

        const ext = file.name.split('.').pop().toLowerCase();
        if (ext === 'doc') {
            alert('UPOZORNĚNÍ: Zastaralý formát .doc (Word 97-2003) nelze z technických důvodů zpracovávat přímo v prohlížeči. Otevřete prosím soubor ve Wordu a uložte jej jako moderní formát .docx. Ten už aplikace bez problému zpracuje!');
            hideLoader(); return;
        } else if (ext === 'xlsx') {
            currentFileType = 'xlsx';
            currentWorkbookExcelJS = new ExcelJS.Workbook();
            await currentWorkbookExcelJS.xlsx.load(arrayBuffer);
        } else if (ext === 'xls' || ext === 'csv') {
            if (ext === 'xls') {
                alert('UPOZORNĚNÍ: Zastaralý formát .xls neumožňuje zachovat vizuální formátování. Aplikace jej zpracuje jako prostá data.');
            }
            currentFileType = 'csv';
            currentWorkbookSheetJS = XLSX.read(arrayBuffer, { type: 'array' });
        } else if (ext === 'txt') {
            currentFileType = 'txt';
            currentTextContent = new TextDecoder('utf-8').decode(arrayBuffer);
        } else if (ext === 'docx') {
            currentFileType = 'docx';
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            currentTextContent = result.value;
        } else if (ext === 'pdf') {
            currentFileType = 'pdf';
            currentTextContent = await extractTextFromPDF(arrayBuffer);
        } else {
            alert('Tento formát není podporován.');
            hideLoader(); return;
        }

        actionSingleFile.classList.remove('hidden');
        actionMultiFile.classList.add('hidden');
        uploadSection.classList.add('hidden');
        actionSelection.classList.remove('hidden');
    } catch (err) {
        console.error(err); alert('Došlo k chybě při čtení souboru.');
    } finally {
        hideLoader();
    }
}

// Navigace
document.getElementById('btn-choose-anon').addEventListener('click', () => {
    actionSelection.classList.add('hidden');
    searchSection.classList.remove('hidden');
    if (currentFileType === 'xlsx' || currentFileType === 'csv') {
        extractExcelHeaders();
        excelColumnsWrapper.classList.remove('hidden');
    } else {
        excelColumnsWrapper.classList.add('hidden');
    }
});

document.getElementById('btn-choose-deanon').addEventListener('click', async () => {
    actionSelection.classList.add('hidden');
    deanonSection.classList.remove('hidden');
    await loadDbMaps();
});

// ==========================================
// MERGE LOGIKA (S NEBO BEZ DEANONYMIZACE)
// ==========================================

function renderMergeFileList() {
    const listEl = document.getElementById('merge-file-list');
    listEl.innerHTML = '';
    uploadedFiles.forEach((file, index) => {
        const li = document.createElement('li');
        li.className = "flex justify-between items-center bg-white p-3 border border-gray-200 rounded-lg shadow-sm";
        
        const fileNameSpan = document.createElement('span');
        fileNameSpan.className = "font-medium text-gray-700 truncate flex-grow mr-4";
        fileNameSpan.textContent = `${index + 1}. ${file.name}`;
        
        const btnGroup = document.createElement('div');
        btnGroup.className = "flex gap-2";
        
        const btnUp = document.createElement('button');
        btnUp.className = `p-2 rounded-lg text-sm transition ${index === 0 ? 'text-gray-300 bg-gray-50 cursor-not-allowed' : 'bg-gray-100 hover:bg-gray-200 text-gray-700'}`;
        btnUp.innerHTML = '<i class="fa-solid fa-arrow-up"></i>';
        btnUp.disabled = index === 0;
        btnUp.onclick = () => {
            if (index > 0) {
                [uploadedFiles[index - 1], uploadedFiles[index]] = [uploadedFiles[index], uploadedFiles[index - 1]];
                renderMergeFileList();
            }
        };

        const btnDown = document.createElement('button');
        btnDown.className = `p-2 rounded-lg text-sm transition ${index === uploadedFiles.length - 1 ? 'text-gray-300 bg-gray-50 cursor-not-allowed' : 'bg-gray-100 hover:bg-gray-200 text-gray-700'}`;
        btnDown.innerHTML = '<i class="fa-solid fa-arrow-down"></i>';
        btnDown.disabled = index === uploadedFiles.length - 1;
        btnDown.onclick = () => {
            if (index < uploadedFiles.length - 1) {
                [uploadedFiles[index], uploadedFiles[index + 1]] = [uploadedFiles[index + 1], uploadedFiles[index]];
                renderMergeFileList();
            }
        };

        btnGroup.appendChild(btnUp);
        btnGroup.appendChild(btnDown);
        
        li.appendChild(fileNameSpan);
        li.appendChild(btnGroup);
        listEl.appendChild(li);
    });
}

btnChooseMergeDeanon.addEventListener('click', () => executeMerge(true));
btnChooseMergeOnly.addEventListener('click', () => executeMerge(false));

async function executeMerge(shouldDeanonymize) {
    showLoader(shouldDeanonymize ? 'Slučuji a de-anonymizuji soubory...' : 'Pouze slučuji soubory...', 'border-t-purple-500');
    
    setTimeout(async () => {
        try {
            let finalMapping = {};
            let deanonRegex = null;

            if (shouldDeanonymize) {
                const allMaps = await db.documentMaps.toArray();
                allMaps.forEach(m => Object.assign(finalMapping, m.mapping));
                deanonRegex = /\{\{[A-Z]+(?:_[a-zA-Z0-9]+)?_\d+\}\}/g;
            }

            if (currentFileType === 'xlsx') {
                const mergedWb = new ExcelJS.Workbook();
                const seenRowsHashes = new Set();
                for (let i = 0; i < uploadedFiles.length; i++) {
                    const wb = new ExcelJS.Workbook();
                    await wb.xlsx.load(await uploadedFiles[i].arrayBuffer());
                    wb.eachSheet(sheet => {
                        let mergedSheet = mergedWb.getWorksheet(sheet.name) || mergedWb.addWorksheet(sheet.name);
                        sheet.eachRow((row, rowNumber) => {
                            if (!row.hasValues) return;
                            const rowVals = [];
                            row.eachCell({ includeEmpty: true }, cell => rowVals.push(getCellValueString(cell.value)));
                            const rowHash = JSON.stringify(rowVals);
                            if (!seenRowsHashes.has(rowHash)) {
                                seenRowsHashes.add(rowHash);
                                const newRow = mergedSheet.addRow([]);
                                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                                    const newCell = newRow.getCell(colNumber);
                                    newCell.value = cell.value;
                                    if (shouldDeanonymize) {
                                        safeReplaceCellValue(newCell, txt => txt.replace(deanonRegex, match => finalMapping[match] !== undefined ? finalMapping[match] : match));
                                    }
                                });
                            }
                        });
                    });
                }
                const buffer = await mergedWb.xlsx.writeBuffer();
                const dateStr = new Date().toISOString().replace(/[-:T]/g, '').slice(0,14);
                const exportName = shouldDeanonymize ? `merged_deanonymized_${dateStr}.xlsx` : `merged_only_${dateStr}.xlsx`;
                downloadBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), exportName);
                hideLoader();
                alert(shouldDeanonymize ? 'Sloučení a de-anonymizace úspěšně dokončena!' : 'Sloučení souborů úspěšně dokončeno!');
                resetApp();

            } else if (currentFileType === 'docx') {
                let mergedHtml = "";
                for (let i = 0; i < uploadedFiles.length; i++) {
                    const arrayBuffer = await uploadedFiles[i].arrayBuffer();
                    const result = await mammoth.convertToHtml({ arrayBuffer });
                    let html = result.value;
                    if (shouldDeanonymize) {
                        html = html.replace(deanonRegex, match => finalMapping[match] !== undefined ? finalMapping[match] : match);
                    }
                    mergedHtml += html;
                    if (i < uploadedFiles.length - 1) mergedHtml += "<br><hr><br>";
                }
                const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${mergedHtml}</body></html>`;
                const blob = htmlDocx.asBlob(wrappedHtml);
                const dateStr = new Date().toISOString().replace(/[-:T]/g, '').slice(0,14);
                const exportName = shouldDeanonymize ? `merged_deanonymized_${dateStr}.docx` : `merged_only_${dateStr}.docx`;
                downloadBlob(blob, exportName);
                hideLoader();
                alert(shouldDeanonymize ? 'Sloučení a de-anonymizace úspěšně dokončena!' : 'Sloučení souborů úspěšně dokončeno!');
                resetApp();
                
            } else if (currentFileType === 'txt' || currentFileType === 'csv' || currentFileType === 'pdf') {
                let mergedTxt = "";
                for (let i = 0; i < uploadedFiles.length; i++) {
                    const arrayBuffer = await uploadedFiles[i].arrayBuffer();
                    let txt = "";
                    if (currentFileType === 'pdf') {
                        txt = await extractTextFromPDF(arrayBuffer);
                    } else {
                        txt = new TextDecoder().decode(arrayBuffer);
                    }
                    if (shouldDeanonymize) {
                        txt = txt.replace(deanonRegex, match => finalMapping[match] !== undefined ? finalMapping[match] : match);
                    }
                    mergedTxt += txt;
                    if (i < uploadedFiles.length - 1) mergedTxt += "\n\n---\n\n";
                }
                const dateStr = new Date().toISOString().replace(/[-:T]/g, '').slice(0,14);
                
                if (currentFileType === 'pdf') {
                    const html = mergedTxt.replace(/\n/g, '<br>');
                    const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
                    const blob = htmlDocx.asBlob(wrappedHtml);
                    const exportName = shouldDeanonymize ? `merged_deanonymized_${dateStr}.docx` : `merged_only_${dateStr}.docx`;
                    downloadBlob(blob, exportName);
                } else {
                    const exportName = shouldDeanonymize ? `merged_deanonymized_${dateStr}.${currentFileType}` : `merged_only_${dateStr}.${currentFileType}`;
                    downloadBlob(new Blob([mergedTxt], { type: 'text/plain' }), exportName);
                }
                
                hideLoader();
                alert(shouldDeanonymize ? 'Sloučení a de-anonymizace úspěšně dokončena!' : 'Sloučení souborů úspěšně dokončeno!');
                resetApp();
            }
        } catch (err) {
            console.error(err);
            alert('Chyba při slučování souborů: ' + (err.message || err));
            hideLoader();
        }
    }, 100);
}

// ==========================================
// ANONYMIZACE LOGIKA
// ==========================================

function extractExcelHeaders() {
    excelHeaders = [];
    if (currentFileType === 'xlsx') {
        currentWorkbookExcelJS.eachSheet((worksheet) => {
            const row = worksheet.getRow(1);
            row.eachCell((cell, colNumber) => {
                const headerName = getCellValueString(cell.value) || `Sloupec ${colNumber}`;
                excelHeaders.push({ id: `col-${worksheet.name}-${colNumber}`, sheet: worksheet.name, colIndex: colNumber, headerName: headerName });
            });
        });
    } else {
        currentWorkbookSheetJS.SheetNames.forEach(sheetName => {
            const sheet = currentWorkbookSheetJS.Sheets[sheetName];
            if(!sheet['!ref']) return;
            const range = XLSX.utils.decode_range(sheet['!ref']);
            const R = range.s.r;
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell = sheet[XLSX.utils.encode_cell({c: C, r: R})];
                let headerName = cell && cell.v ? String(cell.v) : `Sloupec ${C+1}`;
                excelHeaders.push({ id: `col-${sheetName}-${C}`, sheet: sheetName, colIndex: C, headerName: headerName });
            }
        });
    }
    
    excelColumnsContainer.innerHTML = '';
    excelHeaders.forEach(col => {
        const lbl = document.createElement('label');
        lbl.className = 'flex items-center gap-2 p-2 bg-white border border-indigo-100 rounded-lg cursor-pointer hover:bg-indigo-50 transition shadow-sm';
        lbl.innerHTML = `
            <input type="checkbox" value="${col.id}" class="col-checkbox rounded text-indigo-600 focus:ring-indigo-500">
            <span class="text-sm font-medium text-indigo-900 truncate">${col.headerName} <span class="text-xs text-indigo-400">(${col.sheet})</span></span>
        `;
        lbl.querySelector('input').addEventListener('change', updateExportButtonState);
        excelColumnsContainer.appendChild(lbl);
    });
}

document.getElementById('btn-cols-all').addEventListener('click', () => { document.querySelectorAll('.col-checkbox').forEach(cb => cb.checked = true); updateExportButtonState(); });
document.getElementById('btn-cols-none').addEventListener('click', () => { document.querySelectorAll('.col-checkbox').forEach(cb => cb.checked = false); updateExportButtonState(); });

function renderActiveSearchTerms() {
    const container = document.getElementById('active-search-terms');
    if (!container) return;
    container.innerHTML = '';
    searchedCustomWords.forEach(word => {
        const span = document.createElement('span');
        span.className = 'inline-flex items-center gap-1 bg-blue-100 text-blue-800 text-sm font-medium px-3 py-1 rounded-full shadow-sm';
        span.innerHTML = `${word} <button type="button" class="ml-1 text-blue-600 hover:text-red-600 focus:outline-none"><i class="fa-solid fa-xmark"></i></button>`;
        span.querySelector('button').addEventListener('click', () => {
            searchedCustomWords.delete(word);
            renderActiveSearchTerms();
            showLoader('Aktualizuji vyhledávání...');
            setTimeout(() => { performSearch(); hideLoader(); }, 100);
        });
        container.appendChild(span);
    });
}

btnSearch.addEventListener('click', () => {
    const customWordInput = document.getElementById('search-custom').value.trim();
    if (customWordInput) {
        customWordInput.split(',').map(w => w.trim()).filter(w => w.length > 0).forEach(w => searchedCustomWords.add(w));
        document.getElementById('search-custom').value = '';
        renderActiveSearchTerms();
    }
    showLoader('Prohledávám dokument...');
    setTimeout(() => { performSearch(); hideLoader(); }, 100);
});

function performSearch() {
    foundEntities = [];
    entityIdCounter = 0;
    const searchRC = document.getElementById('search-rc').checked;
    const searchEmail = document.getElementById('search-email').checked;

    const searchInText = (text, contextInfo) => {
        if (!text) return;
        text = String(text);
        const addEntity = (match, type) => {
            foundEntities.push({ id: ++entityIdCounter, originalText: match, type: type, context: contextInfo, checked: true });
        };
        if (searchRC) { const m = text.match(REGEX_PATTERNS.RC); if (m) m.forEach(x => addEntity(x, 'RC')); }
        if (searchEmail) { const m = text.match(REGEX_PATTERNS.EMAIL); if (m) m.forEach(x => addEntity(x, 'EMAIL')); }
        if (searchedCustomWords.size > 0) {
            const escapeRegExp = (string) => string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            searchedCustomWords.forEach(word => {
                // Nahradíme běžné mezery regexem \s+, aby to chytlo i nezlomitelné mezery z Wordu atd.
                const customRegex = new RegExp(escapeRegExp(word).replace(/\\s+/g, '\\s+').replace(/\s+/g, '\\s+'), 'gi');
                const m = text.match(customRegex);
                if (m) m.forEach(x => addEntity(x, 'CUSTOM'));
            });
        }
    };

    if (currentFileType === 'xlsx') {
        currentWorkbookExcelJS.eachSheet((worksheet) => {
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    searchInText(getCellValueString(cell.value), `List: ${worksheet.name} | Řádek: ${rowNumber}, Sloupec: ${colNumber}`);
                });
            });
        });
    } else if (currentFileType === 'csv') {
        currentWorkbookSheetJS.SheetNames.forEach(sheetName => {
            const sheet = currentWorkbookSheetJS.Sheets[sheetName];
            if(!sheet['!ref']) return;
            const range = XLSX.utils.decode_range(sheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell = sheet[XLSX.utils.encode_cell({ c: C, r: R })];
                    if (cell && cell.t === 's') searchInText(cell.v, `List: ${sheetName} | Buňka: ${XLSX.utils.encode_cell({ c: C, r: R })}`);
                }
            }
        });
    } else {
        currentTextContent.split('\n').forEach((line, index) => searchInText(line, `Řádek ${index + 1}`));
    }
    renderEntities();
    resultsArea.classList.remove('hidden');
}

function renderEntities() {
    entitiesContainer.innerHTML = '';
    if (foundEntities.length === 0) {
        detectionSummaryEl.textContent = 'Nebyla nalezena žádná shoda.';
        entitiesContainer.innerHTML = '<div class="p-4 text-center text-gray-500 border border-dashed border-gray-300 rounded-lg bg-gray-50">Nic nebylo nalezeno.</div>';
        updateExportButtonState(); return;
    }
    detectionSummaryEl.innerHTML = `Nalezeno <strong class="text-secondary">${foundEntities.length}</strong> výskytů.`;
    foundEntities.forEach(entity => {
        const item = document.createElement('label');
        item.className = 'flex items-center gap-4 p-3 bg-white border border-gray-100 rounded-lg shadow-sm cursor-pointer';
        let typeColor = 'bg-gray-100 text-gray-700'; let typeIcon = 'fa-tag';
        if (entity.type === 'RC') { typeColor = 'bg-purple-100 text-purple-700'; typeIcon = 'fa-id-card'; }
        if (entity.type === 'EMAIL') { typeColor = 'bg-green-100 text-green-700'; typeIcon = 'fa-envelope'; }
        if (entity.type === 'CUSTOM') { typeColor = 'bg-blue-100 text-blue-700'; typeIcon = 'fa-pen'; }

        item.innerHTML = `
            <input type="checkbox" class="entity-checkbox" data-id="${entity.id}" ${entity.checked ? 'checked' : ''}>
            <div class="flex flex-col flex-grow">
                <span class="font-semibold text-gray-800">${entity.originalText}</span>
                <span class="text-xs text-gray-500">${entity.context}</span>
            </div>
            <div class="${typeColor} px-2 py-1 rounded text-xs font-bold flex items-center gap-1 uppercase">
                <i class="fa-solid ${typeIcon}"></i> ${entity.type}
            </div>
        `;
        item.querySelector('input').addEventListener('change', e => { entity.checked = e.target.checked; updateExportButtonState(); });
        entitiesContainer.appendChild(item);
    });
    updateExportButtonState();
}

let allChecked = true;
btnToggleAll.addEventListener('click', () => {
    allChecked = !allChecked;
    foundEntities.forEach(e => e.checked = allChecked);
    btnToggleAll.textContent = allChecked ? 'Odznačit vše' : 'Označit vše';
    renderEntities();
});

function updateExportButtonState() {
    const hasCheckedEntities = foundEntities.some(e => e.checked);
    const hasCheckedCols = document.querySelectorAll('.col-checkbox:checked').length > 0;
    if (hasCheckedEntities || hasCheckedCols) {
        btnAnonymizeExport.disabled = false; btnAnonymizeExport.classList.remove('opacity-50', 'cursor-not-allowed');
    } else {
        btnAnonymizeExport.disabled = true; btnAnonymizeExport.classList.add('opacity-50', 'cursor-not-allowed');
    }
}

btnAnonymizeExport.addEventListener('click', async () => {
    showLoader('Vytvářím anonymizovaný soubor...', 'border-t-secondary');
    setTimeout(async () => {
        try {
            const mapping = {};
            let tagCounters = { RC: 0, EMAIL: 0, CUSTOM: 0, COL: 0 };
            const checkedEntities = foundEntities.filter(e => e.checked);
            checkedEntities.sort((a, b) => b.originalText.length - a.originalText.length);
            checkedEntities.forEach(entity => {
                if (!entity.tag) { 
                    tagCounters[entity.type]++;
                    entity.tag = generateTag(entity.type, tagCounters[entity.type]);
                    mapping[entity.tag] = entity.originalText;
                }
            });

            if (currentFileType === 'xlsx') {
                const wb = await XlsxPopulate.fromDataAsync(fileData.slice(0));
                const checkedCols = Array.from(document.querySelectorAll('.col-checkbox:checked')).map(cb => cb.value);
                const colValueMap = {};
                wb.sheets().forEach(sheet => {
                    const sheetName = sheet.name();
                    const sheetColsToAnon = excelHeaders.filter(h => h.sheet === sheetName && checkedCols.includes(h.id)).map(h => h.colIndex);
                    const usedRange = sheet.usedRange();
                    if (!usedRange) return;
                    usedRange.cells().forEach(row => {
                        row.forEach(cell => {
                            if (!cell || cell.formula()) return;
                            const val = cell.value();
                            if (val === undefined || val === null || val instanceof Date) return;
                            let originalVal = (typeof val === 'object' && typeof val.text === 'function') ? val.text() : String(val);
                            if (cell.rowNumber() > 1 && sheetColsToAnon.includes(cell.columnNumber()) && originalVal.trim() !== '') {
                                let tagKey = `COL_${originalVal}`;
                                if (!colValueMap[tagKey]) {
                                    tagCounters.COL++;
                                    colValueMap[tagKey] = generateTag('COL', tagCounters.COL);
                                    mapping[colValueMap[tagKey]] = originalVal;
                                }
                                cell.value(colValueMap[tagKey]);
                                cell.style("fontColor", "FF0000");
                            } else {
                                let newTxt = originalVal;
                                checkedEntities.forEach(e => { if (newTxt.includes(e.originalText)) newTxt = newTxt.split(e.originalText).join(e.tag); });
                                if (newTxt !== originalVal) {
                                    cell.value(newTxt);
                                    cell.style("fontColor", "FF0000");
                                }
                            }
                        });
                    });
                });
                const blob = await wb.outputAsync();
                downloadBlob(blob, `anonymized_${currentFileName}`);
            } else if (currentFileType === 'csv') {
                const wb = XLSX.read(fileData, { type: 'array' }); 
                const checkedCols = Array.from(document.querySelectorAll('.col-checkbox:checked')).map(cb => cb.value);
                const colValueMap = {};
                wb.SheetNames.forEach(sheetName => {
                    const sheet = wb.Sheets[sheetName];
                    if(!sheet['!ref']) return;
                    const range = XLSX.utils.decode_range(sheet['!ref']);
                    const sheetColsToAnon = excelHeaders.filter(h => h.sheet === sheetName && checkedCols.includes(h.id)).map(h => h.colIndex);
                    for (let R = range.s.r; R <= range.e.r; ++R) {
                        for (let C = range.s.c; C <= range.e.c; ++C) {
                            const cell = sheet[XLSX.utils.encode_cell({ c: C, r: R })];
                            if (cell && cell.v !== undefined) {
                                let txt = String(cell.v);
                                if (R > range.s.r && sheetColsToAnon.includes(C) && txt.trim() !== '') {
                                    let tagKey = `COL_${txt}`;
                                    if (!colValueMap[tagKey]) {
                                        tagCounters.COL++;
                                        colValueMap[tagKey] = generateTag('COL', tagCounters.COL);
                                        mapping[colValueMap[tagKey]] = txt;
                                    }
                                    txt = colValueMap[tagKey];
                                } else {
                                    checkedEntities.forEach(e => { if (txt.includes(e.originalText)) txt = txt.split(e.originalText).join(e.tag); });
                                }
                                cell.v = txt; if (cell.w) delete cell.w;
                            }
                        }
                    }
                });
                const wbout = XLSX.write(wb, { bookType: 'csv', type: 'array' });
                downloadBlob(new Blob([wbout], { type: 'text/csv' }), `anonymized_${currentFileName}`);
            } else if (currentFileType === 'docx') {
                const arrayBuffer = await uploadedFiles[0].arrayBuffer();
                const result = await mammoth.convertToHtml({ arrayBuffer });
                let html = result.value;
                checkedEntities.forEach(e => { if (html.includes(e.originalText)) html = html.split(e.originalText).join(`<span style="color:red;font-weight:bold;">${e.tag}</span>`); });
                const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
                const blob = htmlDocx.asBlob(wrappedHtml);
                downloadBlob(blob, `anonymized_${currentFileName}`);
            } else if (currentFileType === 'pdf') {
                let outText = currentTextContent;
                checkedEntities.forEach(e => { outText = outText.split(e.originalText).join(`<span style="color:red;font-weight:bold;">${e.tag}</span>`); });
                const html = outText.replace(/\n/g, '<br>');
                const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
                const blob = htmlDocx.asBlob(wrappedHtml);
                downloadBlob(blob, `anonymized_${currentFileName.replace('.pdf', '.docx')}`);
            } else {
                let outText = currentTextContent;
                checkedEntities.forEach(e => { outText = outText.split(e.originalText).join(e.tag); });
                downloadBlob(new Blob([outText], { type: 'text/plain;charset=utf-8' }), `anonymized_${currentFileName}`);
            }
            if (Object.keys(mapping).length > 0) {
                await db.documentMaps.add({ docHash: currentFileHash, fileName: currentFileName, createdAt: new Date().toISOString(), mapping: mapping });
                alert('Anonymizace proběhla úspěšně! Mapa uložena.');
            }
            hideLoader();
            resetApp();
        } catch (err) {
            console.error(err); alert('Chyba při exportu.'); hideLoader();
        }
    }, 100);
});

// ==========================================
// DE-ANONYMIZACE LOGIKA PRO 1 SOUBOR
// ==========================================

const radioAuto = document.querySelector('input[name="deanon-mode"][value="auto"]');
const radioManual = document.querySelector('input[name="deanon-mode"][value="manual"]');

radioAuto.addEventListener('change', () => {
    manualMapSelector.classList.add('hidden');
    document.getElementById('btn-execute-deanon').disabled = false;
    document.getElementById('btn-execute-deanon').classList.remove('opacity-50', 'cursor-not-allowed');
});
radioManual.addEventListener('change', () => {
    manualMapSelector.classList.remove('hidden');
    if (!dbMapsSelect.value) {
        document.getElementById('btn-execute-deanon').disabled = true;
        document.getElementById('btn-execute-deanon').classList.add('opacity-50', 'cursor-not-allowed');
    }
});

async function loadDbMaps() {
    try {
        const maps = await db.documentMaps.orderBy('createdAt').reverse().toArray();
        dbMapsSelect.innerHTML = '';
        if (maps.length === 0) {
            dbMapsSelect.innerHTML = '<option value="">Nebyla nalezena žádná uložená mapa.</option>';
            document.getElementById('btn-delete-map').disabled = true;
            radioManual.disabled = true;
            return;
        }
        document.getElementById('btn-delete-map').disabled = false;
        radioManual.disabled = false;
        maps.forEach(map => {
            const date = new Date(map.createdAt).toLocaleString('cs-CZ');
            const keysCount = Object.keys(map.mapping).length;
            const option = document.createElement('option');
            option.value = map.id;
            option.textContent = `${map.fileName} (${date}) - ${keysCount} značek`;
            dbMapsSelect.appendChild(option);
        });
    } catch (err) { console.error('Chyba načítání map:', err); }
}

document.getElementById('btn-delete-map').addEventListener('click', async () => {
    const mapId = parseInt(dbMapsSelect.value);
    if (!mapId) return;
    if (confirm('Opravdu chcete nevratně smazat vybranou mapu klíčů?')) {
        await db.documentMaps.delete(mapId);
        alert('Mapa byla smazána.');
        loadDbMaps();
    }
});

document.getElementById('btn-delete-all-maps').addEventListener('click', async () => {
    if (confirm('VAROVÁNÍ: Chcete smazat KOMPLETNÍ HISTORII všech map?')) {
        await db.documentMaps.clear();
        alert('Databáze byla vyčištěna.');
        loadDbMaps();
    }
});

dbMapsSelect.addEventListener('change', () => {
    if (radioManual.checked) {
        const btn = document.getElementById('btn-execute-deanon');
        if (dbMapsSelect.value) {
            btn.disabled = false;
            btn.classList.remove('opacity-50', 'cursor-not-allowed');
        } else {
            btn.disabled = true;
            btn.classList.add('opacity-50', 'cursor-not-allowed');
        }
    }
});

btnExecuteDeanon.addEventListener('click', async () => {
    showLoader('De-anonymizuji soubor...', 'border-t-emerald-500');
    setTimeout(async () => {
        try {
            const modeEl = document.querySelector('input[name="deanon-mode"]:checked');
            if (!modeEl) throw new Error('Není vybrán režim.');
            const mode = modeEl.value;
            let finalMapping = {};
            if (mode === 'auto') {
                const allMaps = await db.documentMaps.toArray();
                allMaps.forEach(m => Object.assign(finalMapping, m.mapping));
            } else {
                const mapId = parseInt(dbMapsSelect.value);
                if (!mapId) throw new Error('Nebyla vybrána mapa.');
                const mapRecord = await db.documentMaps.get(mapId);
                if (mapRecord) finalMapping = mapRecord.mapping;
            }
            const deanonRegex = /\{\{[A-Z]+(?:_[a-zA-Z0-9]+)?_\d+\}\}/g;
            const replaceText = (txt) => {
                if (typeof txt !== 'string') return txt;
                return txt.replace(deanonRegex, match => finalMapping[match] !== undefined ? finalMapping[match] : match);
            };

            if (currentFileType === 'xlsx') {
                const wb = await XlsxPopulate.fromDataAsync(fileData.slice(0));
                wb.sheets().forEach(sheet => {
                    const usedRange = sheet.usedRange();
                    if (!usedRange) return;
                    usedRange.cells().forEach(row => {
                        row.forEach(cell => {
                            if (!cell || cell.formula()) return;
                            const val = cell.value();
                            if (val === undefined || val === null || val instanceof Date) return;
                            let txt = (typeof val === 'object' && typeof val.text === 'function') ? val.text() : String(val);
                            let newTxt = replaceText(txt);
                            if (txt !== newTxt) cell.value(newTxt);
                        });
                    });
                });
                const blob = await wb.outputAsync();
                downloadBlob(blob, currentFileName.replace('anonymized_', 'deanon_'));
            } else if (currentFileType === 'csv') {
                const wb = XLSX.read(fileData, { type: 'array' });
                wb.SheetNames.forEach(sheetName => {
                    const sheet = wb.Sheets[sheetName];
                    if(!sheet['!ref']) return;
                    const range = XLSX.utils.decode_range(sheet['!ref']);
                    for (let R = range.s.r; R <= range.e.r; ++R) {
                        for (let C = range.s.c; C <= range.e.c; ++C) {
                            const cell = sheet[XLSX.utils.encode_cell({ c: C, r: R })];
                            if (cell && cell.t === 's') {
                                let txt = String(cell.v);
                                let newTxt = replaceText(txt);
                                if (txt !== newTxt) { cell.v = newTxt; if(cell.w) delete cell.w; }
                            }
                        }
                    }
                });
                const wbout = XLSX.write(wb, { bookType: 'csv', type: 'array' });
                downloadBlob(new Blob([wbout], { type: 'text/csv' }), currentFileName.replace('anonymized_', 'deanon_'));
            } else if (currentFileType === 'docx') {
                const arrayBuffer = await uploadedFiles[0].arrayBuffer();
                
                try {
                    const zip = new PizZip(arrayBuffer);
                    const doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                        delimiters: { start: '{{', end: '}}' },
                        nullGetter(part) {
                            if (!part.module) return "{{" + part.value + "}}";
                            return "";
                        }
                    });

                    let templateData = {};
                    for (const key in finalMapping) {
                        const cleanKey = key.replace('{{', '').replace('}}', '');
                        templateData[cleanKey] = finalMapping[key];
                    }

                    doc.render(templateData);
                    const outBlob = doc.getZip().generate({
                        type: "blob",
                        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    });
                    downloadBlob(outBlob, currentFileName.replace('anonymized_', 'deanon_'));
                } catch (error) {
                    alert("Nepodařilo se zachovat formátování: " + (error.message || error));
                    console.warn("DocxTemplater selhal, padáme zpět na HTML režim:", error);
                    const result = await mammoth.convertToHtml({ arrayBuffer: await uploadedFiles[0].arrayBuffer() });
                    let html = result.value;
                    if (!html || html.trim() === '') {
                        alert('Chyba: Dokument je prázdný.');
                        hideLoader(); return;
                    }
                    html = replaceText(html);
                    const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
                    const blob = htmlDocx.asBlob(wrappedHtml);
                    downloadBlob(blob, currentFileName.replace('anonymized_', 'deanon_'));
                }
            } else if (currentFileType === 'pdf') {
                let outText = replaceText(currentTextContent);
                const html = outText.replace(/\n/g, '<br>');
                const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
                const blob = htmlDocx.asBlob(wrappedHtml);
                downloadBlob(blob, currentFileName.replace('anonymized_', 'deanon_').replace('.pdf', '.docx'));
            } else {
                let outText = replaceText(currentTextContent);
                downloadBlob(new Blob([outText], { type: 'text/plain;charset=utf-8' }), currentFileName.replace('anonymized_', 'deanon_'));
            }

            hideLoader();
            alert('De-anonymizace proběhla úspěšně! Vaše originální data jsou zpět.');
            resetApp();

        } catch (err) {
            console.error(err); alert('Chyba při De-anonymizaci.'); hideLoader();
        }
    }, 100);
});
