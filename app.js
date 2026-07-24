// Thresholds por año (UMAs Desarrollo Inmobiliario)
const THRESHOLDS = {
    '2024': 871274.25,
    '2025': 907948.50,
    '2026': 941412.75
};

const BASE_YEAR = 2024;
const TOTAL_YEARS = 3;
const TOTAL_MONTHS = TOTAL_YEARS * 12; // 36 months
const SHORT_MONTHS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

// Global Client Profiles for KYC
let clientProfiles = {};


// KYC Required Fields
const KYC_REQUIRED_FISICA = ['kyc-nombre', 'kyc-apaterno', 'kyc-rfc', 'kyc-curp', 'kyc-fecha', 'kyc-nacionalidad', 'kyc-actividad', 'kyc-telefono', 'kyc-correo', 'kyc-pais-dom', 'kyc-estado', 'kyc-municipio', 'kyc-cp', 'kyc-colonia', 'kyc-calle', 'kyc-numext', 'kyc-tipo-inmueble', 'kyc-valor', 'kyc-folio', 'kyc-dimensiones'];
const KYC_REQUIRED_MORAL = ['kyc-nombre', 'kyc-rfc', 'kyc-fecha', 'kyc-nacionalidad', 'kyc-actividad', 'kyc-telefono', 'kyc-correo', 'kyc-pais-dom', 'kyc-estado', 'kyc-municipio', 'kyc-cp', 'kyc-colonia', 'kyc-calle', 'kyc-numext', 'kyc-tipo-inmueble', 'kyc-valor', 'kyc-folio', 'kyc-dimensiones'];

// Build flat array of 36 month names for UI
const timelineMonths = [];
for (let y = 0; y < TOTAL_YEARS; y++) {
    const year = BASE_YEAR + y;
    for (let m = 0; m < 12; m++) {
        timelineMonths.push({
            name: `${SHORT_MONTHS[m]} ${year.toString().slice(-2)}`,
            year: year,
            monthStr: SHORT_MONTHS[m],
            globalIdx: y * 12 + m
        });
    }
}

document.addEventListener('DOMContentLoaded', () => {
    
    // UI Elements
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const loader = document.getElementById('loader');
    
    const dashboardContainer = document.getElementById('dashboard-container');
    const resultsContainer = document.getElementById('results-container');
    
    const pivotBody = document.getElementById('pivot-body');
    const monthsHeaderRow = document.getElementById('months-header');
    
    const dashClients = document.getElementById('dash-clients');
    const dashNotices = document.getElementById('dash-notices');
    const calendarGrid = document.getElementById('calendar-grid');

    // Modals
    const noticeModal = document.getElementById('notice-modal');
    const modalClose = document.getElementById('modal-close');
    const modalTitle = document.getElementById('modal-title');
    const modalList = document.getElementById('modal-list');

    const kycModal = document.getElementById('kyc-modal');
    const kycClose = document.getElementById('kyc-close');
    const kycForm = document.getElementById('kyc-form');
    const kycSaveBtn = document.getElementById('kyc-save-btn');
    const kycTitle = document.getElementById('kyc-title');
    const kycStatusBadge = document.getElementById('kyc-status-badge');
    const kycUploadArea = document.getElementById('kyc-upload-area');
    const kycFileInput = document.getElementById('kyc-file-input');

    const btnOpenOwnerModal = document.getElementById('btn-open-owner-sidebar'); // Still uses the same ID on the dashboard button
    const ownerModal = document.getElementById('owner-modal');
    const closeOwnerModal = document.getElementById('close-owner-modal');
    
    const ownerExcelArea = document.getElementById('owner-excel-area');
    const ownerFileExcel = document.getElementById('owner-file-excel');
    const ownerExcelListContainer = document.getElementById('owner-excel-list-container');
    
    const globalPdfArea = document.getElementById('global-pdf-area');
    const globalFilePdf = document.getElementById('global-file-pdf');
    const pdfListContainer = document.getElementById('pdf-list-container');

    let currentKycClient = null;
    let globalRawData = {};
    let mainExcelRows = null;
    let ownerExcelRows = null;
    let pdfExtractedData = {}; // Stores {originAccount, paymentMethod} keyed by deposit signature
    
    let attachedOwnerExcel = null;
    let attachedPdfs = []; // { name, text }


    // Build Table Headers
    timelineMonths.forEach(m => {
        const th = document.createElement('th');
        th.textContent = m.monthStr;
        monthsHeaderRow.appendChild(th);
    });

    // Main Drag & Drop
    uploadArea.addEventListener('click', () => fileInput.click());
    setupDragAndDrop(uploadArea, (files) => {
        if (files.length) handleMainFile(files[0]);
    });
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) handleMainFile(e.target.files[0]);
    });

    // KYC Drag & Drop
    kycUploadArea.addEventListener('click', () => kycFileInput.click());
    setupDragAndDrop(kycUploadArea, (files) => {
        if (files.length) handleKycFile(files[0]);
    });
    kycFileInput.addEventListener('change', (e) => {
        if (e.target.files.length) handleKycFile(e.target.files[0]);
    });

    // Owner Modal Toggle (REMOVED: Upload is now directly on dashboard)
    
    // Owner Excel
    if (ownerExcelArea) {
        ownerExcelArea.addEventListener('click', () => ownerFileExcel.click());
        setupDragAndDrop(ownerExcelArea, (files) => {
            if (files.length) handleAportacionesFile(files[0]);
        });
        ownerFileExcel.addEventListener('change', (e) => {
            if (e.target.files.length) handleAportacionesFile(e.target.files[0]);
        });
    }

    // Global PDFs
    if (globalPdfArea) {
        globalPdfArea.addEventListener('click', () => globalFilePdf.click());
        setupDragAndDrop(globalPdfArea, (files) => {
            if (files.length) handlePdfFiles(files);
        });
        globalFilePdf.addEventListener('change', (e) => {
            if (e.target.files.length) handlePdfFiles(e.target.files);
        });
    }

    // Modals Close Events
    [modalClose, noticeModal].forEach(el => el.addEventListener('click', (e) => {
        if (e.target === el) noticeModal.classList.add('hidden');
    }));
    [kycClose, kycModal].forEach(el => el.addEventListener('click', (e) => {
        if (e.target === el) kycModal.classList.add('hidden');
    }));

    // KYC Form Logic
    document.getElementById('kyc-tipo').addEventListener('change', (e) => {
        const isMoral = e.target.value === 'moral' || e.target.value === 'fideicomiso';
        document.getElementById('group-apaterno').style.display = isMoral ? 'none' : 'flex';
        document.getElementById('group-amaterno').style.display = isMoral ? 'none' : 'flex';
        document.getElementById('group-curp').style.display = isMoral ? 'none' : 'flex';
    });

    kycSaveBtn.addEventListener('click', () => {
        if (!currentKycClient) return;
        
        // Save form to profile
        const data = {
            tipo: document.getElementById('kyc-tipo').value
        };
        const inputs = kycForm.querySelectorAll('input, select');
        inputs.forEach(input => {
            if (input.id) {
                data[input.id] = input.value.trim();
            }
        });

        clientProfiles[currentKycClient] = data;
        
        kycModal.classList.add('hidden');
        updateTableKycIcons();
    });

    function setupDragAndDrop(element, callback) {
        element.addEventListener('dragover', (e) => { e.preventDefault(); element.classList.add('dragover'); });
        element.addEventListener('dragleave', () => { element.classList.remove('dragover'); });
        element.addEventListener('drop', (e) => {
            e.preventDefault();
            element.classList.remove('dragover');
            callback(e.dataTransfer.files);
        });
    }

    // -- Main File Processing --
    function getGlobalMonthIdx(dateObj) {
        if (isNaN(dateObj.getTime())) return -1;
        const y = dateObj.getFullYear();
        const m = dateObj.getMonth();
        if (y < BASE_YEAR || y >= BASE_YEAR + TOTAL_YEARS) return -1;
        return (y - BASE_YEAR) * 12 + m;
    }

    function handleMainFile(file) {
        loader.classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            try {
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
                mainExcelRows = rows;
                extractRawData(rows, false);
            } catch (error) {
                console.error("Error reading file:", error);
                alert("Hubo un error al leer el archivo. Asegúrate de que sea un Excel válido.");
                resetUI();
            } finally {
                fileInput.value = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleAportacionesFile(file) {
        attachedOwnerExcel = { name: file.name };
        renderOwnerExcelList();
        
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            try {
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
                ownerExcelRows = rows;
                extractRawData(rows, true);
                alert("Excel del Dueño cargado correctamente.");
            } catch (error) {
                console.error("Error reading aportaciones file:", error);
                alert("Error al procesar archivo de aportaciones.");
            } finally {
                ownerFileExcel.value = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function renderOwnerExcelList() {
        ownerExcelListContainer.innerHTML = '';
        if (attachedOwnerExcel) {
            const div = document.createElement('div');
            div.classList.add('pdf-item');
            div.innerHTML = `<span>📊 ${attachedOwnerExcel.name}</span><button type="button" title="Eliminar">&times;</button>`;
            div.querySelector('button').addEventListener('click', () => {
                attachedOwnerExcel = null;
                delete globalRawData["Aportación Dueño"];
                renderOwnerExcelList();
                saveState();
                if (Object.keys(globalRawData).length > 0) processNoticesAndRender(globalRawData);
            });
            ownerExcelListContainer.appendChild(div);
        }
    }


    async function handlePdfFiles(files) {
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const pdfItem = { name: file.name, text: "" };
            attachedPdfs.push(pdfItem);
            
            // Parse PDF
            try {
                const arrayBuffer = await file.arrayBuffer();
                const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
                for (let p = 1; p <= pdf.numPages; p++) {
                    const page = await pdf.getPage(p);
                    const content = await page.getTextContent();
                    pdfItem.text += content.items.map(item => item.str).join(" ") + " \n ";
                }
            } catch (e) {
                console.error("Error parsing PDF:", file.name, e);
            }
        }
        renderPdfList();
        saveState();
        if (Object.keys(globalRawData).length > 0) processNoticesAndRender(globalRawData);
    }


    function renderPdfList() {
        pdfListContainer.innerHTML = '';
        attachedPdfs.forEach((pdfObj, index) => {
            const div = document.createElement('div');
            div.classList.add('pdf-item');
            div.innerHTML = `<span>📄 ${pdfObj.name}</span><button type="button" title="Eliminar">&times;</button>`;
            div.querySelector('button').addEventListener('click', () => {
                attachedPdfs.splice(index, 1);
                renderPdfList();
                saveState();
                if (Object.keys(globalRawData).length > 0) {
                    processNoticesAndRender(globalRawData);
                }
            });
            pdfListContainer.appendChild(div);
        });
        globalFilePdf.value = '';
    }

    function extractRawData(rows, isAportacion) {
        let headerRowIndex = -1;
        let fechaIdx = -1, abonoIdx = -1, nombreIdx = -1, refIdx = -1, notasIdx = -1;

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            const hasFecha = row.findIndex(c => typeof c === 'string' && c.toLowerCase().includes('fecha')) !== -1;
            const hasAbono = row.findIndex(c => typeof c === 'string' && (c.toLowerCase().includes('abono') || c.toLowerCase().includes('abo') || c.toLowerCase().includes('depósito') || c.toLowerCase().includes('monto'))) !== -1;
            
            if (hasFecha && hasAbono) {
                headerRowIndex = i;
                fechaIdx = row.findIndex(c => typeof c === 'string' && c.toLowerCase().includes('fecha'));
                abonoIdx = row.findIndex(c => typeof c === 'string' && (c.toLowerCase().includes('abono') || c.toLowerCase().includes('abo') || c.toLowerCase().includes('depósito') || c.toLowerCase().includes('monto')));
                nombreIdx = row.findIndex(c => typeof c === 'string' && (c.toLowerCase().includes('nombre') || c.toLowerCase().includes('concepto') || c.toLowerCase().includes('cliente')));
                refIdx = row.findIndex(c => typeof c === 'string' && c.toLowerCase().includes('referencia'));
                notasIdx = row.findIndex(c => typeof c === 'string' && (c.toLowerCase().includes('nota') || c.toLowerCase().includes('observacion')));
                break;
            }
        }

        if (headerRowIndex === -1 || fechaIdx === -1 || abonoIdx === -1) {
            alert("No se pudieron encontrar las columnas 'Fecha' y 'Monto' / 'Abonos'.");
            resetUI();
            return;
        }

        if (!isAportacion) globalRawData = {};
        const monthMap = { 'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11 };

        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const dateVal = row[fechaIdx];
            const depositVal = parseFloat(row[abonoIdx]);
            
            let nameVal = "Desconocido";
            if (isAportacion) {
                nameVal = "Aportación Dueño";
            } else if (nombreIdx !== -1 && row[nombreIdx]) {
                nameVal = String(row[nombreIdx]).trim();
            }

            if (!isNaN(depositVal) && depositVal > 0 && dateVal) {
                let exactDate = null;
                let month = -1;
                
                if (dateVal instanceof Date) {
                    exactDate = dateVal;
                } else if (typeof dateVal === 'number') {
                    exactDate = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
                } else if (typeof dateVal === 'string') {
                    if (dateVal.match(/^\d{4}-\d{2}-\d{2}/)) {
                        // Always parse as local date to avoid UTC timezone shifting
                        const dp = dateVal.substring(0, 10).split('-');
                        exactDate = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));

                    } else {
                        const lowerDate = dateVal.toLowerCase();
                        for (const [mName, mIdx] of Object.entries(monthMap)) {
                            if (lowerDate.includes(mName)) {
                                month = mIdx;
                                break;
                            }
                        }
                        
                        const parts = dateVal.split(/[-/]/);
                        if (parts.length >= 2) {
                            let y = parseInt(parts[2]);
                            if (y < 2000) y += 2000;
                            let m = month !== -1 ? month : parseInt(parts[1]) - 1;
                            let d = parseInt(parts[0]);
                            if (!isNaN(y) && !isNaN(m) && !isNaN(d)) {
                                exactDate = new Date(y, m, d);
                            }
                        }
                    }
                }

                if (exactDate && !isNaN(exactDate.getTime())) {
                    const globalIdx = getGlobalMonthIdx(exactDate);
                    if (globalIdx !== -1) {
                        if (!globalRawData[nameVal]) globalRawData[nameVal] = [];
                        
                        let displayDate = exactDate.toLocaleDateString('es-MX');
                        if (notasIdx !== -1 && typeof row[notasIdx] === 'string' && row[notasIdx].toLowerCase().includes('fecha')) {
                            displayDate = row[notasIdx];
                        }

                        globalRawData[nameVal].push({
                            amount: depositVal,
                            exactDate: exactDate,
                            displayDate: displayDate,
                            ref: refIdx !== -1 && row[refIdx] ? String(row[refIdx]) : '',
                            globalIdx: globalIdx
                        });
                    }
                }
            }
        }

        if (Object.keys(globalRawData).length === 0) {
            alert("No se encontraron depósitos válidos dentro del rango 2024-2026.");
            resetUI();
            return;
        }

        for (const name in globalRawData) {
            globalRawData[name].sort((a, b) => a.exactDate.getTime() - b.exactDate.getTime());
            // Initialize empty profile if doesn't exist
            if (!clientProfiles[name]) {
                clientProfiles[name] = { tipo: 'fisica' }; // Default
            }
        }

        processNoticesAndRender(globalRawData);
    }

    function processNoticesAndRender(rawClientData) {
        let totalNotices = 0;
        let noticesByMonth = {};
        let ownerNoticesByMonth = {};
        
        const tableData = {};
        let ownerTotal = 0;
        
        const allPdfText = attachedPdfs.map(p => p.text).join("  ");

        for (const [name, deposits] of Object.entries(rawClientData)) {
            tableData[name] = Array.from({length: TOTAL_MONTHS}, () => ({ deposits: [], hasNotice: false, endSum: 0, appliedThreshold: 0 }));
            let rollingWindow = [];
            
            if (name === "Aportación Dueño") {
                deposits.forEach(d => ownerTotal += d.amount);
            }
            
            deposits.forEach(d => {
                const depKey = `${name}_${d.displayDate}_${d.amount}`;
                
                // 1. Check if we already have it in the persistent dictionary
                if (pdfExtractedData[depKey]) {
                    d.foundInPdf = true;
                    d.originAccount = pdfExtractedData[depKey].originAccount;
                    d.paymentMethod = pdfExtractedData[depKey].paymentMethod;
                } 
                // 2. Or, if there are PDFs currently loaded, run the regex search
                else if (allPdfText.length > 0) {
                    const amountStr = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(d.amount);
                    const amountStrMX = new Intl.NumberFormat('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(d.amount);
                    
                    const amountRegex = new RegExp(`(?:\\$)?\\s*(?:${amountStr.replace(/,/g, '[,\\s]*')}|${amountStrMX.replace(/,/g, '[,\\s]*')})`, 'i');
                    
                    if (amountRegex.test(allPdfText)) {
                        d.foundInPdf = true;
                        
                        const idx = allPdfText.indexOf(amountStr) !== -1 ? allPdfText.indexOf(amountStr) : allPdfText.indexOf(amountStrMX);
                        let originAccount = null;
                        let paymentMethod = null;
                        
                        if (idx !== -1) {
                            const context = allPdfText.substring(Math.max(0, idx - 200), idx + 200);
                            const accountMatch = context.match(/\b\d{10,20}\b/);
                            if (accountMatch) originAccount = accountMatch[0];
                            
                            const ctxUpper = context.toUpperCase();
                            if (ctxUpper.includes('CHEQUE')) paymentMethod = 'Cheque';
                            else if (ctxUpper.includes('SPEI') || ctxUpper.includes('TRANSFERENCIA')) paymentMethod = 'Transferencia';
                            else if (ctxUpper.includes('EFECTIVO')) paymentMethod = 'Efectivo';
                            else if (ctxUpper.includes('TRASPASO')) paymentMethod = 'Traspaso';
                        }
                        
                        d.originAccount = originAccount;
                        d.paymentMethod = paymentMethod;
                        
                        // Save it to dictionary so we don't need the PDF again!
                        pdfExtractedData[depKey] = { originAccount, paymentMethod };
                    }
                }
            });
            
            for (const dep of deposits) {
                tableData[name][dep.globalIdx].deposits.push(dep);
                rollingWindow.push(dep);
                
                while (rollingWindow.length > 0) {
                    const first = rollingWindow[0];
                    const monthsDiff = (dep.exactDate.getFullYear() - first.exactDate.getFullYear()) * 12 + (dep.exactDate.getMonth() - first.exactDate.getMonth());
                    if (monthsDiff >= 6) rollingWindow.shift();
                    else break;
                }
                
                const currentSum = rollingWindow.reduce((s, d) => s + d.amount, 0);
                const depositYear = dep.exactDate.getFullYear().toString();
                const currentThreshold = THRESHOLDS[depositYear] || THRESHOLDS['2024'];

                tableData[name][dep.globalIdx].endSum = currentSum;
                tableData[name][dep.globalIdx].appliedThreshold = currentThreshold;
                
                if (currentSum > currentThreshold) {
                    tableData[name][dep.globalIdx].hasNotice = true;
                    if (name !== "Aportación Dueño") {
                        totalNotices++;
                    }
                    
                    const targetDict = (name === "Aportación Dueño") ? ownerNoticesByMonth : noticesByMonth;
                    if (!targetDict[dep.globalIdx]) targetDict[dep.globalIdx] = [];
                    
                    let nextMonth = dep.exactDate.getMonth() + 1;
                    let nextYear = dep.exactDate.getFullYear();
                    if (nextMonth > 11) { nextMonth = 0; nextYear++; }
                    
                    targetDict[dep.globalIdx].push({
                        clientName: name,
                        amount: dep.amount,
                        sum: currentSum,
                        exactDate: dep.displayDate,
                        limitDate: new Date(nextYear, nextMonth, 17).toLocaleDateString('es-MX', {day: 'numeric', month: 'long', year: 'numeric'}),
                        contributingDeposits: [...rollingWindow]
                    });
                    
                    rollingWindow = []; // Reset accumulator
                }
            }
        }

        // Render Dashboard
        const clientKeys = Object.keys(rawClientData).filter(name => name !== "Aportación Dueño");
        dashClients.textContent = clientKeys.length;
        dashNotices.textContent = totalNotices;
        const ownerTotalEl = document.getElementById('dash-owner-total');
        if (ownerTotalEl) ownerTotalEl.textContent = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(ownerTotal);
        
        calendarGrid.innerHTML = '';
        const ownerCalendarGridEl = document.getElementById('owner-calendar-grid');
        if (ownerCalendarGridEl) ownerCalendarGridEl.innerHTML = '';

        timelineMonths.forEach((m, globalIdx) => {
            const count = noticesByMonth[globalIdx] ? noticesByMonth[globalIdx].length : 0;
            const ownerCount = ownerNoticesByMonth[globalIdx] ? ownerNoticesByMonth[globalIdx].length : 0;
            
            const box = document.createElement('div');
            box.classList.add('month-box');
            
            if (count > 0 || ownerCount > 0) {
                box.classList.add('active', 'interactive');
                box.addEventListener('click', () => openDualNoticeModal(m.name, noticesByMonth[globalIdx] || [], ownerNoticesByMonth[globalIdx] || []));
            }
            
            if (count > 0 && ownerCount > 0) {
                box.innerHTML = `
                    <div class="month-name">${m.name}</div>
                    <div style="display:flex; justify-content:space-around; width:100%; margin-top:0.5rem;">
                        <div style="text-align:center;"><div class="notice-count" style="font-size:1.2rem;">${count}</div><div style="font-size:0.6rem;">Clientes</div></div>
                        <div style="border-left:1px solid rgba(255,255,255,0.2);"></div>
                        <div style="text-align:center;"><div class="notice-count" style="font-size:1.2rem; color:var(--success);">${ownerCount}</div><div style="font-size:0.6rem;">Dueño</div></div>
                    </div>
                `;
            } else if (count > 0) {
                box.innerHTML = `<div class="month-name">${m.name}</div><div class="notice-count">${count} <span style="font-size:0.8rem; font-weight:normal">Avisos</span></div>`;
            } else if (ownerCount > 0) {
                box.innerHTML = `<div class="month-name">${m.name}</div><div class="notice-count" style="color:var(--success);">${ownerCount} <span style="font-size:0.8rem; font-weight:normal">Dueño</span></div>`;
            } else {
                box.innerHTML = `<div class="month-name">${m.name}</div><div class="notice-count">0</div>`;
            }
            
            calendarGrid.appendChild(box);
        });

        renderTable(tableData);
        loader.classList.add('hidden');
        dashboardContainer.classList.remove('hidden');
        resultsContainer.classList.remove('hidden');
        
        // Save state AFTER all flags (like foundInPdf) have been populated
        saveState();
    }

    function renderTable(tableData) {
        pivotBody.innerHTML = '';
        const formatCurrency = (val) => val === 0 ? '-' : new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(val);

        let ownerRowTr = null;

        for (const [name, monthsData] of Object.entries(tableData)) {
            const isOwner = name === "Aportación Dueño";
            const tr = document.createElement('tr');
            
            if (isOwner) {
                tr.style.borderTop = '3px solid rgba(16, 185, 129, 0.5)';
                tr.style.backgroundColor = 'rgba(16, 185, 129, 0.05)';
            }
            
            // Name cell with KYC status
            const tdName = document.createElement('td');
            const nameContainer = document.createElement('div');
            nameContainer.classList.add('client-name-cell');
            
            if (!isOwner) {
                const statusIcon = document.createElement('span');
                statusIcon.classList.add('status-icon');
                statusIcon.id = `status-${name.replace(/[^a-zA-Z0-9]/g, '-')}`;
                nameContainer.appendChild(statusIcon);
            }
            
            const nameText = document.createElement('span');
            nameText.textContent = name;
            if (isOwner) nameText.style.color = 'var(--success)';
            
            nameContainer.appendChild(nameText);
            tdName.appendChild(nameContainer);
            
            if (!isOwner) nameContainer.addEventListener('click', () => openKycModal(name));
            tr.appendChild(tdName);

            let clientHistoricalTotal = 0;
            monthsData.forEach((monthData) => {
                const monthTotal = monthData.deposits.reduce((sum, dep) => sum + dep.amount, 0);
                clientHistoricalTotal += monthTotal;
                
                const td = document.createElement('td');
                td.classList.add('currency');
                td.textContent = formatCurrency(monthTotal);

                if (monthData.deposits.length > 0) {
                    let tooltip = "DEPÓSITOS DEL MES:\n" + monthData.deposits.map(d => `${d.displayDate} - ${formatCurrency(d.amount)} ${d.ref ? `(${d.ref})` : ''}`).join('\n') + 
                        "\n\n-- PROGRESO LFPIORPI --\n" +
                        `Acumulado (últimos 6m): ${formatCurrency(monthData.endSum)}\n` +
                        `Umbral aplicable: ${formatCurrency(monthData.appliedThreshold)} (${((monthData.endSum / monthData.appliedThreshold) * 100).toFixed(1)}%)`;
                    td.title = tooltip;
                    td.style.cursor = 'help';
                    td.style.textDecoration = 'underline dotted rgba(255,255,255,0.3)';
                    td.style.textUnderlineOffset = '4px';
                }

                if (monthData.hasNotice) td.classList.add('threshold-exceeded');
                tr.appendChild(td);
            });

            const tdTotal = document.createElement('td');
            tdTotal.classList.add('currency');
            tdTotal.style.fontWeight = '600';
            tdTotal.textContent = formatCurrency(clientHistoricalTotal);
            if (isOwner) tdTotal.style.color = 'var(--success)';
            tr.appendChild(tdTotal);
            
            if (isOwner) {
                ownerRowTr = tr;
            } else {
                pivotBody.appendChild(tr);
            }
        }
        
        if (ownerRowTr) {
            pivotBody.appendChild(ownerRowTr);
        }
        
        updateTableKycIcons();
    }

    // -- KYC Feature Logic --

    function checkKycCompleteness(profile) {
        if (!profile) return { isComplete: false, missing: ['No hay datos guardados'] };
        
        const reqFields = (profile.tipo === 'moral' || profile.tipo === 'fideicomiso') ? KYC_REQUIRED_MORAL : KYC_REQUIRED_FISICA;
        const missing = [];
        
        reqFields.forEach(fieldId => {
            const val = profile[fieldId];
            if (!val || val.trim() === '') {
                // Get nice label
                const input = document.getElementById(fieldId);
                const label = input ? input.previousElementSibling.textContent : fieldId;
                missing.push(label);
            }
        });
        
        return { isComplete: missing.length === 0, missing };
    }

    function updateTableKycIcons() {
        saveState();
        for (const [name, profile] of Object.entries(clientProfiles)) {
            const safeId = `status-${name.replace(/[^a-zA-Z0-9]/g, '-')}`;
            const iconEl = document.getElementById(safeId);
            if (!iconEl) continue;

            const { isComplete, missing } = checkKycCompleteness(profile);
            
            if (isComplete) {
                iconEl.textContent = '✅';
                iconEl.removeAttribute('data-missing');
            } else {
                iconEl.textContent = '⚠️';
                iconEl.setAttribute('data-missing', "Faltan Datos:\n- " + missing.join("\n- "));
            }
        }
    }

    function openKycModal(clientName) {
        currentKycClient = clientName;
        kycTitle.textContent = `Expediente: ${clientName}`;
        
        // Clear form
        kycForm.reset();
        document.querySelectorAll('.not-found-text').forEach(el => el.remove());
        
        document.getElementById('kyc-tipo').value = 'fisica';
        document.getElementById('group-apaterno').style.display = 'flex';
        document.getElementById('group-amaterno').style.display = 'flex';
        document.getElementById('group-curp').style.display = 'flex';

        // Load existing data
        const profile = clientProfiles[clientName];
        if (profile) {
            for (const key in profile) {
                const input = document.getElementById(key);
                if (input) input.value = profile[key];
            }
            if (profile.tipo === 'moral' || profile.tipo === 'fideicomiso') {
                document.getElementById('group-apaterno').style.display = 'none';
                document.getElementById('group-amaterno').style.display = 'none';
                document.getElementById('group-curp').style.display = 'none';
            }
        } else {
            // Pre-fill name
            document.getElementById('kyc-nombre').value = clientName;
        }

        updateKycModalBadge(profile);
        kycModal.classList.remove('hidden');
    }

    function updateKycModalBadge(profile) {
        const { isComplete } = checkKycCompleteness(profile);
        if (isComplete) {
            kycStatusBadge.textContent = 'Completo';
            kycStatusBadge.className = 'badge success';
        } else {
            kycStatusBadge.textContent = 'Incompleto';
            kycStatusBadge.className = 'badge warning';
        }
    }

    // Parse Uploaded KYC File (Word/Excel)
    function handleKycFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        
        if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm') {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    // Basic heuristic search across sheets
                    parseExcelKyc(workbook);
                } catch(err) {
                    console.error(err);
                    alert("Error al leer el Excel KYC.");
                }
            };
            reader.readAsArrayBuffer(file);
        } else if (ext === 'docx') {
            const reader = new FileReader();
            reader.onload = (e) => {
                if (window.mammoth) {
                    mammoth.extractRawText({ arrayBuffer: e.target.result })
                        .then(result => parseWordKyc(result.value))
                        .catch(err => alert("Error al leer el Word KYC."));
                } else {
                    alert("Mammoth.js no cargado. No se puede leer el Word.");
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            alert("Formato no soportado. Usa Excel (.xlsx, .xls, .xlsm) o Word (.docx).");
        }
    }

    function parseExcelKyc(workbook) {
        let isSSPLD = false;
        
        // Scan all sheets
        for (let s = 0; s < Math.min(3, workbook.SheetNames.length); s++) {
            const ws = workbook.Sheets[workbook.SheetNames[s]];
            const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
            for(let i = 0; i < Math.min(30, json.length); i++) {
                if(json[i] && json[i].join('').toLowerCase().includes('identificaci') && json[i].join('').toLowerCase().includes('persona')) {
                    isSSPLD = true;
                    break;
                }
            }
        }

        if (isSSPLD) {
            let kycData = {
                tipo: 'fisica',
                fields: {},
                missingInExcel: {}
            };

            const extractField = (headers, data, keyword, inputId) => {
                const idx = headers.findIndex(h => h.includes(keyword));
                if (idx !== -1) {
                    let val = data[idx];
                    if (val !== undefined && val !== null && String(val).trim() !== '') {
                        kycData.fields[inputId] = String(val).trim();
                    } else {
                        kycData.missingInExcel[inputId] = true;
                    }
                }
            };

            const extractDate = (headers, data, keyword, inputId) => {
                const idx = headers.findIndex(h => h.includes(keyword));
                if (idx !== -1) {
                    let val = data[idx];
                    if (val !== undefined && val !== null && String(val).trim() !== '') {
                        if (val instanceof Date) {
                            kycData.fields[inputId] = val.toISOString().split('T')[0];
                        } else if (typeof val === 'number') {
                            const exactDate = new Date(Math.round((val - 25569) * 86400 * 1000));
                            kycData.fields[inputId] = exactDate.toISOString().split('T')[0];
                        } else {
                            // Intenta procesar YYYY-MM-DD HH:MM:SS de string si aplica
                            let strVal = String(val).trim();
                            if(strVal.includes(' ')) strVal = strVal.split(' ')[0];
                            kycData.fields[inputId] = strVal;
                        }
                    } else {
                        kycData.missingInExcel[inputId] = true;
                    }
                }
            };

            for (let s = 0; s < workbook.SheetNames.length; s++) {
                const ws = workbook.Sheets[workbook.SheetNames[s]];
                const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

                for(let i = 0; i < json.length; i++) {
                    const row = json[i];
                    if (!row) continue;
                    const rowStr = row.join('').toLowerCase();
                    
                    // Persona Física
                    if (rowStr.includes('nombre(s)') && rowStr.includes('apellido paterno')) {
                        const headers = row.map(h => String(h).trim().toLowerCase());
                        const nIdx = headers.findIndex(h => h.includes('nombre(s)'));
                        
                        let data = json[i+1] || [];
                        if (!data[nIdx] && json[i+2]) data = json[i+2]; 
                        
                        extractField(headers, data, 'nombre(s)', 'kyc-nombre');
                        extractField(headers, data, 'apellido paterno', 'kyc-apaterno');
                        extractField(headers, data, 'apellido materno', 'kyc-amaterno');
                        extractField(headers, data, 'rfc', 'kyc-rfc');
                        extractField(headers, data, 'curp', 'kyc-curp');
                        extractDate(headers, data, 'fecha nacimiento', 'kyc-fecha');
                        extractField(headers, data, 'nacionalidad', 'kyc-nacionalidad');
                        extractField(headers, data, 'actividad', 'kyc-actividad');
                    }

                    // Persona Moral
                    if (rowStr.includes('denominación o razón social') || rowStr.includes('denominación o raz')) {
                        const headers = row.map(h => String(h).trim().toLowerCase());
                        const dIdx = headers.findIndex(h => h.includes('denominación o raz'));
                        
                        let data = json[i+1] || [];
                        if (!data[dIdx] && json[i+2]) data = json[i+2];
                        
                        if (dIdx !== -1) {
                            kycData.tipo = 'moral';
                            extractField(headers, data, 'denominación o raz', 'kyc-nombre');
                            extractField(headers, data, 'rfc', 'kyc-rfc');
                            extractDate(headers, data, 'fecha de constitu', 'kyc-fecha');
                            extractField(headers, data, 'nacionalidad', 'kyc-nacionalidad');
                            extractField(headers, data, 'giro mercantil', 'kyc-actividad');
                        }
                    }

                    // Domicilio Nacional
                    if (rowStr.includes('código postal') && rowStr.includes('estado') && rowStr.includes('colonia') && !rowStr.includes('tipo de bien')) {
                        const headers = row.map(h => String(h).trim().toLowerCase());
                        const cpIdx = headers.findIndex(h => h.includes('código postal') || h.includes('código'));
                        
                        let data = json[i+1] || [];
                        if (!data[cpIdx] && json[i+2]) data = json[i+2];
                        
                        extractField(headers, data, 'código', 'kyc-cp');
                        extractField(headers, data, 'colonia', 'kyc-colonia');
                        extractField(headers, data, 'calle', 'kyc-calle');
                        extractField(headers, data, 'exterior', 'kyc-numext');
                        extractField(headers, data, 'interior', 'kyc-numint');
                        extractField(headers, data, 'estado', 'kyc-estado');
                        extractField(headers, data, 'municipio', 'kyc-municipio');
                    }

                    // Contacto
                    if (rowStr.includes('teléfono') && rowStr.includes('correo')) {
                        const headers = row.map(h => String(h).trim().toLowerCase());
                        const telIdx = headers.findIndex(h => h.includes('teléfono'));
                        
                        let data = json[i+1] || [];
                        if (!data[telIdx] && json[i+2]) data = json[i+2];
                        
                        extractField(headers, data, 'teléfono', 'kyc-telefono');
                        extractField(headers, data, 'correo', 'kyc-correo');
                    }

                    // Inmuebles
                    if (rowStr.includes('tipo de bien inmue') || rowStr.includes('valor pactado')) {
                        const headers = row.map(h => String(h).trim().toLowerCase());
                        const valIdx = headers.findIndex(h => h.includes('valor pactado'));
                        
                        let data = json[i+1] || [];
                        if (!data[valIdx] && json[i+2]) data = json[i+2];
                        
                        extractField(headers, data, 'tipo de bien', 'kyc-tipo-inmueble');
                        extractField(headers, data, 'valor pactado', 'kyc-valor');
                        extractField(headers, data, 'folio real', 'kyc-folio');
                        extractField(headers, data, 'dimensiones del te', 'kyc-dimensiones');
                    }
                }
            }

            if (Object.keys(kycData.fields).length > 0) {
                document.getElementById('kyc-tipo').value = kycData.tipo;
                document.getElementById('kyc-tipo').dispatchEvent(new Event('change'));
                
                document.querySelectorAll('.not-found-text').forEach(el => el.remove());

                const allInputs = [
                    'kyc-nombre', 'kyc-apaterno', 'kyc-amaterno', 'kyc-rfc', 'kyc-curp', 
                    'kyc-fecha', 'kyc-nacionalidad', 'kyc-actividad', 'kyc-cp', 'kyc-colonia', 
                    'kyc-calle', 'kyc-numext', 'kyc-numint', 'kyc-estado', 'kyc-municipio', 
                    'kyc-telefono', 'kyc-correo', 'kyc-tipo-inmueble', 'kyc-valor', 
                    'kyc-folio', 'kyc-dimensiones'
                ];

                allInputs.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) {
                        el.value = kycData.fields[id] || '';
                        
                        if (kycData.missingInExcel[id] && !kycData.fields[id]) {
                            const msg = document.createElement('small');
                            msg.className = 'not-found-text';
                            msg.textContent = 'No encontrado en el Excel';
                            el.parentNode.appendChild(msg);
                        }
                    }
                });

                alert("Se ha autollenado el expediente usando el formato Oficial SSPLD.");
                return;
            }
        }

        // HEURÍSTICA GENÉRICA (Fallback)
        let rfc = '', curp = '', cp = '', nombre = '', calle = '', tel = '', correo = '';
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' });
        for (let i = 0; i < json.length; i++) {
            for (let j = 0; j < json[i].length; j++) {
                let cell = String(json[i][j]).trim().toLowerCase();
                if (cell === 'rfc') rfc = json[i+1]?.[j] || '';
                if (cell.includes('curp')) curp = json[i+1]?.[j] || '';
                if (cell.includes('código postal')) cp = json[i+1]?.[j] || '';
                if (cell.includes('correo')) correo = json[i+1]?.[j] || '';
                if (cell.includes('teléfono')) tel = json[i+1]?.[j] || '';
                if (cell.includes('calle')) calle = json[i+1]?.[j] || '';
            }
        }

        if (rfc) document.getElementById('kyc-rfc').value = String(rfc).toUpperCase();
        if (curp) document.getElementById('kyc-curp').value = String(curp).toUpperCase();
        if (cp) document.getElementById('kyc-cp').value = cp;
        if (correo) document.getElementById('kyc-correo').value = correo;
        if (tel) document.getElementById('kyc-telefono').value = tel;
        if (calle) document.getElementById('kyc-calle').value = calle;

        alert("Se ha extraído información genérica del Excel. Por favor, revisa el expediente.");
    }

    function parseWordKyc(text) {
        // Encontrar todos los RFCs en el documento
        const rfcMatches = text.match(/[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}/g);
        
        if (rfcMatches && rfcMatches.length > 0) {
            // Si hay más de un RFC, probablemente el primero es el de la empresa o notario
            // Tomamos el segundo (o el último si solo hay dos) que suele ser el del cliente
            const rfc = rfcMatches.length > 1 ? rfcMatches[1] : rfcMatches[0];
            document.getElementById('kyc-rfc').value = rfc;
        }

        // CURP
        const curpMatch = text.match(/[A-Z][AEIOUX][A-Z]{2}\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])[HM](AS|BC|BS|CC|CL|CM|CS|CH|DF|DG|GT|GR|HG|JC|MC|MN|MS|NT|NL|OC|PL|QT|QR|SP|SL|SR|TC|TS|TL|VZ|YN|ZS|NE)[B-DF-HJ-NP-TV-Z]{3}[A-Z\d]\d/i);
        // Email
        const emailMatch = text.match(/[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}/);
        
        if (rfcMatch) document.getElementById('kyc-rfc').value = rfcMatch[0].toUpperCase();
        if (curpMatch) document.getElementById('kyc-curp').value = curpMatch[0].toUpperCase();
        if (emailMatch) document.getElementById('kyc-correo').value = emailMatch[0];

        alert("Se ha extraído información del Word mediante inteligencia. Por favor, revisa los campos llenos.");
    }

    // Modal Helpers
    function generateNoticeHtml(notice, formatCurrency) {
        const item = document.createElement('div');
        item.classList.add('notice-item');
        
        let depositsHtml = '<div style="margin-top:0.75rem; padding-left:1rem; border-left:2px solid var(--border-color);">';
        depositsHtml += '<p style="font-size:0.85rem; color:var(--text-muted); margin-bottom:0.25rem;">Historial de acumulación (últimos 6 meses):</p>';
        
        notice.contributingDeposits.forEach(d => {
            let accountHtml = '';
            
            // Payment method and origin account for EVERYONE
            if (d.originAccount || d.paymentMethod) {
                const parts = [];
                if (d.originAccount) parts.push(`Cuenta Origen: ${d.originAccount}`);
                if (d.paymentMethod) parts.push(d.paymentMethod);
                accountHtml = `<br><span style="font-size:0.75rem; color:var(--success);">${parts.join(' - ')}</span>`;
            } else {
                accountHtml = `<br><span style="font-size:0.75rem; color:var(--warning);">No encontrado en el estado de cuenta</span>`;
            }
            
            depositsHtml += `<div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:0.2rem;">
                <span>${d.displayDate}${accountHtml}</span>
                <span style="font-family:monospace;">${formatCurrency(d.amount)}</span>
            </div>`;
        });
        depositsHtml += '</div>';

        item.innerHTML = `<h4>${notice.clientName}</h4>
            <p><strong>Fecha que detonó:</strong> ${notice.exactDate}</p>
            <p><strong>Operación detonante:</strong> ${formatCurrency(notice.amount)}</p>
            <p><strong>Acumulado:</strong> ${formatCurrency(notice.sum)}</p>
            ${depositsHtml}
            <p style="color: var(--warning); margin-top: 1rem;"><strong>Límite SAT:</strong> 17 de ${notice.limitDate.split(' ')[2]} del ${notice.limitDate.split(' ')[4]}</p>`;
        
        return item;
    }

    function openNoticeModal(monthName, notices) {
        modalTitle.textContent = `Avisos de ${monthName}`;
        modalList.innerHTML = '';
        const formatCurrency = (val) => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(val);
        notices.forEach(notice => {
            modalList.appendChild(generateNoticeHtml(notice, formatCurrency));
        });
        noticeModal.classList.remove('hidden');
    }
    
    function openDualNoticeModal(monthName, clientNotices, ownerNotices) {
        modalTitle.textContent = `Avisos de ${monthName}`;
        modalList.innerHTML = `
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 1.5rem;">
                <div id="dual-client-list">
                    <h3 style="border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem; margin-bottom: 1rem;">Clientes</h3>
                </div>
                <div id="dual-owner-list">
                    <h3 style="border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem; margin-bottom: 1rem; color: var(--success);">Dueño de la Empresa</h3>
                </div>
            </div>
        `;
        
        const formatCurrency = (val) => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(val);
        
        const clientList = modalList.querySelector('#dual-client-list');
        clientNotices.forEach(notice => clientList.appendChild(generateNoticeHtml(notice, formatCurrency)));
        if(clientNotices.length === 0) clientList.innerHTML += `<p style="color:var(--text-muted)">Sin avisos.</p>`;
        
        const ownerList = modalList.querySelector('#dual-owner-list');
        ownerNotices.forEach(notice => ownerList.appendChild(generateNoticeHtml(notice, formatCurrency)));
        if(ownerNotices.length === 0) ownerList.innerHTML += `<p style="color:var(--text-muted)">Sin avisos.</p>`;

        noticeModal.classList.remove('hidden');
    }

    function resetUI() {
        if (loader) loader.classList.add('hidden');
    }

    // --- State Persistence (localStorage + LZString) ---
    // Simple, synchronous, reliable.
    function saveState() {
        try {
            const stateObj = {
                mainExcelRows: mainExcelRows,
                ownerExcelRows: ownerExcelRows,
                pdfExtractedData: pdfExtractedData,
                attachedOwnerExcel: attachedOwnerExcel,
                clientProfiles: clientProfiles
            };
            const json = JSON.stringify(stateObj);
            if (typeof LZString !== 'undefined') {
                localStorage.setItem('amlState', LZString.compressToUTF16(json));
                localStorage.setItem('amlCompressed', '1');
            } else {
                localStorage.setItem('amlState', json);
                localStorage.setItem('amlCompressed', '0');
            }
        } catch (e) {
            console.warn('saveState failed:', e.message);
        }
    }

    function loadState() {
        try {
            const raw = localStorage.getItem('amlState');
            if (!raw) return;

            const compressed = localStorage.getItem('amlCompressed') === '1';
            let json;
            if (compressed && typeof LZString !== 'undefined') {
                json = LZString.decompressFromUTF16(raw);
            } else {
                json = raw;
            }
            if (!json) return;

            const state = JSON.parse(json);
            if (!state) return;

            attachedOwnerExcel = state.attachedOwnerExcel || null;
            clientProfiles = state.clientProfiles || {};
            pdfExtractedData = state.pdfExtractedData || {};

            if (state.mainExcelRows && state.mainExcelRows.length > 0) {
                mainExcelRows = state.mainExcelRows;
                extractRawData(mainExcelRows, false);
            }
            if (state.ownerExcelRows && state.ownerExcelRows.length > 0) {
                ownerExcelRows = state.ownerExcelRows;
                extractRawData(ownerExcelRows, true);
            }

            if (attachedOwnerExcel) renderOwnerExcelList();

        } catch (e) {
            console.warn('loadState failed:', e.message);
        }
    }

    const btnResetData = document.getElementById('btn-reset-data');
    if (btnResetData) {
        btnResetData.addEventListener('click', () => {
            if (confirm('¿Estás seguro de que quieres borrar todos los datos guardados en esta computadora y empezar de cero?')) {
                localStorage.removeItem('amlState');
                localStorage.removeItem('amlCompressed');
                localStorage.removeItem('amlStateIsCompressed');
                location.reload();
            }
        });
    }

    // Boot: load saved state synchronously
    loadState();
});
