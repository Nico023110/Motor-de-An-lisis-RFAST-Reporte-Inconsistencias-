// Initialize Lucide Icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initApp();
});

// Zonas Configuration (matches Python script)
const ZONAS = {
    "ZonaComuna-01": ["C.S. TERRON COLORADO", "P.S. BELLAVISTA", "P.S. VISTAHERMOSA", "P.S. LA PAZ"],
    "ZonaComuna-03": ["P.S. FRAY DAMIAN", "HOSPITAL CAÑAVERALEJO"],
    "ZonaComuna-17": ["C.S. PRIMERO DE MAYO"],
    "ZonaComuna-18": ["C.S. MELENDEZ", "P.S. ALTO POLVORINES", "P.S. ALTO NAPOLES", "P.S. NAPOLES", "P.S. POLVORINES", "P.S. LOURDES"],
    "ZonaComuna-20": ["C.S. SILOE", "P.S. BELEN", "P.S. BRISAS DE MAYO", "P.S. LA ESTRELLA", "P.S. LA SIRENA", "P.S. LA SULTANA"],
    "ZonaRuralNorte": ["P.S. MONTEBELLO", "P.S. EL SALADITO", "P.S. LA ELVIRA", "P.S. FELIDIA", "P.S. PENAS BLANCAS", "P.S. PICHINDE", "P.S. GOLONDRINAS", "P.S. LA LEONERA", "P.S. LA PAZ RURAL", "P.S. ALTO AGUACATAL", "P.S. LA CASTILLA", "P.S. LOS ANDES"],
    "ZonaRuralSur": ["P.S. LA BUITRERA", "P.S. VILLACARMELO", "P.S. PANCE", "P.S. LA VORAGINE", "P.S. EL HORMIGUERO", "P.S. CASCAJAL"]
};

// App State
const state = {
    uploadedFiles: [],
    rawRows: [],
    inconsistencias: [],
    pagination: {
        all: { page: 1, records: [] },
        zonas: { page: 1, records: [] },
        inconsistencias: { page: 1, records: [] },
        pageSize: 50
    }
};

function initApp() {
    // Navigation Tabs
    const navButtons = document.querySelectorAll('.nav-btn');
    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetTab = btn.getAttribute('data-tab');
            navButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.add('hidden');
            });
            document.getElementById(`tab-${targetTab}`).classList.remove('hidden');
        });
    });

    // File Selection
    const fileInput = document.getElementById('file-input');
    fileInput.addEventListener('change', (e) => handleFilesSelect(Array.from(e.target.files)));

    const dropzone = document.getElementById('main-dropzone');
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files.length) {
            handleFilesSelect(Array.from(e.dataTransfer.files));
        }
    });

    // Buttons
    document.getElementById('btn-ejecutar').addEventListener('click', runAuditPipeline);
    document.getElementById('btn-load-demo').addEventListener('click', loadDemoSampleData);

    // Search Box Listeners
    setupSearch('search-auditoria', 'all');
    setupSearch('search-inconsistencias', 'inconsistencias');

    // Filter Zone Select
    document.getElementById('select-zona-filter').addEventListener('change', (e) => {
        const selected = e.target.value;
        if (selected === 'ALL') {
            state.pagination.zonas.records = state.inconsistencias;
        } else {
            const ipsList = ZONAS[selected] ? ZONAS[selected].map(x => x.toUpperCase()) : [];
            state.pagination.zonas.records = state.inconsistencias.filter(item => {
                const pre = (item.NOMBRE_PRE || '').toUpperCase();
                return ipsList.some(ips => pre.includes(ips));
            });
        }
        state.pagination.zonas.page = 1;
        renderTabTable('zonas');
    });

    // Pagination Listeners
    setupPaginationControls('all', 'btn-prev-page', 'btn-next-page');
    setupPaginationControls('zonas', 'btn-prev-zonas', 'btn-next-zonas');
    setupPaginationControls('inconsistencias', 'btn-prev-inconsistencias', 'btn-next-inconsistencias');

    // Export Button
    document.getElementById('btn-export-excel').addEventListener('click', exportAuditExcelReport);
}

// Handle File Selection & Parsing
async function handleFilesSelect(files) {
    if (!files || !files.length) return;

    showLoader(`Leyendo y procesando ${files.length} archivo(s)...`);
    state.uploadedFiles = files;
    state.rawRows = [];

    const fileTagsList = document.getElementById('file-tags-list');
    fileTagsList.innerHTML = '';

    for (let file of files) {
        const span = document.createElement('span');
        span.className = 'file-tag';
        span.textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        fileTagsList.appendChild(span);

        try {
            const rows = await parseFile(file);
            if (rows && rows.length) {
                state.rawRows = state.rawRows.concat(rows);
            }
        } catch (err) {
            console.warn(`Error al leer ${file.name}:`, err);
        }
    }

    document.getElementById('btn-ejecutar').disabled = state.rawRows.length === 0;
    hideLoader();

    if (state.rawRows.length > 0) {
        runAuditPipeline();
    }
}

// File Parsing Supporting Excel & CSV
async function parseFile(file) {
    const ext = file.name.toLowerCase();

    if (ext.endsWith('.xlsx')) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.worksheets[0];
            const rows = [];
            const headers = [];
            let headerParsed = false;

            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                const values = row.values;
                if (!headerParsed) {
                    for (let i = 1; i < values.length; i++) {
                        headers.push(values[i] ? values[i].toString().trim().toLowerCase() : `col_${i}`);
                    }
                    headerParsed = true;
                } else {
                    const rowObj = {};
                    for (let i = 1; i <= headers.length; i++) {
                        const header = headers[i - 1];
                        let val = values[i];
                        rowObj[header] = val !== null && val !== undefined ? val.toString().trim() : '';
                    }
                    rows.push(rowObj);
                }
            });
            return rows;
        } catch (err) {
            return parseCSVText(await file.text());
        }
    } else {
        const arrayBuffer = await file.arrayBuffer();
        const decoder = new TextDecoder('iso-8859-1');
        const text = decoder.decode(arrayBuffer);
        return parseCSVText(text);
    }
}

// Multi-Delimiter CSV Parser
function parseCSVText(text) {
    if (!text || !text.trim()) return [];

    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length < 2) return [];

    const firstLine = lines[0];
    const semiCount = (firstLine.match(/;/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const delimiter = semiCount >= commaCount ? ';' : ',';

    const headers = firstLine.split(delimiter).map(h => h.replace(/^["']|["']$/g, '').trim().toLowerCase());
    const result = [];

    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(delimiter);
        const row = {};
        headers.forEach((header, index) => {
            let val = values[index] !== undefined ? values[index] : '';
            row[header] = val.replace(/^["']|["']$/g, '').trim();
        });
        result.push(row);
    }
    return result;
}

// Core Audit Engine (Implements exact 7 Python script rules)
function runAuditPipeline() {
    if (!state.rawRows.length) return;

    showLoader('Ejecutando Motor de Auditoría (7 Validaciones Clínicas)...');
    setTimeout(() => {
        state.inconsistencias = [];

        state.rawRows.forEach((r, idx) => {
            // Helper para obtener valor ignorando pequeñas variaciones en headers
            const getVal = (keys) => {
                for (let k of keys) {
                    for (let key in r) {
                        if (key.trim().toLowerCase() === k.toLowerCase()) return String(r[key]).trim();
                    }
                }
                return '';
            };

            const codigo = getVal(['codigo', 'cod']);
            const nombre = getVal(['nombre', 'actividad', 'servicio']);
            const finalidad = getVal(['finalidad']);
            const finalidad_rips = getVal(['finalidad_rips', 'finalidad_']);
            const nombre_finalidad = getVal(['nombre_finalidad', 'nombre_fin']);
            const cexterna = getVal(['cexterna', 'causa_externa']);
            const nombre_cexterna = getVal(['nombre_cexterna', 'nombre_cex']);
            const atencion = getVal(['atencion', 'fecha_atencion']);
            const cajero = getVal(['cajero']);
            const nombre_cajero = getVal(['nombre_cajero', 'nombre_caj']);
            const nombre_prestador = getVal(['nombre_prestador', 'nombre_pre', 'ips', 'prestador']).replace(/ VISTA HERMOSA/g, ' VISTAHERMOSA');
            const centroprod = getVal(['centroprod', 'cod_centroprod']);
            const nombre_centroproduccion = getVal(['nombre_centroproduccion', 'nombre_cen']);
            const profesional = getVal(['profesional', 'profesiona']);
            const nombre_profesional = getVal(['nombre_profesional', 'nombre_pro']);
            const documento = getVal(['documento', 'cedula', 'num_doc']);
            const factura = getVal(['factura', 'num_factura']);
            const dx_principal = getVal(['dx_principal', 'dx_princip', 'cie10']);
            const nombre_dx_principal = getVal(['nombre_dx_principal', 'nombre_dx_']);
            const dx_relacionado = getVal(['dx_relacionado', 'dx_relacio']);
            const nombre_dx_relacionado = getVal(['nombre_dx_relacionado', 'nombre_dx2']);

            const rowData = {
                CODIGO: codigo,
                NOMBRE: nombre,
                FINALIDAD: finalidad,
                FINALIDAD_RIPS: finalidad_rips,
                NOMBRE_FINALIDAD: nombre_finalidad,
                CEXTERNA: cexterna,
                NOMBRE_CEX: nombre_cexterna,
                ATENCION: atencion,
                CAJERO: cajero,
                NOMBRE_CAJ: nombre_cajero,
                NOMBRE_PRE: nombre_prestador,
                CENTROPROD: centroprod,
                NOMBRE_CEN: nombre_centroproduccion,
                PROFESIONA: profesional,
                NOMBRE_PRO: nombre_profesional,
                DOCUMENTO: documento,
                FACTURA: factura,
                DX_PRINCIP: dx_principal,
                NOMBRE_DX_: nombre_dx_principal,
                DX_RELACIO: dx_relacionado,
                NOMBRE_DX2: nombre_dx_relacionado
            };

            const centroUpper = nombre_centroproduccion.toUpperCase();
            const nombreUpper = nombre.toUpperCase();

            // 1. VALIDACIÓN 01 - Causa Externa Incorrecta PyM
            if (cexterna === '38' && finalidad_rips === '11' && 
                (centroUpper.includes('CURSO DE VIDA') || centroUpper.includes('PLANIFICACION FAMILIAR')) &&
                (nombreUpper.includes('PRIMERA VEZ') || nombreUpper.includes('SEGUIMIENTO'))) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "La causa externa no puede ser la 38 (Enfermedad General), debe ser la 40 (Promoción y mantenimiento)."
                });
                return;
            }

            // 2. VALIDACIÓN 02 - Finalidad Incorrecta en Programas de Control
            if (['1415', '1416', '1417'].some(cp => centroprod.includes(cp)) &&
                (nombreUpper.includes('CONTROL') || nombreUpper.includes('SEGUIMIENTO')) &&
                !['16', '17', '23', '0', '28'].includes(finalidad_rips) &&
                (!dx_relacionado || dx_relacionado === dx_principal)) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "La finalidad debe ser 28 (tratamiento) porque son consultas de control y seguimiento de pacientes con diagnósticos ya definidos (ej. hipertensión)."
                });
                return;
            }

            // 3. VALIDACIÓN 03 - Finalidad Incorrecta en Consultas de Primera Vez
            if (nombreUpper.includes('PRIMERA VEZ') &&
                ['1415', '1416', '1417'].includes(centroprod) &&
                !['15', '23', '0'].includes(finalidad_rips)) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "La finalidad debe ser 27 (diagnóstico) porque es una consulta por primera vez, el objetivo principal es evaluar al paciente para confirmar o definir un diagnóstico."
                });
                return;
            }

            // 4. VALIDACIÓN 04 - Finalidad Incorrecta en Controles de Odontología
            if (nombreUpper.includes('CONTROL') &&
                ['1300', '1303'].includes(centroprod) &&
                !['16', '17', '0', '23'].includes(finalidad_rips)) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "Debe ser finalidad 28 (Tratamiento) porque la consulta de control o seguimiento en odontología hace parte de la continuidad del manejo clínico del paciente."
                });
                return;
            }

            // 5. VALIDACIÓN 05 - Finalidad Incorrecta en Planificación Familiar
            if (nombreUpper.includes('CONSULTA') &&
                centroprod === '1405' &&
                !['19', '21', '23', '25', '0'].includes(finalidad_rips)) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "Debe ser finalidad 31 (Planificación familiar y anticoncepción) y la causa externa 40 (Promoción y mantenimiento) porque la consulta corresponde a una atención de primera vez dentro del programa de planificación familiar (PYP PF)."
                });
                return;
            }

            // 6. VALIDACIÓN 06 - Finalidad Incorrecta en Detección Temprana
            if (nombreUpper.includes('CONSULTA') &&
                ['1408', '1409', '1439', '1440'].includes(centroprod) &&
                !['12', '15', '16', '23', '0'].includes(finalidad_rips)) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "La finalidad debe ser 24 (detección temprana de enfermedad general) porque las atenciones corresponden a consultas de primera vez dentro de programas de detección de cáncer, cuyo objetivo es identificar de manera oportuna posibles enfermedades, no realizar tratamiento ni solo valoración general."
                });
                return;
            }

            // 7. VALIDACIÓN 07 - Finalidad Incorrecta en Educación Individual
            if (nombreUpper.includes('EDUCACION INDIVIDUAL') &&
                !['0', '19', '20', '23', '28', '29', '30', '32', '33', '34', '38', '39', '40', '41', '42'].includes(finalidad_rips)) {
                state.inconsistencias.push({
                    ...rowData,
                    "Inconsistencia a Corregir": "La finalidad registrada no corresponde a una actividad de educación individual. Estas atenciones deben registrarse con una finalidad de Promoción de la Salud (finalidades 40 a 54), ya que su objetivo es la educación y promoción, no el diagnóstico, tratamiento o seguimiento clínico."
                });
                return;
            }
        });

        updateUI();
        hideLoader();
    }, 600);
}

// Update UI & KPI Cards
function updateUI() {
    const totalEvaluados = state.rawRows.length;
    const totalInconsistencias = state.inconsistencias.length;
    const conformes = totalEvaluados - totalInconsistencias;
    const tasaConformidad = totalEvaluados ? ((conformes / totalEvaluados) * 100).toFixed(1) : 100;

    document.getElementById('kpi-total').textContent = totalEvaluados.toLocaleString();
    document.getElementById('kpi-inconsistencias').textContent = totalInconsistencias.toLocaleString();
    document.getElementById('kpi-conformes').textContent = conformes.toLocaleString();
    document.getElementById('kpi-tasa').textContent = `${tasaConformidad}%`;

    document.getElementById('badge-count').textContent = `${totalInconsistencias.toLocaleString()} inconsistencias`;
    document.getElementById('badge-zonas').textContent = `${Object.keys(ZONAS).length} zonas activas`;
    document.getElementById('badge-inconsistencias').textContent = `${totalInconsistencias.toLocaleString()} registros`;

    document.getElementById('btn-export-excel').disabled = totalInconsistencias === 0;

    state.pagination.all.records = state.inconsistencias; state.pagination.all.page = 1;
    state.pagination.zonas.records = state.inconsistencias; state.pagination.zonas.page = 1;
    state.pagination.inconsistencias.records = state.inconsistencias; state.pagination.inconsistencias.page = 1;

    renderTabTable('all');
    renderTabTable('zonas');
    renderTabTable('inconsistencias');
}

// Render Table according to Tab
function renderTabTable(tabKey) {
    const tbodyId = tabKey === 'all' ? 'table-body' : `table-body-${tabKey}`;
    const pageInfoId = tabKey === 'all' ? 'page-info' : `page-info-${tabKey}`;
    const pageNumId = tabKey === 'all' ? 'current-page-num' : `page-num-${tabKey}`;
    const prevBtnId = tabKey === 'all' ? 'btn-prev-page' : `btn-prev-${tabKey}`;
    const nextBtnId = tabKey === 'all' ? 'btn-next-page' : `btn-next-${tabKey}`;

    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = '';

    const tabState = state.pagination[tabKey];
    const records = tabState.records;
    const pageSize = state.pagination.pageSize;
    const totalPages = Math.max(1, Math.ceil(records.length / pageSize));
    const currentPage = Math.min(tabState.page, totalPages);
    tabState.page = currentPage;

    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, records.length);
    const pageRecords = records.slice(startIdx, endIdx);

    if (!records.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="${tabKey === 'zonas' ? 6 : 7}" class="empty-table-msg">
                    <i data-lucide="check-circle-2" style="color: #10B981"></i>
                    <p>No se encontraron inconsistencias para esta vista.</p>
                </td>
            </tr>
        `;
        document.getElementById(pageInfoId).textContent = 'Mostrando 0 registros';
        document.getElementById(prevBtnId).disabled = true;
        document.getElementById(nextBtnId).disabled = true;
        lucide.createIcons();
        return;
    }

    pageRecords.forEach(row => {
        const tr = document.createElement('tr');
        if (tabKey === 'zonas') {
            let zonaName = 'Otras Zonas / Cabecera';
            const preUpper = (row.NOMBRE_PRE || '').toUpperCase();
            for (let z in ZONAS) {
                if (ZONAS[z].some(ips => preUpper.includes(ips))) {
                    zonaName = z;
                    break;
                }
            }

            tr.innerHTML = `
                <td><span class="badge badge-primary">${zonaName}</span></td>
                <td><strong>${row.NOMBRE_PRE || 'E.S.E. LADERA'}</strong></td>
                <td><code>${row.CODIGO || '-'}</code></td>
                <td>${row.NOMBRE || 'PACIENTE AUDITADO'}</td>
                <td><code>${row.FACTURA || '-'}</code></td>
                <td><span style="color: #EF4444; font-weight: 500">${row["Inconsistencia a Corregir"]}</span></td>
            `;
        } else {
            tr.innerHTML = `
                <td><code>${row.CODIGO || '-'}</code></td>
                <td><code>${row.DOCUMENTO || '-'}</code></td>
                <td><strong>${row.NOMBRE || '-'}</strong></td>
                <td>${row.NOMBRE_PRE || 'E.S.E. LADERA'}</td>
                <td><code>${row.CENTROPROD || '-'}</code></td>
                <td><span class="badge badge-warning">RIPS ${row.FINALIDAD_RIPS || '-'}</span></td>
                <td><span style="color: #EF4444; font-weight: 500">${row["Inconsistencia a Corregir"]}</span></td>
            `;
        }
        tbody.appendChild(tr);
    });

    document.getElementById(pageInfoId).textContent = `Mostrando ${(startIdx + 1).toLocaleString()} a ${endIdx.toLocaleString()} de ${records.length.toLocaleString()} registros`;
    document.getElementById(pageNumId).textContent = `Página ${currentPage} de ${totalPages}`;
    document.getElementById(prevBtnId).disabled = currentPage === 1;
    document.getElementById(nextBtnId).disabled = currentPage === totalPages;
}

function setupPaginationControls(tabKey, prevBtnId, nextBtnId) {
    document.getElementById(prevBtnId).addEventListener('click', () => {
        if (state.pagination[tabKey].page > 1) {
            state.pagination[tabKey].page--;
            renderTabTable(tabKey);
        }
    });

    document.getElementById(nextBtnId).addEventListener('click', () => {
        const totalPages = Math.ceil(state.pagination[tabKey].records.length / state.pagination.pageSize);
        if (state.pagination[tabKey].page < totalPages) {
            state.pagination[tabKey].page++;
            renderTabTable(tabKey);
        }
    });
}

function setupSearch(inputId, tabKey) {
    document.getElementById(inputId).addEventListener('input', (e) => {
        const q = e.target.value.toLowerCase();
        const source = state.inconsistencias;

        state.pagination[tabKey].records = source.filter(row => JSON.stringify(row).toLowerCase().includes(q));
        state.pagination[tabKey].page = 1;
        renderTabTable(tabKey);
    });
}

// Load Demo Dataset Simulation
function loadDemoSampleData() {
    showLoader('Cargando dataset de prueba de atenciones (1.800 registros)...');
    setTimeout(() => {
        const sampleRows = [];
        const prestadores = [
            'C.S. TERRON COLORADO', 'P.S. BELLAVISTA', 'P.S. VISTAHERMOSA',
            'P.S. FRAY DAMIAN', 'HOSPITAL CAÑAVERALEJO', 'C.S. PRIMERO DE MAYO',
            'C.S. MELENDEZ', 'P.S. ALTO POLVORINES', 'C.S. SILOE', 'P.S. MONTEBELLO'
        ];

        for (let i = 1; i <= 1800; i++) {
            const isError = i % 2 === 0;
            const ips = prestadores[i % prestadores.length];

            sampleRows.push({
                codigo: `ACT-${1000 + i}`,
                nombre: isError ? 'CONSULTA DE PRIMERA VEZ POR MEDICINA GENERAL' : 'CONSULTA DE CONTROL POR MEDICINA GENERAL',
                finalidad: '10',
                finalidad_rips: isError ? '11' : '28',
                nombre_finalidad: 'DIAGNOSTICO',
                cexterna: isError ? '38' : '40',
                nombre_cexterna: isError ? 'ENFERMEDAD GENERAL' : 'PROMOCION Y MANTENIMIENTO',
                atencion: '2026-07-15',
                cajero: 'CAJ-01',
                nombre_cajero: 'CAJERO SISTEMA AUDITORIA',
                nombre_prestador: ips,
                centroprod: isError ? '1415' : '1000',
                nombre_centroproduccion: isError ? 'CURSO DE VIDA ADULTO' : 'CONSULTA EXTERNA GENERAL',
                profesional: 'MED-992',
                nombre_profesional: `DR. MEDICO AUDITADO #${i}`,
                documento: `11440${i + 500}`,
                factura: `FAC-2026-${i + 100}`,
                dx_principal: 'I10X',
                nombre_dx_principal: 'HIPERTENSION ARTERIAL',
                dx_relacionado: '',
                nombre_dx_relacionado: ''
            });
        }

        state.rawRows = sampleRows;
        document.getElementById('btn-ejecutar').disabled = false;
        runAuditPipeline();
    }, 600);
}

// Export Full Multi-Tab Excel Report (Matches original Python script formatting)
async function exportAuditExcelReport() {
    if (!state.inconsistencias.length) return;

    const workbook = new ExcelJS.Workbook();

    // 1. Hoja Principal
    const mainSheet = workbook.addWorksheet('Todas las inconsistencias');
    setupSheetHeadersAndData(mainSheet, state.inconsistencias);

    // 2. Hojas por Zona
    for (let nombreHoja in ZONAS) {
        const ipsList = ZONAS[nombreHoja].map(x => x.toUpperCase());
        const datosZona = state.inconsistencias.filter(item => {
            const pre = (item.NOMBRE_PRE || '').toUpperCase();
            return ipsList.some(ips => pre.includes(ips));
        });

        if (datosZona.length > 0) {
            const sheet = workbook.addWorksheet(nombreHoja);
            setupSheetHeadersAndData(sheet, datosZona);
        }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Reporte_Inconsistencias_Clinicas_${new Date().toISOString().slice(0, 10)}.xlsx`;
    link.click();
}

function setupSheetHeadersAndData(sheet, records) {
    const columns = [
        'CODIGO', 'NOMBRE', 'FINALIDAD', 'FINALIDAD_RIPS', 'NOMBRE_FINALIDAD',
        'CEXTERNA', 'NOMBRE_CEX', 'ATENCION', 'CAJERO', 'NOMBRE_CAJ',
        'NOMBRE_PRE', 'CENTROPROD', 'NOMBRE_CEN', 'PROFESIONA', 'NOMBRE_PRO',
        'DOCUMENTO', 'FACTURA', 'DX_PRINCIP', 'NOMBRE_DX_', 'DX_RELACIO',
        'NOMBRE_DX2', 'Inconsistencia a Corregir'
    ];

    sheet.columns = columns.map(c => ({
        header: c,
        key: c,
        width: c === 'Inconsistencia a Corregir' ? 45 : (c.includes('NOMBRE') ? 28 : 16)
    }));

    // Header formatting: Yellow background (FFFF00) like Python script
    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF00' }
        };
        cell.font = { bold: true, color: { argb: '000000' } };
    });

    // Add rows and format 'Inconsistencia a Corregir' column text as Red (FF0000)
    records.forEach(item => {
        const row = sheet.addRow(item);
        const cellInc = row.getCell('Inconsistencia a Corregir');
        if (cellInc) {
            cellInc.font = { color: { argb: 'FF0000' }, bold: true };
        }
    });
}

function showLoader(msg) {
    document.getElementById('loader-message').textContent = msg;
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}
