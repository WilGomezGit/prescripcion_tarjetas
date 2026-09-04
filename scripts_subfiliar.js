// Variables globales Sub Filiar
let dataSubFiliar = [];
let filteredData = [];
let headers = [];
let dataTrabajadores = [];

const MAX_FILAS_TABLA = 500;

function showLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'flex';
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'none';
}

// ============================================
// PROCESAR ARCHIVO SUB FILIAR (con filtro correcto)
// ============================================
function processFileSubFiliar() {
    const input = document.getElementById('fileInputSubFiliar');
    const file = input.files[0];

    if (!file) {
        alert('Por favor selecciona un archivo Excel.');
        return;
    }

    showLoading();

    setTimeout(() => {
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const arrayBuffer = e.target.result;
                let workbook;

                try {
                    const data = new Uint8Array(arrayBuffer);
                    workbook = XLSX.read(data, { type: 'array' });
                } catch (err) {
                    // Fallback para .xls
                    const binary = new Uint8Array(arrayBuffer).reduce((acc, byte) => acc + String.fromCharCode(byte), '');
                    workbook = XLSX.read(binary, { type: 'binary' });
                }

                if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                    throw new Error('El archivo no contiene hojas de cálculo.');
                }

                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                // Buscar fila de encabezados
                let headerIndex = -1;
                for (let i = 0; i < rows.length; i++) {
                    const joined = rows[i].map(c => String(c).toLowerCase()).join('|');
                    if (joined.includes('documento') && (joined.includes('valor a prescribir') || joined.includes('saldo actual'))) {
                        headerIndex = i;
                        break;
                    }
                }

                if (headerIndex === -1) {
                    throw new Error('No se encontró la fila de encabezados (Documento, Valor a Prescribir).');
                }

                headers = rows[headerIndex].map(h => String(h).trim());
                const dataRows = rows.slice(headerIndex + 1);

                // Convertir filas en objetos
                dataSubFiliar = dataRows.map(row => {
                    const obj = {};
                    headers.forEach((h, idx) => {
                        if (h === 'Documento') obj[h] = String(row[idx] || '').trim().replace(/['’]/g, '');
                        else obj[h] = row[idx] ? String(row[idx]).trim() : '';
                    });
                    return obj;
                }).filter(o => o['Documento'] && o['Documento'] !== '' && o['Documento'] !== 'TOTAL');

                // 🔥 FILTRO CORREGIDO: solo > 10000
                filteredData = dataSubFiliar.filter(item => {
                    const valor = parseFloat(item['Valor a Prescribir']) || 0;
                    return valor > 10000;
                });

                // Mostrar tabla con SOLO los filtrados (no todos)
                renderTableSubFiliar(filteredData.slice(0, MAX_FILAS_TABLA));

                // Mostrar botón de descarga si hay registros filtrados
                const btn = document.getElementById('btnDescargarExcel');
                if (btn) btn.style.display = filteredData.length > 0 ? 'block' : 'none';

                if (filteredData.length === 0) {
                    alert('No se encontraron registros con Valor a Prescribir > 10000.');
                } else {
                    // Aviso si se truncó la tabla por límite
                    if (filteredData.length > MAX_FILAS_TABLA) {
                        alert(`Se muestran solo las primeras ${MAX_FILAS_TABLA} filas. Total filtrados: ${filteredData.length}`);
                    } else {
                        alert(`Se procesaron ${filteredData.length} registros con Valor a Prescribir > 10000.`);
                    }
                }

                hideLoading();
            } catch (error) {
                console.error(error);
                alert('Ocurrió un error: ' + error.message);
                hideLoading();
            }
        };

        reader.onerror = function() {
            alert('No se pudo leer el archivo.');
            hideLoading();
        };

        reader.readAsArrayBuffer(file);
    }, 100);
}

// ============================================
// RENDERIZAR TABLA (usa la data que se pasa)
// ============================================
function renderTableSubFiliar(data) {
    const thead = document.getElementById('theadPrescripcion');
    const tbody = document.getElementById('tbodyPrescripcion');

    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (data.length === 0) return;

    const headerRow = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    data.forEach(item => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
            const td = document.createElement('td');
            td.textContent = item[h];
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

// ============================================
// LIMPIAR TABLA
// ============================================
function clearTableSubFiliar() {
    dataSubFiliar = [];
    filteredData = [];
    headers = [];

    const input = document.getElementById('fileInputSubFiliar');
    if (input) input.value = '';

    const zona = document.getElementById('zonaSubFiliar');
    if (zona) {
        zona.classList.remove('file-loaded', 'dragover');
        const t = zona.querySelector('.drop-title');
        if (t) t.innerHTML = 'Arrastra archivos Excel aquí o haz clic para seleccionar';
    }

    document.getElementById('theadPrescripcion').innerHTML = '';
    document.getElementById('tbodyPrescripcion').innerHTML = '';

    const btn = document.getElementById('btnDescargarExcel');
    if (btn) btn.style.display = 'none';
}

// ============================================
// DESCARGAR EXCEL DE DOCUMENTOS (solo >10000)
// ============================================
async function downloadFilteredExcel() {
    if (filteredData.length === 0) {
        alert('No hay datos filtrados para descargar.');
        return;
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Documentos');
    ws.columns = [{ width: 15 }];

    filteredData.forEach(item => {
        let d = String(item['Documento']).trim().replace(/['’]/g, '');
        if (/^\d+$/.test(d)) d = Number(d);
        ws.addRow([d]);
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    const hoy = new Date();
    const fecha = `${('0' + hoy.getDate()).slice(-2)}-${('0' + (hoy.getMonth() + 1)).slice(-2)}-${hoy.getFullYear()}`;
    link.href = url;
    link.download = `DocumentosPrescripcion_${fecha}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// ============================================
// PROCESAR TRABAJADORES (SUB FILIAR)
// ============================================
function processFileTrabajadores() {
    const input = document.getElementById('fileInputTrabajadores');
    const file = input.files[0];

    if (!file) {
        alert('Por favor selecciona el archivo de trabajadores.');
        return;
    }

    showLoading();

    setTimeout(() => {
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const arrayBuffer = e.target.result;
                let wb;

                try {
                    const data = new Uint8Array(arrayBuffer);
                    wb = XLSX.read(data, { type: 'array' });
                } catch (err) {
                    const bin = new Uint8Array(arrayBuffer).reduce((a, b) => a + String.fromCharCode(b), '');
                    wb = XLSX.read(bin, { type: 'binary' });
                }

                const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });

                let hIdx = -1;
                for (let i = 0; i < rows.length; i++) {
                    const joined = rows[i].map(c => String(c).toLowerCase()).join('|');
                    if (joined.includes('documento') && joined.includes('nombre') && joined.includes('trabajador')) {
                        hIdx = i;
                        break;
                    }
                }

                if (hIdx === -1) throw new Error('No se encontró encabezado Documento / Nombre Trabajador');

                const raw = rows[hIdx].map(h => String(h).trim());
                const map = raw.map((h, i) => ({ i, name: h.toLowerCase().replace(/\s+/g, '_') }));

                const dIdx = map.find(h => h.name === 'documento')?.i
                    ?? map.find(h => h.name.includes('documento') && !h.name.includes('tipo'))?.i
                    ?? map.find(h => h.name.includes('documento'))?.i;
                const nIdx = map.find(h => h.name.includes('nombre') && h.name.includes('trabajador') && !h.name.includes('empresa'))?.i;
                const cIdx = map.find(h => h.name.includes('celular'))?.i;
                const eIdx = map.find(h => h.name.includes('correo') || h.name.includes('corrreo'))?.i;

                if (dIdx === undefined || nIdx === undefined) throw new Error('No se encontraron columnas Documento o Nombre');

                const dataRows = rows.slice(hIdx + 1);
                dataTrabajadores = dataRows.map(row => ({
                    Documento: String(row[dIdx] || '').trim().replace(/['’]/g, ''),
                    'Nombre Trabajador': String(row[nIdx] || '').trim(),
                    Celular: cIdx !== undefined ? String(row[cIdx] || '').trim() : '',
                    Corrreo: eIdx !== undefined ? String(row[eIdx] || '').trim() : ''
                })).filter(o => o.Documento && o.Documento !== '');

                const btn = document.getElementById('btnDescargarPrescripcionConsumos');
                if (btn) btn.style.display = dataTrabajadores.length > 0 ? 'block' : 'none';

                if (dataTrabajadores.length === 0) {
                    alert('No se encontraron trabajadores con documento válido.');
                } else {
                    alert(`Se procesaron ${dataTrabajadores.length} trabajadores. Puedes descargar el archivo.`);
                }

                hideLoading();
            } catch (error) {
                console.error(error);
                alert('Ocurrió un error: ' + error.message);
                hideLoading();
            }
        };

        reader.onerror = function() {
            alert('No se pudo leer el archivo.');
            hideLoading();
        };

        reader.readAsArrayBuffer(file);
    }, 100);
}

// ============================================
// DESCARGAR PRESCRIPCIONCONSUMOS.XLSX
// ============================================
async function downloadPrescripcionConsumos() {
    if (dataTrabajadores.length === 0) {
        alert('No hay datos para descargar. Primero procesa el archivo de trabajadores.');
        return;
    }

    const maskDocument = (doc) => {
        const clean = String(doc).trim().replace(/['’]/g, '');
        if (clean.length <= 3) return clean;
        return '******' + clean.slice(-3);
    };

    const outputData = dataTrabajadores.map(item => ({
        Documento: maskDocument(item.Documento),
        'Nombre Trabajador': item['Nombre Trabajador'],
        Celular: item.Celular,
        Corrreo: item.Corrreo
    }));

    outputData.sort((a, b) => a['Nombre Trabajador'].localeCompare(b['Nombre Trabajador'], 'es'));

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('PrescripcionConsumos');
    ws.columns = [
        { header: 'Documento', key: 'Documento', width: 20 },
        { header: 'Nombre Trabajador', key: 'Nombre Trabajador', width: 30 },
        { header: 'Celular', key: 'Celular', width: 15 },
        { header: 'Corrreo', key: 'Corrreo', width: 30 }
    ];

    outputData.forEach(item => ws.addRow(item));

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = 'Prescripcion_Consumos_MM_AAAA_Mtx_Gmail.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
