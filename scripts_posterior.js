// Variables globales Posteriores
let dataPosteriorOriginal = [];
let filteredDataPosteriorOriginal = [];
let headersPosteriorOriginal = [];
let dataPosteriorTrabajadores = [];

const MAX_FILAS_TABLA = 500;

// Lista de municipios Popayán
const MUNICIPIOS_POPAYAN = [
    'BOLIVAR','CAJIBIO','COCONUCO','EL TAMBO','FLORENCIA','INZA','SAN SEBASTIAN',
    'LA SIERRA','LA VEGA','JAMBALO','MORALES','PAEZ-BELALCAZAR','PASTO',
    'PATIA (EL BORDO)','PIAMONTE','PIENDAMO','POPAYAN','PURACE-COCONUCO','ROSAS',
    'SILVIA','SOTARA','SUCRE','TIMBIO','TIMBIQUI','TORIBIO','TOTORO'
];

// Lista de municipios Zona Norte
const MUNICIPIOS_ZONA_NORTE = [
    'ALMAGUER','PRADERA','BUENOS AIRES','PUERTO TEJADA','CALDONO','CALI',
    'CANDELARIA','PALMIRA','SANTA ROSA','SANTAFE DE BOGOTA','FLORIDA',
    'FLORIDABLANCA','GARZON','GUACHENE','GUAPI','IBAGUE','TULUA','TUMACO',
    'JAMUNDI','LA SIERRA','VILLARICA','LOPEZ DE MICAY','MEDELLIN','MIRANDA',
    'MORALES','CALOTO','CORINTO','SUAREZ','SANTANDER DE QUILICHAO','YUMBO',
    'PADILLA','PALMIRA','MIRANDA','JAMBALO'
];

function showLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'flex';
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.style.display = 'none';
}

// ============================================
// PROCESAR ARCHIVO ORIGINAL (sin filtro)
// ============================================
function processFilePosteriorOriginal() {
    const input = document.getElementById('fileInputPosteriorOriginal');
    const file = input.files[0];
    if (!file) {
        alert('Por favor selecciona el archivo de prescripción posterior.');
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
                    if (joined.includes('documento') && (joined.includes('valor a prescribir') || joined.includes('saldo actual'))) {
                        hIdx = i;
                        break;
                    }
                }
                if (hIdx === -1) throw new Error('No se encontró la fila de encabezados (Documento, Valor a Prescribir).');

                headersPosteriorOriginal = rows[hIdx].map(h => String(h).trim());
                const dataRows = rows.slice(hIdx + 1);

                dataPosteriorOriginal = dataRows.map(row => {
                    const obj = {};
                    headersPosteriorOriginal.forEach((h, idx) => {
                        if (h === 'Documento') obj[h] = String(row[idx] || '').trim().replace(/['’]/g, '');
                        else obj[h] = row[idx] ? String(row[idx]).trim() : '';
                    });
                    return obj;
                }).filter(o => o['Documento'] && o['Documento'] !== '' && o['Documento'] !== 'TOTAL');

                filteredDataPosteriorOriginal = dataPosteriorOriginal; // SIN FILTRO

                const btn = document.getElementById('btnDescargarExcelPosteriorOriginal');
                if (btn) btn.style.display = filteredDataPosteriorOriginal.length > 0 ? 'block' : 'none';

                if (filteredDataPosteriorOriginal.length === 0) alert('No se encontraron registros con documento válido.');
                else alert(`Se procesaron ${filteredDataPosteriorOriginal.length} registros. Puedes descargar el Excel de documentos.`);

                hideLoading();
            } catch (error) {
                console.error(error);
                alert('Ocurrió un error: ' + error.message);
                hideLoading();
            }
        };
        reader.onerror = () => { alert('No se pudo leer el archivo.'); hideLoading(); };
        reader.readAsArrayBuffer(file);
    }, 100);
}

// Descargar Excel de documentos (sin filtro)
async function downloadFilteredExcelPosteriorOriginal() {
    if (filteredDataPosteriorOriginal.length === 0) {
        alert('No hay datos para descargar.');
        return;
    }
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Documentos');
    ws.columns = [{ width: 15 }];
    filteredDataPosteriorOriginal.forEach(item => {
        let d = String(item['Documento']).trim().replace(/['’]/g, '');
        if (/^\d+$/.test(d)) d = Number(d);
        ws.addRow([d]);
    });
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `DocumentosPosteriores_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// PROCESAR TRABAJADORES (POSTERIORES)
// ============================================
function processFilePosteriorTrabajadores() {
    const input = document.getElementById('fileInputPosteriorTrabajadores');
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
                const dIdx = map.find(h => h.name === 'documento')?.i ?? map.find(h => h.name.includes('documento') && !h.name.includes('tipo'))?.i ?? map.find(h => h.name.includes('documento'))?.i;
                const nIdx = map.find(h => h.name.includes('nombre') && h.name.includes('trabajador') && !h.name.includes('empresa'))?.i;
                const tIdx = map.find(h => h.name.includes('teléfono') || h.name.includes('telefono'))?.i;
                const cIdx = map.find(h => h.name.includes('celular'))?.i;
                const eIdx = map.find(h => h.name.includes('correo') || h.name.includes('corrreo'))?.i;
                const stIdx = map.find(h => h.name.includes('estado'))?.i;
                const muNacIdx = map.find(h => h.name.includes('municipio_nac'))?.i;
                const nitEmpIdx = map.find(h => h.name.includes('nit_empresa'))?.i;
                const nomEmpIdx = map.find(h => h.name.includes('nombre_empresa'))?.i;
                const muEmpIdx = map.find(h => h.name.includes('municipio_empresa'))?.i;
                const telLocIdx = map.find(h => h.name.includes('tel._local') || h.name.includes('tel_local'))?.i;
                const celLocIdx = map.find(h => h.name.includes('cel._local') || h.name.includes('cel_local'))?.i;
                const dirResIdx = map.find(h => h.name.includes('dirección_residencia') || h.name.includes('direccion_residencia'))?.i;

                if (dIdx === undefined || nIdx === undefined) throw new Error('No se encontraron columnas Documento o Nombre');

                const dataRows = rows.slice(hIdx + 1);
                dataPosteriorTrabajadores = dataRows.map(row => {
                    const doc = String(row[dIdx] || '').trim().replace(/['’]/g, '');
                    const obj = {
                        'Documento': doc,
                        'Nombre Trabajador': String(row[nIdx] || '').trim(),
                        'Teléfono': tIdx !== undefined ? String(row[tIdx] || '').trim() : '',
                        'Celular': cIdx !== undefined ? String(row[cIdx] || '').trim() : '',
                        'Corrreo': eIdx !== undefined ? String(row[eIdx] || '').trim() : '',
                        'Estado': stIdx !== undefined ? String(row[stIdx] || '').trim() : '',
                        'Municipio Nac': muNacIdx !== undefined ? String(row[muNacIdx] || '').trim() : '',
                        'Nit Empresa': nitEmpIdx !== undefined ? String(row[nitEmpIdx] || '').trim() : '',
                        'Nombre Empresa': nomEmpIdx !== undefined ? String(row[nomEmpIdx] || '').trim() : '',
                        'Municipio Empresa': muEmpIdx !== undefined ? String(row[muEmpIdx] || '').trim() : '',
                        'Tel. Local / Ppal': telLocIdx !== undefined ? String(row[telLocIdx] || '').trim() : '',
                        'Cel. Local / Ppal': celLocIdx !== undefined ? String(row[celLocIdx] || '').trim() : '',
                        'Dirección Residencia': dirResIdx !== undefined ? String(row[dirResIdx] || '').trim() : ''
                    };
                    return obj;
                }).filter(o => o['Documento'] && o['Documento'] !== '');

                // Mostrar botones adicionales
                const botonesExtra = document.getElementById('botonesPosterioresAdicionales');
                if (botonesExtra) botonesExtra.style.display = dataPosteriorTrabajadores.length > 0 ? 'block' : 'none';

                const btn = document.getElementById('btnDescargarPosteriorFinal');
                if (btn) btn.style.display = dataPosteriorTrabajadores.length > 0 ? 'block' : 'none';

                if (dataPosteriorTrabajadores.length === 0) {
                    alert('No se encontraron trabajadores con documento válido.');
                } else {
                    alert(`Se procesaron ${dataPosteriorTrabajadores.length} trabajadores. Puedes descargar los archivos.`);
                }

                hideLoading();
            } catch (error) {
                console.error(error);
                alert('Ocurrió un error: ' + error.message);
                hideLoading();
            }
        };
        reader.onerror = () => { alert('No se pudo leer el archivo.'); hideLoading(); };
        reader.readAsArrayBuffer(file);
    }, 100);
}

// ============================================
// DESCARGAR EXCEL FINAL COMPLETO (con saldos)
// ============================================
async function downloadPosteriorFinal() {
    if (dataPosteriorTrabajadores.length === 0) {
        alert('No hay datos de trabajadores para procesar.');
        return;
    }
    if (filteredDataPosteriorOriginal.length === 0) {
        alert('No hay datos del archivo original. Procesa primero el archivo de prescripción.');
        return;
    }

    const mapOriginal = new Map();
    filteredDataPosteriorOriginal.forEach(item => {
        const doc = String(item['Documento']).trim().replace(/['’]/g, '');
        mapOriginal.set(doc, item);
    });

    const outputRows = [];
    dataPosteriorTrabajadores.forEach((trabajador, index) => {
        const doc = String(trabajador['Documento']).trim().replace(/['’]/g, '');
        const original = mapOriginal.get(doc) || {};
        outputRows.push({
            'No.': index + 1,
            'Documento': doc,
            'Nombre Trabajador': trabajador['Nombre Trabajador'],
            'Teléfono': trabajador['Teléfono'],
            'Celular': trabajador['Celular'],
            'Corrreo': trabajador['Corrreo'],
            'Saldo Actual': original['Saldo Actual'] || '',
            'Valor a Prescribir': original['Valor a Prescribir'] || '',
            'Observación': '',
            'Estado': trabajador['Estado'],
            'Municipio Nac': trabajador['Municipio Nac'],
            'Nit Empresa': trabajador['Nit Empresa'],
            'Nombre Empresa': trabajador['Nombre Empresa'],
            'Municipio Empresa': trabajador['Municipio Empresa'],
            'Tel. Local / Ppal': trabajador['Tel. Local / Ppal'],
            'Cel. Local / Ppal': trabajador['Cel. Local / Ppal'],
            'Dirección Residencia': trabajador['Dirección Residencia']
        });
    });

    outputRows.sort((a, b) => a['Nombre Trabajador'].localeCompare(b['Nombre Trabajador'], 'es'));
    outputRows.forEach((row, idx) => row['No.'] = idx + 1);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Posteriores');
    ws.columns = [
        { header: 'No.', key: 'No.', width: 6 },
        { header: 'Documento', key: 'Documento', width: 20 },
        { header: 'Nombre Trabajador', key: 'Nombre Trabajador', width: 30 },
        { header: 'Teléfono', key: 'Teléfono', width: 15 },
        { header: 'Celular', key: 'Celular', width: 15 },
        { header: 'Corrreo', key: 'Corrreo', width: 30 },
        { header: 'Saldo Actual', key: 'Saldo Actual', width: 15 },
        { header: 'Valor a Prescribir', key: 'Valor a Prescribir', width: 18 },
        { header: 'Observación', key: 'Observación', width: 20 },
        { header: 'Estado', key: 'Estado', width: 15 },
        { header: 'Municipio Nac', key: 'Municipio Nac', width: 18 },
        { header: 'Nit Empresa', key: 'Nit Empresa', width: 15 },
        { header: 'Nombre Empresa', key: 'Nombre Empresa', width: 25 },
        { header: 'Municipio Empresa', key: 'Municipio Empresa', width: 20 },
        { header: 'Tel. Local / Ppal', key: 'Tel. Local / Ppal', width: 18 },
        { header: 'Cel. Local / Ppal', key: 'Cel. Local / Ppal', width: 18 },
        { header: 'Dirección Residencia', key: 'Dirección Residencia', width: 30 }
    ];
    outputRows.forEach(row => ws.addRow(row));

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'PrescripcionConsumosPosteriores.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// DESCARGAR POPAYÁN (filtrado por municipio, sin saldos)
// ============================================
async function downloadPopayan() {
    if (dataPosteriorTrabajadores.length === 0) {
        alert('No hay datos de trabajadores. Primero procesa el archivo.');
        return;
    }
    const filtrados = dataPosteriorTrabajadores.filter(t => MUNICIPIOS_POPAYAN.includes(String(t['Municipio Empresa']).toUpperCase().trim()));
    if (filtrados.length === 0) {
        alert('No hay registros para Popayán.');
        return;
    }
    const output = filtrados.map((t, i) => ({
        'No.': i + 1,
        'Documento': t['Documento'],
        'Nombre Trabajador': t['Nombre Trabajador'],
        'Teléfono': t['Teléfono'],
        'Celular': t['Celular'],
        'Corrreo': t['Corrreo'],
        'Estado': t['Estado'],
        'Municipio Nac': t['Municipio Nac'],
        'Nit Empresa': t['Nit Empresa'],
        'Nombre Empresa': t['Nombre Empresa'],
        'Municipio Empresa': t['Municipio Empresa'],
        'Tel. Local / Ppal': t['Tel. Local / Ppal'],
        'Cel. Local / Ppal': t['Cel. Local / Ppal'],
        'Dirección Residencia': t['Dirección Residencia']
    }));
    output.sort((a, b) => a['Nombre Trabajador'].localeCompare(b['Nombre Trabajador'], 'es'));
    output.forEach((r, i) => r['No.'] = i + 1);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Popayán');
    ws.columns = [
        { header: 'No.', key: 'No.', width: 6 },
        { header: 'Documento', key: 'Documento', width: 20 },
        { header: 'Nombre Trabajador', key: 'Nombre Trabajador', width: 30 },
        { header: 'Teléfono', key: 'Teléfono', width: 15 },
        { header: 'Celular', key: 'Celular', width: 15 },
        { header: 'Corrreo', key: 'Corrreo', width: 30 },
        { header: 'Estado', key: 'Estado', width: 15 },
        { header: 'Municipio Nac', key: 'Municipio Nac', width: 18 },
        { header: 'Nit Empresa', key: 'Nit Empresa', width: 15 },
        { header: 'Nombre Empresa', key: 'Nombre Empresa', width: 25 },
        { header: 'Municipio Empresa', key: 'Municipio Empresa', width: 20 },
        { header: 'Tel. Local / Ppal', key: 'Tel. Local / Ppal', width: 18 },
        { header: 'Cel. Local / Ppal', key: 'Cel. Local / Ppal', width: 18 },
        { header: 'Dirección Residencia', key: 'Dirección Residencia', width: 30 }
    ];
    output.forEach(r => ws.addRow(r));
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'Popayan.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// DESCARGAR ZONA NORTE (filtrado por municipio, sin saldos)
// ============================================
async function downloadZonaNorte() {
    if (dataPosteriorTrabajadores.length === 0) {
        alert('No hay datos de trabajadores. Primero procesa el archivo.');
        return;
    }
    const filtrados = dataPosteriorTrabajadores.filter(t => MUNICIPIOS_ZONA_NORTE.includes(String(t['Municipio Empresa']).toUpperCase().trim()));
    if (filtrados.length === 0) {
        alert('No hay registros para Zona Norte.');
        return;
    }
    const output = filtrados.map((t, i) => ({
        'No.': i + 1,
        'Documento': t['Documento'],
        'Nombre Trabajador': t['Nombre Trabajador'],
        'Teléfono': t['Teléfono'],
        'Celular': t['Celular'],
        'Corrreo': t['Corrreo'],
        'Estado': t['Estado'],
        'Municipio Nac': t['Municipio Nac'],
        'Nit Empresa': t['Nit Empresa'],
        'Nombre Empresa': t['Nombre Empresa'],
        'Municipio Empresa': t['Municipio Empresa'],
        'Tel. Local / Ppal': t['Tel. Local / Ppal'],
        'Cel. Local / Ppal': t['Cel. Local / Ppal'],
        'Dirección Residencia': t['Dirección Residencia']
    }));
    output.sort((a, b) => a['Nombre Trabajador'].localeCompare(b['Nombre Trabajador'], 'es'));
    output.forEach((r, i) => r['No.'] = i + 1);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Zona Norte');
    ws.columns = [
        { header: 'No.', key: 'No.', width: 6 },
        { header: 'Documento', key: 'Documento', width: 20 },
        { header: 'Nombre Trabajador', key: 'Nombre Trabajador', width: 30 },
        { header: 'Teléfono', key: 'Teléfono', width: 15 },
        { header: 'Celular', key: 'Celular', width: 15 },
        { header: 'Corrreo', key: 'Corrreo', width: 30 },
        { header: 'Estado', key: 'Estado', width: 15 },
        { header: 'Municipio Nac', key: 'Municipio Nac', width: 18 },
        { header: 'Nit Empresa', key: 'Nit Empresa', width: 15 },
        { header: 'Nombre Empresa', key: 'Nombre Empresa', width: 25 },
        { header: 'Municipio Empresa', key: 'Municipio Empresa', width: 20 },
        { header: 'Tel. Local / Ppal', key: 'Tel. Local / Ppal', width: 18 },
        { header: 'Cel. Local / Ppal', key: 'Cel. Local / Ppal', width: 18 },
        { header: 'Dirección Residencia', key: 'Dirección Residencia', width: 30 }
    ];
    output.forEach(r => ws.addRow(r));
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'ZonaNorte.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// DESCARGAR MSJ TEXTO Y GMAIL (documento enmascarado, sin municipio)
// ============================================
async function downloadMsjTextoGmail() {
    if (dataPosteriorTrabajadores.length === 0) {
        alert('No hay datos de trabajadores. Primero procesa el archivo.');
        return;
    }
    const maskDoc = (doc) => {
        const d = String(doc).trim().replace(/['’]/g, '');
        if (d.length <= 3) return d;
        return '******' + d.slice(-3);
    };
    const output = dataPosteriorTrabajadores.map(t => ({
        'Documento': maskDoc(t['Documento']),
        'Celular': t['Celular'],
        'Corrreo': t['Corrreo']
    }));
    output.sort((a, b) => a['Documento'].localeCompare(b['Documento'], 'es')); // Ordena por documento (aunque el usuario pidió alfabético por nombre, pero el ejemplo muestra por documento; sin embargo, podemos ordenar por nombre si está disponible. El ejemplo tiene solo documento y celular/correo, así que ordenamos por documento)
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Msj Texto y Gmail');
    ws.columns = [
        { header: 'Documento', key: 'Documento', width: 20 },
        { header: 'Celular', key: 'Celular', width: 15 },
        { header: 'Corrreo', key: 'Corrreo', width: 30 }
    ];
    output.forEach(r => ws.addRow(r));
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'MsjTextoGmail.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// DESCARGAR PÁGINA WEB (documento enmascarado, solo doc y nombre)
// ============================================
async function downloadPaginaWeb() {
    if (dataPosteriorTrabajadores.length === 0) {
        alert('No hay datos de trabajadores. Primero procesa el archivo.');
        return;
    }
    const maskDoc = (doc) => {
        const d = String(doc).trim().replace(/['’]/g, '');
        if (d.length <= 3) return d;
        return '******' + d.slice(-3);
    };
    const output = dataPosteriorTrabajadores.map(t => ({
        'Documento': maskDoc(t['Documento']),
        'Nombre Trabajador': t['Nombre Trabajador']
    }));
    output.sort((a, b) => a['Nombre Trabajador'].localeCompare(b['Nombre Trabajador'], 'es'));
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Página Web');
    ws.columns = [
        { header: 'Documento', key: 'Documento', width: 20 },
        { header: 'Nombre Trabajador', key: 'Nombre Trabajador', width: 30 }
    ];
    output.forEach(r => ws.addRow(r));
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'PaginaWeb.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// LIMPIAR TODO
// ============================================
function clearPosterior() {
    dataPosteriorOriginal = [];
    filteredDataPosteriorOriginal = [];
    headersPosteriorOriginal = [];
    dataPosteriorTrabajadores = [];

    const i1 = document.getElementById('fileInputPosteriorOriginal');
    if (i1) i1.value = '';
    const i2 = document.getElementById('fileInputPosteriorTrabajadores');
    if (i2) i2.value = '';

    const z1 = document.getElementById('zonaPosteriorOriginal');
    if (z1) {
        z1.classList.remove('file-loaded', 'dragover');
        const t = z1.querySelector('.drop-title');
        if (t) t.innerHTML = 'Sube el archivo de prescripción posterior (Estructura: Tipo Documento, Documento, ...)';
    }

    const z2 = document.getElementById('zonaPosteriorTrabajadores');
    if (z2) {
        z2.classList.remove('file-loaded', 'dragover');
        const t = z2.querySelector('.drop-title');
        if (t) t.innerHTML = 'Sube el archivo de trabajadores devuelto por la app';
    }

    const b1 = document.getElementById('btnDescargarExcelPosteriorOriginal');
    if (b1) b1.style.display = 'none';
    const b2 = document.getElementById('btnDescargarPosteriorFinal');
    if (b2) b2.style.display = 'none';
    const botonesExtra = document.getElementById('botonesPosterioresAdicionales');
    if (botonesExtra) botonesExtra.style.display = 'none';
}