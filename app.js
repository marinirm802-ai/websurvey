// Global State
let currentDataset = [];
let filteredDataset = [];
let currentPage = 1;
const rowsPerPage = 10;
let pendingUploadedData = [];

// Chart Instances
let spatialChartInstance = null;
let elevationChartInstance = null;

// Palette Warna per Kode Point
const CODE_COLORS = {
    'C': { bg: '#2563eb', border: '#1d4ed8', name: 'Crest (C)' },
    'T': { bg: '#10b981', border: '#047857', name: 'Toe (T)' },
    'SH': { bg: '#f59e0b', border: '#b45309', name: 'Stripping/Spot (SH)' },
    'DEFAULT': { bg: '#6b7280', border: '#374151', name: 'Lainnya' }
};

// Data Sampel Awal (Menyerupai WK_260626.xlsx)
const mockInitialData = [
    { point_id: 1, type: 2, northing: 9917186.171, easting: 509658.828, elevation: 98.117, code: 'C' },
    { point_id: 2, type: 2, northing: 9917194.116, easting: 509660.371, elevation: 94.464, code: 'C' },
    { point_id: 3, type: 2, northing: 9917195.455, easting: 509657.244, elevation: 94.837, code: 'C' },
    { point_id: 4, type: 2, northing: 9917192.629, easting: 509650.825, elevation: 94.924, code: 'C' },
    { point_id: 5, type: 2, northing: 9917189.839, easting: 509642.824, elevation: 95.209, code: 'C' },
    { point_id: 71, type: 3, northing: 9917250.100, easting: 509610.500, elevation: 91.116, code: 'T' },
    { point_id: 72, type: 3, northing: 9917252.300, easting: 509612.200, elevation: 92.400, code: 'T' },
    { point_id: 162, type: 8, northing: 9917300.500, easting: 509710.000, elevation: 86.572, code: 'SH' },
    { point_id: 163, type: 8, northing: 9917305.200, easting: 509715.100, elevation: 87.200, code: 'SH' }
];

// Initialize Dashboard
document.addEventListener('DOMContentLoaded', () => {
    currentDataset = [...mockInitialData];
    initDashboard();
    setupEventListeners();
});

function initDashboard() {
    updateFiltersOptions();
    applyFiltersAndRender();
}

// -------------------------------------------------------------
// EVENT LISTENERS & EXCEL UPLOADER LOGIC
// -------------------------------------------------------------
function setupEventListeners() {
    const fileInput = document.getElementById('excelFileInput');
    fileInput.addEventListener('change', handleFileUpload);

    document.getElementById('btnReplaceData').addEventListener('click', () => {
        currentDataset = [...pendingUploadedData];
        hideUploadModeBar();
        initDashboard();
    });

    document.getElementById('btnAppendData').addEventListener('click', () => {
        currentDataset = [...currentDataset, ...pendingUploadedData];
        hideUploadModeBar();
        initDashboard();
    });

    document.getElementById('searchInput').addEventListener('input', applyFiltersAndRender);
    document.getElementById('filterCodeSelect').addEventListener('change', applyFiltersAndRender);
    document.getElementById('btnExportCSV').addEventListener('click', exportToCSV);
}

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('fileNameDisplay').textContent = file.name;

    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert Sheet to JSON Array
        const rawJson = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        pendingUploadedData = parseRawSurveyData(rawJson);

        if (pendingUploadedData.length > 0) {
            document.getElementById('loadedInfoText').textContent = `Terbaca ${pendingUploadedData.length} titik dari file ${file.name}.`;
            document.getElementById('uploadModeBar').classList.remove('hidden');
        } else {
            alert('Format data Excel tidak sesuai. Pastikan data berisi koordinat Northing, Easting, Elevasi, dan Kode.');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseRawSurveyData(rawRows) {
    const parsed = [];
    
    // Auto Detect Header atau Data Tanpa Header (persis format WK_260626.xlsx)
    rawRows.forEach((row, idx) => {
        if (!row || row.length < 4) return;

        // Skip jika baris berisi teks header (bukan angka di kolom koordinat)
        if (isNaN(parseFloat(row[1])) || isNaN(parseFloat(row[2]))) return;

        // Mendukung format WK_260626: [Type, Northing, Easting, Elevation, Code, Point_ID]
        const type = parseInt(row[0]) || 0;
        const northing = parseFloat(row[1]);
        const easting = parseFloat(row[2]);
        const elevation = parseFloat(row[3]);
        const code = String(row[4] || 'SH').trim().toUpperCase();
        const point_id = row[5] ? parseInt(row[5]) : (idx + 1);

        if (!isNaN(northing) && !isNaN(easting) && !isNaN(elevation)) {
            parsed.push({
                point_id: point_id,
                type: type,
                northing: northing,
                easting: easting,
                elevation: elevation,
                code: code
            });
        }
    });

    return parsed;
}

function hideUploadModeBar() {
    document.getElementById('uploadModeBar').classList.add('hidden');
}

// -------------------------------------------------------------
// FILTER & RENDER DISPATCHER
// -------------------------------------------------------------
function applyFiltersAndRender() {
    const searchValue = document.getElementById('searchInput').value.toLowerCase();
    const selectedCode = document.getElementById('filterCodeSelect').value;

    filteredDataset = currentDataset.filter(item => {
        const matchesSearch = String(item.point_id).toLowerCase().includes(searchValue) ||
                              item.code.toLowerCase().includes(searchValue);
        const matchesCode = selectedCode === 'ALL' || item.code === selectedCode;
        return matchesSearch && matchesCode;
    });

    currentPage = 1;

    renderKPIs();
    renderCharts();
    renderTable();
}

function updateFiltersOptions() {
    const select = document.getElementById('filterCodeSelect');
    const uniqueCodes = [...new Set(currentDataset.map(d => d.code))];
    
    select.innerHTML = '<option value="ALL">Semua Kode</option>';
    uniqueCodes.forEach(code => {
        select.innerHTML += `<option value="${code}">${code}</option>`;
    });
}

// -------------------------------------------------------------
// METRIK KPI & STATISTIK
// -------------------------------------------------------------
function renderKPIs() {
    const data = filteredDataset;
    const totalPoints = data.length;

    document.getElementById('kpiTotalPoints').textContent = totalPoints.toLocaleString('id-ID');

    if (totalPoints === 0) {
        document.getElementById('kpiElevRange').textContent = '-';
        document.getElementById('kpiAvgElev').textContent = '0';
        document.getElementById('kpiCodeDistribution').textContent = '-';
        document.getElementById('statMinZ').textContent = '-';
        document.getElementById('statMaxZ').textContent = '-';
        document.getElementById('statStdDev').textContent = '-';
        return;
    }

    const elevations = data.map(d => d.elevation);
    const minZ = Math.min(...elevations);
    const maxZ = Math.max(...elevations);
    const avgZ = elevations.reduce((a, b) => a + b, 0) / totalPoints;

    // Standard Deviation Calculation
    const variance = elevations.reduce((a, b) => a + Math.pow(b - avgZ, 2), 0) / totalPoints;
    const stdDev = Math.sqrt(variance);

    document.getElementById('kpiElevRange').textContent = `${minZ.toFixed(2)} - ${maxZ.toFixed(2)}`;
    document.getElementById('kpiAvgElev').textContent = avgZ.toFixed(2);
    document.getElementById('statMinZ').textContent = `${minZ.toFixed(2)} m`;
    document.getElementById('statMaxZ').textContent = `${maxZ.toFixed(2)} m`;
    document.getElementById('statStdDev').textContent = stdDev.toFixed(2);

    // Count per Code
    const codeCounts = {};
    data.forEach(d => {
        codeCounts[d.code] = (codeCounts[d.code] || 0) + 1;
    });

    const distText = Object.keys(codeCounts)
        .map(code => `<span class="bg-slate-100 px-2 py-1 rounded">${code}: ${codeCounts[code]}</span>`)
        .join(' ');
    document.getElementById('kpiCodeDistribution').innerHTML = distText;
}

// -------------------------------------------------------------
// VISUALISASI CHARTJS
// -------------------------------------------------------------
function renderCharts() {
    renderSpatialScatterChart();
    renderElevationBarChart();
}

function renderSpatialScatterChart() {
    const ctx = document.getElementById('spatialScatterChart').getContext('2d');

    // Grouping Data by Code for Chart.js Datasets
    const groups = {};
    filteredDataset.forEach(pt => {
        if (!groups[pt.code]) groups[pt.code] = [];
        groups[pt.code].push({ x: pt.easting, y: pt.northing, z: pt.elevation, id: pt.point_id });
    });

    const datasets = Object.keys(groups).map(code => {
        const colorConfig = CODE_COLORS[code] || CODE_COLORS['DEFAULT'];
        return {
            label: colorConfig.name || code,
            data: groups[code],
            backgroundColor: colorConfig.bg,
            borderColor: colorConfig.border,
            borderWidth: 1,
            pointRadius: 4,
            pointHoverRadius: 6
        };
    });

    if (spatialChartInstance) spatialChartInstance.destroy();

    spatialChartInstance = new Chart(ctx, {
        type: 'scatter',
        data: { datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const raw = ctx.raw;
                            return `ID: ${raw.id} (${ctx.dataset.label}) | X: ${raw.x.toFixed(2)}, Y: ${raw.y.toFixed(2)}, Z: ${raw.z.toFixed(2)}m`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    title: { display: true, text: 'Easting (X)', font: { size: 11 } },
                    ticks: { font: { size: 10 } }
                },
                y: {
                    title: { display: true, text: 'Northing (Y)', font: { size: 11 } },
                    ticks: { font: { size: 10 } }
                }
            }
        }
    });
}

function renderElevationBarChart() {
    const ctx = document.getElementById('elevationBarChart').getContext('2d');

    // Calculate Average Elevation per Code
    const codeMap = {};
    filteredDataset.forEach(d => {
        if (!codeMap[d.code]) codeMap[d.code] = { totalZ: 0, count: 0 };
        codeMap[d.code].totalZ += d.elevation;
        codeMap[d.code].count += 1;
    });

    const labels = Object.keys(codeMap);
    const avgData = labels.map(code => (codeMap[code].totalZ / codeMap[code].count).toFixed(2));
    const bgColors = labels.map(code => (CODE_COLORS[code] ? CODE_COLORS[code].bg : CODE_COLORS['DEFAULT'].bg));

    if (elevationChartInstance) elevationChartInstance.destroy();

    elevationChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Rata-Rata Elevasi (m)',
                data: avgData,
                backgroundColor: bgColors,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: false,
                    title: { display: true, text: 'Elevasi (mdpl)', font: { size: 11 } }
                }
            }
        }
    });
}

// -------------------------------------------------------------
// TABEL DATA & PAGINATION
// -------------------------------------------------------------
function renderTable() {
    const tbody = document.getElementById('surveyTableBody');
    tbody.innerHTML = '';

    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = filteredDataset.slice(startIdx, endIdx);

    if (pageData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" class="text-center py-6 text-gray-400">Tidak ada data uji petik yang cocok</td></tr>`;
        renderPaginationControls(0);
        return;
    }

    pageData.forEach(item => {
        const codeClass = `badge-${item.code.toLowerCase()}`;
        const badgeHTML = `<span class="badge-code ${codeClass}">${item.code}</span>`;

        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="px-4 py-2.5 font-medium text-gray-900">${item.point_id}</td>
            <td class="px-4 py-2.5">${badgeHTML}</td>
            <td class="px-4 py-2.5 text-right font-mono">${item.northing.toFixed(3)}</td>
            <td class="px-4 py-2.5 text-right font-mono">${item.easting.toFixed(3)}</td>
            <td class="px-4 py-2.5 text-right font-mono font-semibold text-emerald-700">${item.elevation.toFixed(3)}</td>
        `;
        tbody.appendChild(row);
    });

    renderPaginationControls(filteredDataset.length);
}

function renderPaginationControls(totalItems) {
    const totalPages = Math.ceil(totalItems / rowsPerPage) || 1;
    const container = document.getElementById('paginationControls');
    const info = document.getElementById('tablePaginationInfo');

    const start = totalItems === 0 ? 0 : (currentPage - 1) * rowsPerPage + 1;
    const end = Math.min(currentPage * rowsPerPage, totalItems);
    info.textContent = `Menampilkan ${start} - ${end} dari ${totalItems} titik`;

    container.innerHTML = `
        <button ${currentPage === 1 ? 'disabled' : ''} onclick="changePage(${currentPage - 1})" class="px-2.5 py-1 bg-white border border-gray-300 rounded hover:bg-gray-100 disabled:opacity-50 transition">Prev</button>
        <span class="px-2 py-1 font-semibold">${currentPage} / ${totalPages}</span>
        <button ${currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${currentPage + 1})" class="px-2.5 py-1 bg-white border border-gray-300 rounded hover:bg-gray-100 disabled:opacity-50 transition">Next</button>
    `;
}

function changePage(newPage) {
    currentPage = newPage;
    renderTable();
}

// -------------------------------------------------------------
// EXPORT TO CSV
// -------------------------------------------------------------
function exportToCSV() {
    if (filteredDataset.length === 0) {
        alert('Tidak ada data untuk di-export');
        return;
    }

    const headers = ['Point_ID', 'Type', 'Northing', 'Easting', 'Elevation', 'Code'];
    const csvRows = [headers.join(',')];

    filteredDataset.forEach(row => {
        csvRows.push([row.point_id, row.type, row.northing, row.easting, row.elevation, row.code].join(','));
    });

    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.setAttribute('href', url);
    a.setAttribute('download', `Uji_Petik_Topografi_${new Date().toISOString().slice(0, 10)}.csv`);
    a.click();
}
