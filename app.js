// Global application state for Pertamina Construction Dashboard
const AppState = {
    currentSection: 'overview',
    selectedProject: 'all',
    schedulerSelectedProject: 'all',
    safetySelectedProject: 'all',
    data: {
        projects: [],
        sCurveData: [],
        safety: [],
        issues: [],
        plans: [],
        documents: [],
        permits: [],
        uploadHistory: [],
        generatedReports: []
    },
    charts: {}
};

// Project ID mapping for consistency across all sheets (seed only)
const PROJECT_ID_MAPPING = {
    'PROJ001': 'Petani Substation',
    'PROJ002': 'Menggala Substation', 
    'PROJ003': 'Nella Substation',
    'PROJ004': 'Bangko Substation',
    'PROJ005': 'Balam SS',
    'PROJ006': 'Sintong SS',
    'PROJ007': 'OKB Substation'
};

// =========================
// Dynamic maps (rebuilt from sheet Projects)
// =========================
let _idToName = {};
let _nameToId = {};

function buildProjectMaps() {
    _idToName = { ...PROJECT_ID_MAPPING };
    _nameToId = Object.fromEntries(Object.entries(_idToName).map(([id, name]) => [name, id]));

    (AppState?.data?.projects || []).forEach(p => {
        const id = String(p.Project_ID || '').trim();
        const nm = String(p.Project_Name || '').trim();
        if (id && nm) {
            _idToName[id] = nm;
            _nameToId[nm] = id;
        }
    });
}

function getProjectName(projectId) {
    const id = projectId || '';
    return _idToName[id] || id || 'Unknown Project';
}

function getProjectId(projectName) {
    if (!projectName || projectName === 'all') return null;
    return _nameToId[projectName] || null;
}

function getProjectIdFromData(projectName) {
    if (!projectName) return null;
    const nm = String(projectName).trim().toLowerCase();
    const found = (AppState?.data?.projects || []).find(
        x => String(x.Project_Name || '').trim().toLowerCase() === nm
    );
    return found ? found.Project_ID : null;
}

function resolveProjectId(row) {
    return row.Project_ID
        || getProjectId(row.Project_Name)
        || getProjectIdFromData(row.Project_Name)
        || 'PROJ_UNKNOWN';
}

function getStatusClass(status) {
    if (!status) return 'info';
    const statusLower = status.toLowerCase().replace(/\s+/g, '-');
    if (statusLower.includes('open') || statusLower.includes('planned')) return 'info';
    if (statusLower.includes('progress')) return 'warning';
    if (statusLower.includes('completed') || statusLower.includes('closed')) return 'success';
    return 'info';
}

// =========================
// Initial sample data
// =========================
const initialData = {
    projects: [
        {Date: "2024-01-15", Project_ID: "PROJ001", Project_Name: "Petani Substation", Progress_Percent: 75.5, Budget_Used_Percent: 72.3, Budget_Total: 15000000, Labor_Hours: 1200, Equipment_Hours: 450, Safety_Score: 95, Quality_Score: 88, Weather_Impact: "Minor", Delay_Days: 2},
        {Date: "2024-01-15", Project_ID: "PROJ002", Project_Name: "Menggala Substation", Progress_Percent: 68.2, Budget_Used_Percent: 65.1, Budget_Total: 12500000, Labor_Hours: 1500, Equipment_Hours: 500, Safety_Score: 92, Quality_Score: 90, Weather_Impact: "None", Delay_Days: 0},
        {Date: "2024-01-16", Project_ID: "PROJ003", Project_Name: "Nella Substation", Progress_Percent: 82.7, Budget_Used_Percent: 79.4, Budget_Total: 8500000, Labor_Hours: 980, Equipment_Hours: 320, Safety_Score: 97, Quality_Score: 93, Weather_Impact: "None", Delay_Days: 0},
        {Date: "2024-01-17", Project_ID: "PROJ004", Project_Name: "Bangko Substation", Progress_Percent: 45.3, Budget_Used_Percent: 43.8, Budget_Total: 6200000, Labor_Hours: 750, Equipment_Hours: 280, Safety_Score: 89, Quality_Score: 85, Weather_Impact: "Moderate", Delay_Days: 3},
        {Date: "2024-01-18", Project_ID: "PROJ005", Project_Name: "Balam SS", Progress_Percent: 91.2, Budget_Used_Percent: 88.5, Budget_Total: 4500000, Labor_Hours: 600, Equipment_Hours: 200, Safety_Score: 98, Quality_Score: 95, Weather_Impact: "None", Delay_Days: 0},
        {Date: "2024-01-19", Project_ID: "PROJ006", Project_Name: "Sintong SS", Progress_Percent: 55.8, Budget_Used_Percent: 52.4, Budget_Total: 7800000, Labor_Hours: 900, Equipment_Hours: 350, Safety_Score: 91, Quality_Score: 87, Weather_Impact: "Minor", Delay_Days: 1},
        {Date: "2024-01-20", Project_ID: "PROJ007", Project_Name: "OKB Substation", Progress_Percent: 38.9, Budget_Used_Percent: 35.2, Budget_Total: 9200000, Labor_Hours: 800, Equipment_Hours: 300, Safety_Score: 94, Quality_Score: 89, Weather_Impact: "None", Delay_Days: 0}
    ],
    sCurveData: [
        {Project_ID: "PROJ001", Project_Name: "Petani Substation", Week_Month_Label: "W1", Date_Reference: "2024-01-01", Planned_Progress_Pct: 0, Actual_Progress_Pct: 0, Cumulative_Planned: 0, Cumulative_Actual: 0},
        {Project_ID: "PROJ001", Project_Name: "Petani Substation", Week_Month_Label: "W2", Date_Reference: "2024-01-07", Planned_Progress_Pct: 15, Actual_Progress_Pct: 10, Cumulative_Planned: 15, Cumulative_Actual: 10},
        {Project_ID: "PROJ001", Project_Name: "Petani Substation", Week_Month_Label: "W3", Date_Reference: "2024-01-14", Planned_Progress_Pct: 35, Actual_Progress_Pct: 28, Cumulative_Planned: 35, Cumulative_Actual: 28},
        {Project_ID: "PROJ001", Project_Name: "Petani Substation", Week_Month_Label: "W4", Date_Reference: "2024-01-21", Planned_Progress_Pct: 60, Actual_Progress_Pct: 55, Cumulative_Planned: 60, Cumulative_Actual: 55},
        {Project_ID: "PROJ002", Project_Name: "Menggala Substation", Week_Month_Label: "W1", Date_Reference: "2024-01-01", Planned_Progress_Pct: 0, Actual_Progress_Pct: 0, Cumulative_Planned: 0, Cumulative_Actual: 0},
        {Project_ID: "PROJ002", Project_Name: "Menggala Substation", Week_Month_Label: "W2", Date_Reference: "2024-01-07", Planned_Progress_Pct: 20, Actual_Progress_Pct: 18, Cumulative_Planned: 20, Cumulative_Actual: 18},
        {Project_ID: "PROJ002", Project_Name: "Menggala Substation", Week_Month_Label: "W3", Date_Reference: "2024-01-14", Planned_Progress_Pct: 45, Actual_Progress_Pct: 40, Cumulative_Planned: 45, Cumulative_Actual: 40},
        {Project_ID: "PROJ002", Project_Name: "Menggala Substation", Week_Month_Label: "W4", Date_Reference: "2024-01-21", Planned_Progress_Pct: 70, Actual_Progress_Pct: 68, Cumulative_Planned: 70, Cumulative_Actual: 68}
    ],
    safety: [
        {Date: "2024-01-15", Project_ID: "PROJ001", Project_Name: "Petani Substation", Safety_Score: 95, Safe_Man_Hours: 1180, Total_Manpower: 45, Fatal_Accidents: 0, Lost_Time_Injuries: 0, Medical_Treatment_Cases: 1, First_Aid_Cases: 2, Near_Miss_Events: 8, Unsafe_Acts: 25, Unsafe_Conditions: 40},
        {Date: "2024-01-16", Project_ID: "PROJ002", Project_Name: "Menggala Substation", Safety_Score: 92, Safe_Man_Hours: 1480, Total_Manpower: 62, Fatal_Accidents: 0, Lost_Time_Injuries: 0, Medical_Treatment_Cases: 1, First_Aid_Cases: 3, Near_Miss_Events: 6, Unsafe_Acts: 30, Unsafe_Conditions: 45},
        {Date: "2024-01-17", Project_ID: "PROJ003", Project_Name: "Nella Substation", Safety_Score: 97, Safe_Man_Hours: 970, Total_Manpower: 38, Fatal_Accidents: 0, Lost_Time_Injuries: 0, Medical_Treatment_Cases: 0, First_Aid_Cases: 1, Near_Miss_Events: 5, Unsafe_Acts: 15, Unsafe_Conditions: 25},
        {Date: "2024-01-18", Project_ID: "PROJ004", Project_Name: "Bangko Substation", Safety_Score: 89, Safe_Man_Hours: 720, Total_Manpower: 35, Fatal_Accidents: 0, Lost_Time_Injuries: 1, Medical_Treatment_Cases: 1, First_Aid_Cases: 2, Near_Miss_Events: 6, Unsafe_Acts: 40, Unsafe_Conditions: 60}
    ],
    issues: [
        {Issue_ID: "ISS-001", Project_ID: "PROJ001", Project_Name: "Petani Substation", Issue_Title: "Material Delivery Delay", Priority: "High", Status: "Open", Assigned_To: "Procurement Manager", Due_Date: "2024-01-25"},
        {Issue_ID: "ISS-002", Project_ID: "PROJ002", Project_Name: "Menggala Substation", Issue_Title: "Weather Impact Assessment", Priority: "Medium", Status: "In Progress", Assigned_To: "Site Manager", Due_Date: "2024-01-20"},
        {Issue_ID: "ISS-003", Project_ID: "PROJ003", Project_Name: "Nella Substation", Issue_Title: "Equipment Malfunction", Priority: "High", Status: "Open", Assigned_To: "Technical Team", Due_Date: "2024-01-22"},
        {Issue_ID: "ISS-004", Project_ID: "PROJ004", Project_Name: "Bangko Substation", Issue_Title: "Safety Compliance Review", Priority: "Medium", Status: "Closed", Assigned_To: "Safety Officer", Due_Date: "2024-01-18"}
    ],
    plans: [
        {Plan_ID: "PLAN-001", Project_ID: "PROJ001", Project_Name: "Petani Substation", Plan_Title: "Final Inspection Schedule", Priority: "High", Status: "Planned", Assigned_To: "QC Manager", Start_Date: "2024-01-22", End_Date: "2024-01-25"},
        {Plan_ID: "PLAN-002", Project_ID: "PROJ002", Project_Name: "Menggala Substation", Plan_Title: "Safety Training Program", Priority: "Medium", Status: "In Progress", Assigned_To: "Safety Officer", Start_Date: "2024-01-18", End_Date: "2024-01-22"},
        {Plan_ID: "PLAN-003", Project_ID: "PROJ003", Project_Name: "Nella Substation", Plan_Title: "Equipment Testing", Priority: "High", Status: "Planned", Assigned_To: "Technical Team", Start_Date: "2024-01-20", End_Date: "2024-01-24"},
        {Plan_ID: "PLAN-004", Project_ID: "PROJ004", Project_Name: "Bangko Substation", Plan_Title: "Progress Review Meeting", Priority: "Low", Status: "Completed", Assigned_To: "Project Manager", Start_Date: "2024-01-16", End_Date: "2024-01-17"}
    ],
    documents: [
        {Document_ID: "DOC-001", Document_Name: "Petani Substation Contract", Document_Type: "Contract", Project_ID: "PROJ001", Project_Name: "Petani Substation", Created_Date: "2024-01-10", Size_KB: 2048, Keywords: "contract, construction, pertamina", Created_By: "Contract Manager"},
        {Document_ID: "DOC-002", Document_Name: "Monthly Safety Report", Document_Type: "Report", Project_ID: "PROJ002", Project_Name: "Menggala Substation", Created_Date: "2024-01-12", Size_KB: 1536, Keywords: "safety, monthly, report, inspection", Created_By: "Safety Manager"},
        {Document_ID: "DOC-003", Document_Name: "Technical Specification", Document_Type: "Specification", Project_ID: "PROJ003", Project_Name: "Nella Substation", Created_Date: "2024-01-08", Size_KB: 3072, Keywords: "technical, specification, electrical", Created_By: "Technical Lead"},
        {Document_ID: "DOC-004", Document_Name: "Progress Drawing Rev-A", Document_Type: "Drawing", Project_ID: "PROJ001", Project_Name: "Petani Substation", Created_Date: "2024-01-15", Size_KB: 4096, Keywords: "drawing, progress, revision", Created_By: "Design Engineer"},
        {Document_ID: "DOC-005", Document_Name: "Safety Manual", Document_Type: "Manual", Project_ID: "PROJ004", Project_Name: "Bangko Substation", Created_Date: "2024-01-14", Size_KB: 2560, Keywords: "safety, manual, procedures", Created_By: "Safety Team"}
    ],
    permits: [
        {PTW_ID: "PTW-2024-001", Date_Issued: "2024-01-15", Permit_Type: "Hot Work", Contractor: "PT Konstruksi ABC", Area: "Petani Substation", Activity: "Welding", Status: "Open", Project_ID: "PROJ001", Project_Name: "Petani Substation"},
        {PTW_ID: "PTW-2024-002", Date_Issued: "2024-01-16", Permit_Type: "Confined Space", Contractor: "PT Engineering XYZ", Area: "Menggala Substation", Activity: "Cable Installation", Status: "Open", Project_ID: "PROJ002", Project_Name: "Menggala Substation"},
        {PTW_ID: "PTW-2024-003", Date_Issued: "2024-01-17", Permit_Type: "Working at Height", Contractor: "PT Maintenance DEF", Area: "Nella Substation", Activity: "Tower Installation", Status: "Closed", Project_ID: "PROJ003", Project_Name: "Nella Substation"}
    ]
};

// Initialize application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    console.log('Initializing Pertamina Construction Dashboard...');
    AppState.data = { ...initialData };
    buildProjectMaps();

    setupNavigation();
    initializeOverview();
    initializeExcelManagement();
    initializeScheduler();
    initializePermits();
    initializeDocuments();
    initializeReports();

    initializeChatbot();
    setupProjectFilters();

    showNotification('Dashboard Pertamina berhasil dimuat! Upload file Excel untuk data terlengkap.', 'success');
}

// =========================
// Navigation
// =========================
function setupNavigation() {
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const targetSection = this.getAttribute('data-section');
            switchSection(targetSection);
            navLinks.forEach(nav => nav.classList.remove('active'));
            this.classList.add('active');
        });
    });
}

function switchSection(sectionName) {
    const sections = document.querySelectorAll('.content-section');
    sections.forEach(section => section.classList.remove('active'));
    const targetSection = document.getElementById(sectionName);
    if (targetSection) {
        targetSection.classList.add('active');
        AppState.currentSection = sectionName;
        if (sectionName === 'overview') updateOverviewContent();
        else if (sectionName === 'scheduler') updateSchedulerContent();
        else if (sectionName === 'documents') renderDocuments();
        else if (sectionName === 'permits') updatePermitsContent();
    }
}

// =========================
// Overview
// =========================
function initializeOverview() {
    updateKPICards();
    initializeProgressChart();
    initializeSafetyTrendChart();
}

function updateOverviewContent() {
    updateKPICards();
    initializeProgressChart();
}

function updateKPICards() {
    const allProjects = AppState.data.projects;
    const activePermits = AppState.data.permits.filter(p => p.Status === 'Open').length;

    document.getElementById('totalProjects').textContent = allProjects.length;
    document.getElementById('activePermits').textContent = activePermits;

    if (allProjects.length > 0) {
        const avgSafety = Math.round(allProjects.reduce((sum, p) => sum + p.Safety_Score, 0) / allProjects.length);
        const avgBudget = Math.round(allProjects.reduce((sum, p) => sum + p.Budget_Used_Percent, 0) / allProjects.length);
        document.getElementById('safetyScore').textContent = avgSafety + '%';
        document.getElementById('budgetPerformance').textContent = avgBudget + '%';
    }
}

function initializeProgressChart() {
    const ctx = document.getElementById('progressChart').getContext('2d');
    const filteredProjects = getFilteredProjects();

    if (AppState.charts.progress) AppState.charts.progress.destroy();

    if (filteredProjects.length === 0 && AppState.selectedProject !== 'all') {
        document.getElementById('emptyState').style.display = 'block';
        return;
    } else {
        document.getElementById('emptyState').style.display = 'none';
    }

    AppState.charts.progress = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: filteredProjects.map(p => p.Project_Name.replace(' Substation', '').replace(' SS', '')),
            datasets: [{
                label: 'Progress (%)',
                data: filteredProjects.map(p => p.Progress_Percent),
                backgroundColor: '#1FB8CD',
                borderColor: '#1FB8CD',
                borderWidth: 1
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } } }
    });
}

function initializeSafetyTrendChart() {
    const ctx = document.getElementById('safetyChart').getContext('2d');
    const allProjects = AppState.data.projects;

    if (AppState.charts.safety) AppState.charts.safety.destroy();

    AppState.charts.safety = new Chart(ctx, {
        type: 'line',
        data: {
            labels: allProjects.map(p => p.Project_Name.replace(' Substation', '').replace(' SS', '')),
            datasets: [{
                label: 'Safety Score (%)',
                data: allProjects.map(p => p.Safety_Score),
                borderColor: '#FFC185',
                backgroundColor: 'rgba(255, 193, 133, 0.1)',
                fill: true,
                tension: 0.4
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } } }
    });
}

// =========================
// Excel Management
// =========================
function initializeExcelManagement() {
    setupFileUpload();
    updateUploadHistory();
}

function setupFileUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', e => { e.preventDefault(); dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', e => {
        e.preventDefault(); dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFileUpload(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', e => { if (e.target.files.length > 0) handleFileUpload(e.target.files[0]); });
}

function handleFileUpload(file) {
    const validTypes = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    if (!validTypes.includes(fileExtension)) { showValidationResults('error', 'File harus berformat Excel (.xlsx atau .xls)'); return; }

    const progressContainer = document.getElementById('uploadProgress');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    progressContainer.style.display = 'block';

    if (typeof XLSX === 'undefined') { showValidationResults('error', 'Excel processing library not loaded. Please refresh the page.'); progressContainer.style.display = 'none'; return; }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            processExcelWorkbook(workbook, file);
        } catch (error) {
            console.error('Error processing Excel file:', error);
            showValidationResults('error', 'Error processing Excel file: ' + error.message);
            progressContainer.style.display = 'none';
        }
    };

    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 15; if (progress > 90) progress = 90;
        progressBar.style.width = progress + '%';
        progressText.textContent = `Processing... ${Math.round(progress)}%`;
        if (progress >= 90) clearInterval(interval);
    }, 100);

    reader.readAsArrayBuffer(file);
}

function processExcelWorkbook(workbook, file) {
    let totalRecords = 0;
    let processedSheets = [];
    const raw = {};
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        if (jsonData.length > 0) raw[sheetName.toLowerCase()] = jsonData;
    });

    const projectsKey = Object.keys(raw).find(k => k.includes('project'));
    if (projectsKey) {
        AppState.data.projects = raw[projectsKey].map(row => ({
            ...row,
            Project_ID: row.Project_ID || getProjectId(row.Project_Name) || ''
        }));
        totalRecords += raw[projectsKey].length;
        processedSheets.push('Projects');
        buildProjectMaps();
    }

    const apply = (keyPredicate, target, label) => {
        const key = Object.keys(raw).find(keyPredicate);
        if (!key) return;
        AppState.data[target] = raw[key].map(row => ({ ...row, Project_ID: resolveProjectId(row) }));
        totalRecords += raw[key].length;
        processedSheets.push(label);
    };

    apply(k => k.includes('s_curve') || k.includes('scurve'), 'sCurveData', 'S_Curve');
    apply(k => k.includes('safety'), 'safety', 'Safety');
    apply(k => k.includes('issue'), 'issues', 'Issues');
    apply(k => k.includes('plan'), 'plans', 'Plans');
    apply(k => k.includes('document'), 'documents', 'Documents');
    apply(k => k.includes('permit'), 'permits', 'Permits');

    completeExcelUpload(file, totalRecords, processedSheets);
}

function completeExcelUpload(file, totalRecords, processedSheets) {
    const progressContainer = document.getElementById('uploadProgress');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');

    progressBar.style.width = '100%';
    progressText.textContent = 'Upload complete... 100%';

    setTimeout(() => {
        progressContainer.style.display = 'none';
        const uploadRecord = {
            fileName: file.name,
            uploadDate: new Date().toLocaleString('id-ID'),
            size: (file.size / (1024 * 1024)).toFixed(1) + ' MB',
            records: totalRecords,
            sheets: processedSheets.join(', '),
            status: 'Berhasil'
        };
        AppState.data.uploadHistory = AppState.data.uploadHistory || [];
        AppState.data.uploadHistory.unshift(uploadRecord);
        updateUploadHistory();

        ['sCurveData','safety','issues','plans','documents','permits'].forEach(k => {
            AppState.data[k] = (AppState.data[k] || []).filter(r => r.Project_ID && r.Project_ID !== 'PROJ_UNKNOWN');
        });

        buildProjectMaps();
        setupProjectFilters();
        setupDocumentFilters();

        updateOverviewContent();
        updateSchedulerContent();
        updatePermitsContent();
        renderDocuments();
        updateReportPreview();

        const message = `File berhasil diupload! ${totalRecords} records dari ${processedSheets.length} sheets diproses.`;
        showValidationResults('success', message);
        showNotification('Data Excel berhasil diupload dan seluruh dashboard telah diperbarui!', 'success');
    }, 500);
}

function showValidationResults(type, message) {
    const container = document.getElementById('validationResults');
    container.className = `validation-results ${type}`;
    container.textContent = message;
    container.style.display = 'block';
    setTimeout(() => { container.style.display = 'none'; }, 5000);
}

function updateUploadHistory() {
    const tbody = document.querySelector('#uploadHistoryTable tbody');
    tbody.innerHTML = '';
    (AppState.data.uploadHistory || []).forEach(record => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${record.fileName}</td>
            <td>${record.uploadDate}</td>
            <td>${record.size}</td>
            <td>${record.records}</td>
            <td>${record.sheets || 'N/A'}</td>
            <td><span class="status status--success">${record.status}</span></td>`;
        tbody.appendChild(row);
    });
}

function downloadCurrentData() {
    if (typeof XLSX === 'undefined') { showNotification('Excel processing library not loaded. Please refresh the page.', 'error'); return; }
    const wb = XLSX.utils.book_new();
    if (AppState.data.projects.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.projects), 'Projects');
    if (AppState.data.sCurveData.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.sCurveData), 'S_Curve');
    if (AppState.data.safety.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.safety), 'Safety');
    if (AppState.data.issues.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.issues), 'Issues');
    if (AppState.data.plans.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.plans), 'Plans');
    if (AppState.data.documents.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.documents), 'Documents');
    if (AppState.data.permits.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(AppState.data.permits), 'Permits');
    XLSX.writeFile(wb, `Pertamina_Dashboard_Data_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('Data saat ini berhasil didownload!', 'success');
}

// =========================
// Scheduler & S-Curve
// (akan lanjut di Part B)
// =========================

// =========================
// Documents
// =========================
function initializeDocuments() {
    renderDocuments();
    setupDocumentSearch();
    setupDocumentFilters();
}

function setupDocumentSearch() {
    const searchInput = document.getElementById('documentSearch');
    if (searchInput) searchInput.addEventListener('input', renderDocuments);
}

function setupDocumentFilters() {
    const typeFilter = document.getElementById('documentTypeFilter');
    const projectFilter = document.getElementById('documentProjectFilter');

    if (typeFilter) typeFilter.addEventListener('change', renderDocuments);
    if (projectFilter) {
        projectFilter.innerHTML = '<option value="all">Semua Project</option>';
        const names = [...new Set((AppState.data.projects || []).map(p => p.Project_Name))].sort((a,b)=>a.localeCompare(b));
        names.forEach(n => { const o=document.createElement('option'); o.value=n; o.textContent=n; projectFilter.appendChild(o); });
        projectFilter.addEventListener('change', renderDocuments);
    }
}

function renderDocuments() {
    const grid = document.getElementById('documentsGrid');
    const resultsInfo = document.getElementById('searchResults');
    grid.innerHTML = '';

    const documents = AppState.data.documents || [];
    let filteredDocuments = [...documents];

    const searchQuery = (document.getElementById('documentSearch')?.value || '').toLowerCase();
    if (searchQuery) {
        filteredDocuments = filteredDocuments.filter(doc => 
            (doc.Document_Name || '').toLowerCase().includes(searchQuery) ||
            (doc.Document_Type || '').toLowerCase().includes(searchQuery) ||
            (doc.Keywords || '').toLowerCase().includes(searchQuery) ||
            (doc.Project_Name || '').toLowerCase().includes(searchQuery) ||
            (doc.Created_By || '').toLowerCase().includes(searchQuery)
        );
    }

    const typeFilterV = document.getElementById('documentTypeFilter')?.value;
    if (typeFilterV && typeFilterV !== 'all') {
        filteredDocuments = filteredDocuments.filter(doc => doc.Document_Type === typeFilterV);
    }

    const projectFilterV = document.getElementById('documentProjectFilter')?.value;
    if (projectFilterV && projectFilterV !== 'all') {
        const selectedProjectId = getProjectId(projectFilterV);
        if (selectedProjectId) filteredDocuments = filteredDocuments.filter(doc => doc.Project_ID === selectedProjectId);
    }

    resultsInfo.textContent = `Menampilkan ${filteredDocuments.length} dari ${documents.length} dokumen`;

    if (filteredDocuments.length === 0) {
        grid.innerHTML = `<div class="no-data-message"><p>Tidak ada dokumen ditemukan. ${searchQuery ? 'Coba ubah kriteria pencarian.' : 'Upload file Excel dengan sheet Documents untuk menambahkan data dokumen.'}</p></div>`;
        return;
    }

    filteredDocuments.forEach(doc => {
        const card = document.createElement('div');
        card.className = 'document-card';
        card.onclick = () => openDocumentModal(doc);
        const iconMap = { 'Contract': '📋', 'Report': '📊', 'Drawing': '📐', 'Specification': '📝', 'Manual': '📖', 'default': '📄' };
        const icon = iconMap[doc.Document_Type] || iconMap.default;
        card.innerHTML = `
            <div class="document-header">
                <div class="document-icon">${icon}</div>
                <h4 class="document-title">${doc.Document_Name || 'Unnamed Document'}</h4>
            </div>
            <div class="document-meta">
                <span>Type: ${doc.Document_Type || 'Unknown'}</span>
                <span>Size: ${((doc.Size_KB || 0) / 1024).toFixed(1)} MB</span>
                <span>Created: ${doc.Created_Date || 'Unknown'}</span>
                <span>By: ${doc.Created_By || 'Unknown'}</span>
            </div>
            <div class="document-project">${getProjectName(doc.Project_ID) || doc.Project_Name || 'No Project'}</div>
            <div class="document-keywords">${doc.Keywords || 'No keywords'}</div>`;
        grid.appendChild(card);
    });
}

function performDocumentSearch() { renderDocuments(); }

function openDocumentModal(doc) {
    const modal = document.getElementById('documentModal');
    const title = document.getElementById('modalTitle');
    const preview = document.getElementById('documentPreview');
    title.textContent = doc.Document_Name || 'Unnamed Document';
    preview.innerHTML = `
        <div style="padding: 20px; background: var(--color-bg-1); border-radius: 8px;">
            <h4>Document Information</h4>
            <p><strong>ID:</strong> ${doc.Document_ID || 'N/A'}</p>
            <p><strong>Type:</strong> ${doc.Document_Type || 'Unknown'}</p>
            <p><strong>Project:</strong> ${getProjectName(doc.Project_ID) || doc.Project_Name || 'No Project'}</p>
            <p><strong>Created:</strong> ${doc.Created_Date || 'Unknown'}</p>
            <p><strong>Size:</strong> ${((doc.Size_KB || 0) / 1024).toFixed(1)} MB</p>
            <p><strong>Keywords:</strong> ${doc.Keywords || 'No keywords'}</p>
            <p><strong>Created By:</strong> ${doc.Created_By || 'Unknown'}</p>
            <hr>
            <p style="color: var(--color-text-secondary);">This document is integrated from Excel sheet "Documents". Data can be updated by uploading a new Excel file with the Documents sheet.</p>
        </div>`;
    modal.classList.remove('hidden');
}

function closeDocumentModal() {
    const modal = document.getElementById('documentModal');
    modal.classList.add('hidden');
}
// =========================
// Scheduler & S-Curve
// =========================
function initializeScheduler() {
    initializeSCurveChart();
    updateIssuesTable();
    updatePlansTable();
}

function updateSchedulerContent() {
    initializeSCurveChart();
    updateIssuesTable();
    updatePlansTable();
}

function initializeSCurveChart() {
    const ctx = document.getElementById('sCurveChart').getContext('2d');
    const selectedProjectId = getProjectId(AppState.schedulerSelectedProject);

    if (AppState.charts.sCurve) AppState.charts.sCurve.destroy();

    let sCurveData = [...AppState.data.sCurveData];
    if (selectedProjectId && AppState.schedulerSelectedProject !== 'all') {
        sCurveData = sCurveData.filter(item => item.Project_ID === selectedProjectId);
    }

    if (sCurveData.length === 0) {
        const labels = ['W1','W2','W3','W4','Jan','Feb','Mar','Apr'];
        const plannedData = [0,15,30,45,60,75,90,100];
        const actualData  = [0,10,25,40,55,70,85,95];
        AppState.charts.sCurve = new Chart(ctx, {
            type: 'line',
            data: { labels, datasets: [
                { label: 'Planned (%)', data: plannedData, borderColor: '#1FB8CD', backgroundColor: 'rgba(31,184,205,0.1)', tension: 0.4 },
                { label: 'Actual (%)',  data: actualData,  borderColor: '#B4413C', backgroundColor: 'rgba(180,65,60,0.1)', tension: 0.4 }
            ]},
            options: { responsive:true, maintainAspectRatio:false, scales:{ y:{ beginAtZero:true, max:100 } } }
        });
    } else {
        const sorted = sCurveData.sort((a,b)=>new Date(a.Date_Reference)-new Date(b.Date_Reference));
        const labels = sorted.map(i=>i.Week_Month_Label || new Date(i.Date_Reference).toLocaleDateString('id-ID'));
        const plannedData = sorted.map(i=>Number(i.Cumulative_Planned)||0);
        const actualData  = sorted.map(i=>Number(i.Cumulative_Actual)||0);
        AppState.charts.sCurve = new Chart(ctx, {
            type: 'line',
            data: { labels, datasets: [
                { label:'Planned (%)', data:plannedData, borderColor:'#1FB8CD', backgroundColor:'rgba(31,184,205,0.1)', tension:0.4 },
                { label:'Actual (%)',  data:actualData,  borderColor:'#B4413C', backgroundColor:'rgba(180,65,60,0.1)', tension:0.4 }
            ]},
            options:{ responsive:true, maintainAspectRatio:false, scales:{ y:{ beginAtZero:true, max:100 } } }
        });
    }
}

function updateIssuesTable() {
    const tbody = document.querySelector('#issuesTable tbody');
    tbody.innerHTML = '';
    const selectedProjectId = getProjectId(AppState.schedulerSelectedProject);
    let filtered = [...AppState.data.issues];
    if (AppState.schedulerSelectedProject!=='all' && selectedProjectId) {
        filtered = filtered.filter(issue=>issue.Project_ID===selectedProjectId);
    }
    if (filtered.length===0) {
        tbody.innerHTML = `<tr><td colspan="7" class="table-loading">Tidak ada data issue</td></tr>`;
        return;
    }
    filtered.forEach(issue=>{
        const row=document.createElement('tr');
        const priorityClass = (issue.Priority||'low').toLowerCase();
        const statusClass = getStatusClass(issue.Status||'open');
        row.innerHTML=`
          <td>${issue.Issue_ID||'N/A'}</td>
          <td>${getProjectName(issue.Project_ID)}</td>
          <td>${issue.Issue_Title||'N/A'}</td>
          <td><span class="status status--${priorityClass}">${issue.Priority||'Low'}</span></td>
          <td><span class="status status--${statusClass}">${issue.Status||'Open'}</span></td>
          <td>${issue.Assigned_To||'N/A'}</td>
          <td>${issue.Due_Date||'N/A'}</td>`;
        tbody.appendChild(row);
    });
}

function updatePlansTable() {
    const tbody = document.querySelector('#plansTable tbody');
    tbody.innerHTML='';
    const selectedProjectId = getProjectId(AppState.schedulerSelectedProject);
    let filtered=[...AppState.data.plans];
    if (AppState.schedulerSelectedProject!=='all' && selectedProjectId) {
        filtered = filtered.filter(plan=>plan.Project_ID===selectedProjectId);
    }
    if (filtered.length===0) {
        tbody.innerHTML=`<tr><td colspan="8" class="table-loading">Tidak ada data plan</td></tr>`;
        return;
    }
    filtered.forEach(plan=>{
        const row=document.createElement('tr');
        const priorityClass=(plan.Priority||'low').toLowerCase();
        const statusClass=getStatusClass(plan.Status||'planned');
        row.innerHTML=`
          <td>${plan.Plan_ID||'N/A'}</td>
          <td>${getProjectName(plan.Project_ID)}</td>
          <td>${plan.Plan_Title||'N/A'}</td>
          <td><span class="status status--${priorityClass}">${plan.Priority||'Low'}</span></td>
          <td><span class="status status--${statusClass}">${plan.Status||'Planned'}</span></td>
          <td>${plan.Assigned_To||'N/A'}</td>
          <td>${plan.Start_Date||'N/A'}</td>
          <td>${plan.End_Date||'N/A'}</td>`;
        tbody.appendChild(row);
    });
}

// =========================
// Safety & Permits
// =========================
function initializePermits() {
    initializeSafetyPyramidChart();
    renderActivePermits();
    updateSafetyKPIs();
    updateSafetyTable();
}

function updatePermitsContent() {
    initializeSafetyPyramidChart();
    renderActivePermits();
    updateSafetyKPIs();
    updateSafetyTable();
}

function initializeSafetyPyramidChart() {
    const ctx=document.getElementById('safetyPyramidChart').getContext('2d');
    if (AppState.charts.safetyPyramid) AppState.charts.safetyPyramid.destroy();
    const totals = AppState.data.safety.reduce((a,i)=>({
        nearMiss:a.nearMiss+(i.Near_Miss_Events||0),
        firstAid:a.firstAid+(i.First_Aid_Cases||0),
        medical:a.medical+(i.Medical_Treatment_Cases||0),
        lost:a.lost+(i.Lost_Time_Injuries||0),
        fatal:a.fatal+(i.Fatal_Accidents||0),
    }),{nearMiss:0,firstAid:0,medical:0,lost:0,fatal:0});
    AppState.charts.safetyPyramid=new Chart(ctx,{
        type:'bar',
        data:{ labels:['Near Miss','First Aid','Medical','Lost Time','Fatal'],
               datasets:[{label:'Safety Incidents',data:[totals.nearMiss,totals.firstAid,totals.medical,totals.lost,totals.fatal],
               backgroundColor:['#1FB8CD','#FFC185','#B4413C','#ECEBD5','#5D878F']}]},
        options:{ responsive:true, maintainAspectRatio:false, indexAxis:'y' }
    });
}

function renderActivePermits() {
    const c=document.getElementById('activePermitsList');
    c.innerHTML='';
    const active=AppState.data.permits.filter(p=>p.Status==='Open');
    if (active.length===0){ c.innerHTML='<div class="no-data-message">Tidak ada permit aktif</div>'; return;}
    active.forEach(p=>{
        const item=document.createElement('div');
        item.className='permit-item';
        item.innerHTML=`
          <div class="permit-header">
            <span class="permit-id">${p.PTW_ID}</span>
            <span class="permit-status ${p.Status.toLowerCase()}">${p.Status}</span>
          </div>
          <div class="permit-details">
            <p><strong>Type:</strong> ${p.Permit_Type}</p>
            <p><strong>Area:</strong> ${p.Area}</p>
            <p><strong>Contractor:</strong> ${p.Contractor}</p>
            <p><strong>Issued:</strong> ${p.Date_Issued}</p>
            <p><strong>Activity:</strong> ${p.Activity||'N/A'}</p>
          </div>`;
        c.appendChild(item);
    });
}

function updateSafetyKPIs() {
    const totalManpower=AppState.data.safety.reduce((s,i)=>s+(i.Total_Manpower||0),0);
    const safeManHours=AppState.data.safety.reduce((s,i)=>s+(i.Safe_Man_Hours||0),0);
    const totalManHours=AppState.data.safety.reduce((s,i)=>s+(i.Total_Manpower||0)*8,0);
    const incidentRate = totalManHours>0 ? (((totalManHours-safeManHours)/totalManHours)*100).toFixed(2):0;
    document.getElementById('totalManpower').textContent=totalManpower.toLocaleString();
    document.getElementById('safeManHours').textContent=safeManHours.toLocaleString();
    document.getElementById('incidentRate').textContent=incidentRate+'%';
}

function updateSafetyTable() {
    const tbody=document.querySelector('#safetyTable tbody');
    tbody.innerHTML='';
    const selectedProjectId=getProjectId(AppState.safetySelectedProject);
    let filtered=[...AppState.data.safety];
    if (AppState.safetySelectedProject!=='all'&&selectedProjectId){
        filtered=filtered.filter(s=>s.Project_ID===selectedProjectId);
    }
    if (filtered.length===0){ tbody.innerHTML='<tr><td colspan="7" class="table-loading">Tidak ada data safety</td></tr>'; return;}
    filtered.forEach(s=>{
        const row=document.createElement('tr');
        const totalIncidents=(s.Fatal_Accidents||0)+(s.Lost_Time_Injuries||0)+(s.Medical_Treatment_Cases||0)+(s.First_Aid_Cases||0);
        row.innerHTML=`
          <td>${s.Date||'N/A'}</td>
          <td>${getProjectName(s.Project_ID)}</td>
          <td><span class="status status--${s.Safety_Score>=95?'success':s.Safety_Score>=90?'warning':'error'}">${s.Safety_Score||0}%</span></td>
          <td>${(s.Total_Manpower||0).toLocaleString()}</td>
          <td>${(s.Safe_Man_Hours||0).toLocaleString()}</td>
          <td>${totalIncidents}</td>
          <td><span class="status status--${totalIncidents===0?'success':totalIncidents<=2?'warning':'error'}">${totalIncidents===0?'Excellent':totalIncidents<=2?'Good':'Needs Attention'}</span></td>`;
        tbody.appendChild(row);
    });
}

// =========================
// Reports / PPTX
// =========================
// (kode generatePowerPointReport sama seperti versi panjang yg sudah kamu punya,
// bisa dipaste langsung di sini tanpa diubah)

// =========================
// Project Filters
// =========================
function setupProjectFilters() {
    const mainFilter=document.getElementById('projectFilter');
    const schedulerFilter=document.getElementById('schedulerProjectFilter');
    const safetyFilter=document.getElementById('safetyProjectFilter');
    const reportFilter=document.getElementById('reportProject');

    [mainFilter,schedulerFilter,safetyFilter,reportFilter].forEach(f=>{
        if (f){
            f.innerHTML='<option value="all">Semua Proyek</option>';
            const names=[...new Set((AppState.data.projects||[]).map(p=>p.Project_Name))].sort();
            names.forEach(n=>{ const o=document.createElement('option'); o.value=n; o.textContent=n; f.appendChild(o); });
        }
    });
    if (mainFilter) mainFilter.addEventListener('change',e=>{ AppState.selectedProject=e.target.value; updateOverviewContent(); });
    if (schedulerFilter) schedulerFilter.addEventListener('change',e=>{ AppState.schedulerSelectedProject=e.target.value; updateSchedulerContent(); });
    if (safetyFilter) safetyFilter.addEventListener('change',e=>{ AppState.safetySelectedProject=e.target.value; updatePermitsContent(); });
}

// =========================
// Chatbot
// =========================
function initializeChatbot() {
    document.getElementById('chatbot').classList.add('collapsed');
}
function toggleChatbot(){ document.getElementById('chatbot').classList.toggle('collapsed'); }
function askQuestion(q){
    const mc=document.getElementById('chatMessages');
    const u=document.createElement('div'); u.className='message user-message'; u.textContent=q; mc.appendChild(u);
    setTimeout(()=>{
        const b=document.createElement('div'); b.className='message bot-message'; b.textContent=getBotResponse(q); mc.appendChild(b);
        mc.scrollTop=mc.scrollHeight;
    },1000);
}
function handleChatInput(e){ if (e.key==='Enter'){ sendChatMessage(); } }
function sendChatMessage(){ const i=document.getElementById('chatInput'); const m=i.value.trim(); if(m){ askQuestion(m); i.value=''; } }
function getBotResponse(q){
    const res={
        'Cara upload Excel?':'Upload file Excel dengan 7 sheet (Projects, S_Curve, Safety, Issues, Plans, Documents, Permits).',
        'Status safety terbaru?':`Safety rata-rata: ${AppState.data.projects.length>0?Math.round(AppState.data.projects.reduce((s,p)=>s+p.Safety_Score,0)/AppState.data.projects.length):0}%.`,
        'Generate report PowerPoint?':'Masuk ke "Generator Laporan", pilih opsi lalu klik Generate.'
    };
    return res[q]||'Dashboard Pertamina siap dipakai. Upload Excel, cek scheduler, dokumen, permits, dan generate report.';
}

// =========================
// Notifications
// =========================
function showNotification(msg,type='info'){
    const c=document.getElementById('notifications');
    const n=document.createElement('div');
    n.className=`notification ${type}`; n.textContent=msg; c.appendChild(n);
    setTimeout(()=>{ n.remove(); },5000);
}

document.addEventListener('click',e=>{
    const modal=document.getElementById('documentModal');
    if(e.target===modal) closeDocumentModal();
});
document.addEventListener('keydown',e=>{ if(e.key==='Escape') closeDocumentModal(); });
