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
    // seed from static mapping for backward-compat
    _idToName = { ...PROJECT_ID_MAPPING };
    _nameToId = Object.fromEntries(Object.entries(_idToName).map(([id, name]) => [name, id]));

    // merge from uploaded/initial data
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
// Initial sample data (unchanged)
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
    // Load initial data
    AppState.data = { ...initialData };
    buildProjectMaps();

    // Setup navigation
    setupNavigation();

    // Initialize all sections
    initializeOverview();
    initializeExcelManagement();
    initializeScheduler();
    initializePermits();
    initializeDocuments();
    initializeReports();

    // Initialize chatbot
    initializeChatbot();

    // Setup project filters
    setupProjectFilters();

    showNotification('Dashboard Pertamina berhasil dimuat! Upload file Excel untuk data terlengkap.', 'success');
}

// Navigation setup
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
        if (sectionName === 'overview') {
            updateOverviewContent();
        } else if (sectionName === 'scheduler') {
            updateSchedulerContent();
        } else if (sectionName === 'documents') {
            renderDocuments();
        } else if (sectionName === 'permits') {
            updatePermitsContent();
        }
    }
}

// Overview Section
function initializeOverview() {
    updateKPICards();
    initializeProgressChart();
    initializeSafetyTrendChart();
}

function updateOverviewContent() {
    updateKPICards();
    initializeProgressChart();
    // Safety trend always shows all projects
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
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: { y: { beginAtZero: true, max: 100 } }
        }
    });
}

// Safety Trend Chart shows ALL projects
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

// Excel Management
function initializeExcelManagement() {
    setupFileUpload();
    updateUploadHistory();
}

function setupFileUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', function(e) { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', function(e) { e.preventDefault(); dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', function(e) {
        e.preventDefault(); dropZone.classList.remove('dragover');
        const files = e.dataTransfer.files; if (files.length > 0) handleFileUpload(files[0]);
    });

    fileInput.addEventListener('change', function(e) { if (e.target.files.length > 0) handleFileUpload(e.target.files[0]); });
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

    // fake progress
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 15; if (progress > 90) progress = 90;
        progressBar.style.width = progress + '%';
        progressText.textContent = `Processing... ${Math.round(progress)}%`;
        if (progress >= 90) clearInterval(interval);
    }, 100);

    reader.readAsArrayBuffer(file);
}

// Two-stage Excel processing to ensure maps
function processExcelWorkbook(workbook, file) {
    let totalRecords = 0;
    let processedSheets = [];

    // 1) read all to memory
    const raw = {};
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        if (jsonData.length > 0) raw[sheetName.toLowerCase()] = jsonData;
    });

    // 2) commit Projects first (if any) and rebuild maps
    const projectsKey = Object.keys(raw).find(k => k.includes('project'));
    if (projectsKey) {
        AppState.data.projects = raw[projectsKey].map(row => ({
            ...row,
            Project_ID: row.Project_ID || getProjectId(row.Project_Name) || row.Project_ID || ''
        }));
        totalRecords += raw[projectsKey].length;
        processedSheets.push('Projects');
        buildProjectMaps();
    }

    // 3) commit others with resolver
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

        // Cleanup unknowns (optional but tidy)
        ['sCurveData','safety','issues','plans','documents','permits'].forEach(k => {
            AppState.data[k] = (AppState.data[k] || []).filter(r => r.Project_ID && r.Project_ID !== 'PROJ_UNKNOWN');
        });

        // Rebuild maps and refresh UI
        buildProjectMaps();
        setupProjectFilters();
        setupDocumentFilters(); // ensure Documents project filter refreshed

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
            <td><span class="status status--success">${record.status}</span></td>
        `;
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

// Scheduler & S-Curve
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
        const labels = ['W1', 'W2', 'W3', 'W4', 'Jan', 'Feb', 'Mar', 'Apr'];
        const plannedData = [0, 15, 30, 45, 60, 75, 90, 100];
        const actualData = [0, 10, 25, 40, 55, 70, 85, 95];
        AppState.charts.sCurve = new Chart(ctx, {
            type: 'line',
            data: {
                labels,
                datasets: [
                    { label: 'Planned (%)', data: plannedData, borderColor: '#1FB8CD', backgroundColor: 'rgba(31, 184, 205, 0.1)', fill: false, tension: 0.4 },
                    { label: 'Actual (%)',  data: actualData,  borderColor: '#B4413C', backgroundColor: 'rgba(180, 65, 60, 0.1)',  fill: false, tension: 0.4 }
                ]
            },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } }, plugins: { title: { display: true, text: `S-Curve: ${AppState.schedulerSelectedProject === 'all' ? 'All Projects' : AppState.schedulerSelectedProject}` } } }
        });
    } else {
        const sortedData = sCurveData.sort((a, b) => new Date(a.Date_Reference) - new Date(b.Date_Reference));
        const labels = sortedData.map(item => item.Week_Month_Label || new Date(item.Date_Reference).toLocaleDateString('id-ID'));
        const plannedData = sortedData.map(item => Number(item.Cumulative_Planned) || 0);
        const actualData = sortedData.map(item => Number(item.Cumulative_Actual) || 0);
        AppState.charts.sCurve = new Chart(ctx, {
            type: 'line',
            data: { labels, datasets: [
                { label: 'Planned (%)', data: plannedData, borderColor: '#1FB8CD', backgroundColor: 'rgba(31, 184, 205, 0.1)', fill: false, tension: 0.4 },
                { label: 'Actual (%)',  data: actualData,  borderColor: '#B4413C', backgroundColor: 'rgba(180, 65, 60, 0.1)',  fill: false, tension: 0.4 }
            ]},
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } }, plugins: { title: { display: true, text: `S-Curve: ${AppState.schedulerSelectedProject === 'all' ? 'All Projects' : AppState.schedulerSelectedProject} (Excel Data)` } } }
        });
    }
}

function updateIssuesTable() {
    const tbody = document.querySelector('#issuesTable tbody');
    tbody.innerHTML = '';
    const selectedProjectId = getProjectId(AppState.schedulerSelectedProject);
    let filteredIssues = [...AppState.data.issues];
    if (AppState.schedulerSelectedProject !== 'all' && selectedProjectId) {
        filteredIssues = filteredIssues.filter(issue => issue.Project_ID === selectedProjectId);
    }
    if (filteredIssues.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="table-loading">${AppState.schedulerSelectedProject === 'all' ? 'Tidak ada data issue tersedia' : `Tidak ada issue untuk proyek ${AppState.schedulerSelectedProject}`}</td></tr>`;
        return;
    }
    filteredIssues.forEach(issue => {
        const row = document.createElement('tr');
        const priorityClass = (issue.Priority || 'low').toLowerCase();
        const statusClass = getStatusClass(issue.Status || 'open');
        row.innerHTML = `
            <td>${issue.Issue_ID || 'N/A'}</td>
            <td>${getProjectName(issue.Project_ID)}</td>
            <td>${issue.Issue_Title || 'N/A'}</td>
            <td><span class="status status--${priorityClass}">${issue.Priority || 'Low'}</span></td>
            <td><span class="status status--${statusClass}">${issue.Status || 'Open'}</span></td>
            <td>${issue.Assigned_To || 'N/A'}</td>
            <td>${issue.Due_Date || 'N/A'}</td>`;
        tbody.appendChild(row);
    });
}

function updatePlansTable() {
    const tbody = document.querySelector('#plansTable tbody');
    tbody.innerHTML = '';
    const selectedProjectId = getProjectId(AppState.schedulerSelectedProject);
    let filteredPlans = [...AppState.data.plans];
    if (AppState.schedulerSelectedProject !== 'all' && selectedProjectId) {
        filteredPlans = filteredPlans.filter(plan => plan.Project_ID === selectedProjectId);
    }
    if (filteredPlans.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" class="table-loading">${AppState.schedulerSelectedProject === 'all' ? 'Tidak ada data plan tersedia' : `Tidak ada plan untuk proyek ${AppState.schedulerSelectedProject}`}</td></tr>`;
        return;
    }
    filteredPlans.forEach(plan => {
        const row = document.createElement('tr');
        const priorityClass = (plan.Priority || 'low').toLowerCase();
        const statusClass = getStatusClass(plan.Status || 'planned');
        row.innerHTML = `
            <td>${plan.Plan_ID || 'N/A'}</td>
            <td>${getProjectName(plan.Project_ID)}</td>
            <td>${plan.Plan_Title || 'N/A'}</td>
            <td><span class="status status--${priorityClass}">${plan.Priority || 'Low'}</span></td>
            <td><span class="status status--${statusClass}">${plan.Status || 'Planned'}</span></td>
            <td>${plan.Assigned_To || 'N/A'}</td>
            <td>${plan.Start_Date || 'N/A'}</td>
            <td>${plan.End_Date || 'N/A'}</td>`;
        tbody.appendChild(row);
    });
}

// Safety & Permits
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
    const ctx = document.getElementById('safetyPyramidChart').getContext('2d');
    if (AppState.charts.safetyPyramid) AppState.charts.safetyPyramid.destroy();
    const totals = AppState.data.safety.reduce((acc, item) => {
        acc.nearMiss += item.Near_Miss_Events || 0;
        acc.firstAid += item.First_Aid_Cases || 0;
        acc.medicalTreatment += item.Medical_Treatment_Cases || 0;
        acc.lostTime += item.Lost_Time_Injuries || 0;
        acc.fatal += item.Fatal_Accidents || 0;
        return acc;
    }, { nearMiss: 0, firstAid: 0, medicalTreatment: 0, lostTime: 0, fatal: 0 });

    AppState.charts.safetyPyramid = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Near Miss', 'First Aid', 'Medical Treatment', 'Lost Time', 'Fatal'],
            datasets: [{
                label: 'Safety Incidents',
                data: [totals.nearMiss, totals.firstAid, totals.medicalTreatment, totals.lostTime, totals.fatal],
                backgroundColor: ['#1FB8CD', '#FFC185', '#B4413C', '#ECEBD5', '#5D878F'],
                borderColor: ['#1FB8CD', '#FFC185', '#B4413C', '#ECEBD5', '#5D878F'],
                borderWidth: 1
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', scales: { x: { beginAtZero: true } } }
    });
}

function renderActivePermits() {
    const container = document.getElementById('activePermitsList');
    container.innerHTML = '';
    const activePermits = AppState.data.permits.filter(p => p.Status === 'Open');
    if (activePermits.length === 0) { container.innerHTML = '<div class="no-data-message">Tidak ada permit aktif</div>'; return; }
    activePermits.forEach(permit => {
        const item = document.createElement('div');
        item.className = 'permit-item';
        item.innerHTML = `
            <div class="permit-header">
                <span class="permit-id">${permit.PTW_ID}</span>
                <span class="permit-status ${permit.Status.toLowerCase()}">${permit.Status}</span>
            </div>
            <div class="permit-details">
                <p><strong>Type:</strong> ${permit.Permit_Type}</p>
                <p><strong>Area:</strong> ${permit.Area}</p>
                <p><strong>Contractor:</strong> ${permit.Contractor}</p>
                <p><strong>Issued:</strong> ${permit.Date_Issued}</p>
                <p><strong>Activity:</strong> ${permit.Activity || 'N/A'}</p>
            </div>`;
        container.appendChild(item);
    });
}

function updateSafetyKPIs() {
    const totalManpower = AppState.data.safety.reduce((sum, item) => sum + (item.Total_Manpower || 0), 0);
    const safeManHours = AppState.data.safety.reduce((sum, item) => sum + (item.Safe_Man_Hours || 0), 0);
    const totalManHours = AppState.data.safety.reduce((sum, item) => sum + (item.Total_Manpower || 0) * 8, 0);
    let incidentRateNum = 0;
    if (totalManHours > 0) {
        const unsafe = Math.max(0, totalManHours - safeManHours);
        incidentRateNum = (unsafe / totalManHours) * 100;
    }
    const incidentRate = incidentRateNum.toFixed(2);

    document.getElementById('totalManpower').textContent = totalManpower.toLocaleString();
    document.getElementById('safeManHours').textContent = safeManHours.toLocaleString();
    document.getElementById('incidentRate').textContent = incidentRate + '%';
}

function updateSafetyTable() {
    const tbody = document.querySelector('#safetyTable tbody');
    tbody.innerHTML = '';
    const selectedProjectId = getProjectId(AppState.safetySelectedProject);
    let filteredSafety = [...AppState.data.safety];
    if (AppState.safetySelectedProject !== 'all' && selectedProjectId) {
        filteredSafety = filteredSafety.filter(safety => safety.Project_ID === selectedProjectId);
    }
    if (filteredSafety.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="table-loading">Tidak ada data safety tersedia</td></tr>';
        return;
    }
    filteredSafety.forEach(safety => {
        const row = document.createElement('tr');
        const totalIncidents = (safety.Fatal_Accidents || 0) + (safety.Lost_Time_Injuries || 0) + (safety.Medical_Treatment_Cases || 0) + (safety.First_Aid_Cases || 0);
        row.innerHTML = `
            <td>${safety.Date || 'N/A'}</td>
            <td>${getProjectName(safety.Project_ID)}</td>
            <td><span class="status status--${safety.Safety_Score >= 95 ? 'success' : safety.Safety_Score >= 90 ? 'warning' : 'error'}">${safety.Safety_Score || 0}%</span></td>
            <td>${(safety.Total_Manpower || 0).toLocaleString()}</td>
            <td>${(safety.Safe_Man_Hours || 0).toLocaleString()}</td>
            <td>${totalIncidents}</td>
            <td><span class="status status--${totalIncidents === 0 ? 'success' : totalIncidents <= 2 ? 'warning' : 'error'}">${totalIncidents === 0 ? 'Excellent' : totalIncidents <= 2 ? 'Good' : 'Needs Attention'}</span></td>`;
        tbody.appendChild(row);
    });
}

// Documents
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
        // rebuild each time this is called
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
        grid.innerHTML = `
            <div class="no-data-message">
                <p>Tidak ada dokumen ditemukan. ${searchQuery ? 'Coba ubah kriteria pencarian.' : 'Upload file Excel dengan sheet Documents untuk menambahkan data dokumen.'}</p>
            </div>`;
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

// =======================================================
// Reports / PPTX  —>> UPDATED WITH YOUR DESIGN
// =======================================================

// --- PPTX helpers & theme (NEW) ---
const ASSETS_BASE = './assets'; // ubah jika folder berbeda
const PPT_THEME = {
  bg: 'FFFFFF',
  brandDark: '1F4E5C',
  brandTeal: '21808D',
  accent: 'FFC185',
  muted: '6B7C85',
  tableHeaderBg: 'EDF2F5',
  tileBg: 'F5F7F9'
};

// Convert image url to dataURL (works in same-origin / CORS enabled)
async function imgToDataURL(url) {
  const res = await fetch(url, { cache: 'no-store' });
  const blob = await res.blob();
  return await new Promise((resolve) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.readAsDataURL(blob);
  });
}

function addDividerSlide(pptx, title, logos) {
  const s = pptx.addSlide({ bkgd: 'BFBFBF' });
  if (logos?.pertamina) s.addImage({ data: logos.pertamina, x: 8.8, y: 0.2, w: 1, h: 1 });
  if (logos?.danan)     s.addImage({ data: logos.danan,     x: 0.2, y: 0.2, w: 1.6, h: 1 });
  s.addShape(pptx.ShapeType.rect, { x: 2, y: 2.2, w: 6, h: 1.2, fill: 'FFFFFF' });
  s.addText(title, { x: 2, y: 2.2, w: 6, h: 1.2, fontSize: 28, align: 'center', bold: true, color: PPT_THEME.brandDark });
  return s;
}

function addZebraTable(slide, rows) {
  slide.addTable(rows, {
    x: 0.5, y: 1.5, w: 9, h: 4.6,
    border: { type: 'solid', color: 'D5DCE3', pt: 1 },
    fill: 'FFFFFF',
    header: true,
    fillHdr: PPT_THEME.tableHeaderBg,
    fontSize: 12,
    alternateRowFill: 'FAFBFC'
  });
}

function initializeReports() {
    updateReportsTable();
    setupReportGeneration();
    setupDateInputs();
    updateReportPreview();
}

function setupDateInputs() {
    const today = new Date();
    const twoWeeksAgo = new Date(today.getTime() - (14 * 24 * 60 * 60 * 1000));
    document.getElementById('startDate').value = twoWeeksAgo.toISOString().split('T')[0];
    document.getElementById('endDate').value = today.toISOString().split('T')[0];
}

function setupReportGeneration() {
    const generateBtn = document.getElementById('generateReport');
    generateBtn.addEventListener('click', generateReport);
}

function updateReportPreview() {
    const totalProjects = AppState.data.projects.length;
    const activePermits = AppState.data.permits.filter(p => p.Status === 'Open').length;
    const totalIssues = AppState.data.issues.length;
    const totalPlans = AppState.data.plans.length;
    const avgSafety = totalProjects > 0 
        ? Math.round(AppState.data.projects.reduce((sum, p) => sum + p.Safety_Score, 0) / totalProjects)
        : 0;
    document.getElementById('previewProjects').textContent = totalProjects;
    document.getElementById('previewPermits').textContent = activePermits;
    document.getElementById('previewIssues').textContent = totalIssues;
    document.getElementById('previewPlans').textContent = totalPlans;
    document.getElementById('previewSafety').textContent = avgSafety + '%';
}

function generateReport() {
    const type = document.getElementById('reportType').value;
    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;
    const project = document.getElementById('reportProject').value;

    showNotification('Generating PowerPoint report...', 'info');
    // UPDATED: now async but we don't await to keep UX responsive
    generatePowerPointReport(type, project, startDate, endDate);

    const reportRecord = {
        fileName: `Pertamina_${type}_Report_${startDate}_${endDate}.pptx`,
        type: type,
        project: project === 'all' ? 'Semua Proyek' : project,
        period: `${startDate} to ${endDate}`,
        createdDate: new Date().toLocaleString('id-ID'),
        status: 'Generated'
    };
    AppState.data.generatedReports = AppState.data.generatedReports || [];
    AppState.data.generatedReports.unshift(reportRecord);
    updateReportsTable();
}

// ==================== REPLACED WITH YOUR DESIGN ====================
async function generatePowerPointReport(type, project, startDate, endDate) {
  try {
    if (typeof PptxGenJS === 'undefined') {
      showNotification('PowerPoint library not loaded. Generating fallback report...', 'warning');
      generateFallbackReport(type, project, startDate, endDate);
      return;
    }

    // Preload assets
    const coverPath = `${ASSETS_BASE}/cover_weekly.png`;
    const logoPertaminaPath = `${ASSETS_BASE}/logo_pertamina.png`;
    const logoDananPath = `${ASSETS_BASE}/logo_danan.png`;

    const [coverImg, logoPertamina, logoDanan] = await Promise.all([
      imgToDataURL(coverPath).catch(()=>null),
      imgToDataURL(logoPertaminaPath).catch(()=>null),
      imgToDataURL(logoDananPath).catch(()=>null),
    ]);

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'Pertamina Construction Dashboard';
    pptx.company = 'Pertamina';
    pptx.title = `Pertamina Construction ${type?.charAt(0).toUpperCase() + type?.slice(1)} Report`;

    const filteredProjects = project === 'all'
      ? AppState.data.projects
      : AppState.data.projects.filter(p => p.Project_Name === project);

    // KPIs
    const totalProjects = filteredProjects.length || AppState.data.projects.length;
    const baseProjects = filteredProjects.length ? filteredProjects : AppState.data.projects;
    const avgProgress = totalProjects ? Math.round(baseProjects.reduce((s,p)=>s+(p.Progress_Percent||0),0)/totalProjects) : 0;
    const avgSafety   = totalProjects ? Math.round(baseProjects.reduce((s,p)=>s+(p.Safety_Score||0),0)/totalProjects) : 0;
    const activePermits = AppState.data.permits.filter(p=>p.Status==='Open').length;
    const totalManpower = AppState.data.safety.reduce((sum,i)=>sum+(i.Total_Manpower||0),0);
    const safeManHours  = AppState.data.safety.reduce((sum,i)=>sum+(i.Safe_Man_Hours||0),0);

    // ---------- Slide 1: COVER ----------
    {
      const s = pptx.addSlide();
      if (coverImg) s.addImage({ data: coverImg, x: 0, y: 0, w: 10, h: 5.63 });
      if (logoPertamina) s.addImage({ data: logoPertamina, x: 8.8, y: 0.2, w: 1.1, h: 1.1 });
      if (logoDanan)     s.addImage({ data: logoDanan,     x: 0.2, y: 0.2, w: 1.7, h: 1.1 });

      s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 0.6, w: 8.0, h: 2.3, fill: { color: '2E6F86', transparency: 35 } });
      s.addText('Construction Weekly Report', { x: 1.3, y: 0.8, w: 7.4, h: 0.8, fontSize: 28, bold: true, color: 'FFFFFF' });
      s.addText(project === 'all' ? 'Semua Proyek' : project, { x: 1.3, y: 1.5, w: 7.4, h: 0.6, fontSize: 24, bold: true, color: 'FFFFFF' });
      s.addText(`Period: ${startDate} to ${endDate}`, { x: 1.3, y: 2.1, w: 7.4, h: 0.6, fontSize: 22, bold: true, color: 'FFFFFF' });
    }

    // ---------- Slide 2: EXECUTIVE SUMMARY (Divider) ----------
    const logos = { pertamina: logoPertamina, danan: logoDanan };
    addDividerSlide(pptx, 'Executive Summary', logos);

    // ---------- Slide 3: EXECUTIVE SUMMARY CONTENT ----------
    {
      const s = pptx.addSlide();
      if (logoPertamina) s.addImage({ data: logoPertamina, x: 8.8, y: 0.2, w: 1.1, h: 1.1 });
      if (logoDanan)     s.addImage({ data: logoDanan,     x: 0.2, y: 0.2, w: 1.7, h: 1.1 });

      const lines = [
        `Total Projects: ${totalProjects}`,
        `Average Progress: ${avgProgress}%`,
        `Average Safety Score: ${avgSafety}%`,
        `Active Permits: ${activePermits}`,
        `Generated: ${new Date().toLocaleString('id-ID')}`
      ];
      s.addText(lines.join('\n\n'), { x: 1.0, y: 2.0, w: 8.0, h: 3.2, fontSize: 22, color: PPT_THEME.brandDark, bold: false });
    }

    // ---------- Slide 4: PROGRESS DETAILS (Divider) ----------
    addDividerSlide(pptx, 'Progress Project Details', logos);

    // ---------- Slide 5: PROJECT TABLE ----------
    {
      const s = pptx.addSlide();
      if (logoPertamina) s.addImage({ data: logoPertamina, x: 8.8, y: 0.2, w: 1.1, h: 1.1 });
      if (logoDanan)     s.addImage({ data: logoDanan,     x: 0.2, y: 0.2, w: 1.7, h: 1.1 });

      const rows = [[
        { text: 'Project Name', options: { bold: true, color: PPT_THEME.brandDark } },
        { text: 'Progress %',  options: { bold: true, color: PPT_THEME.brandDark } },
        { text: 'Safety Score',options: { bold: true, color: PPT_THEME.brandDark } },
        { text: 'Budget Used %', options: { bold: true, color: PPT_THEME.brandDark } }
      ]];

      baseProjects.forEach(p=>{
        rows.push([
          p.Project_Name?.replace(' Substation','') || '—',
          `${Number(p.Progress_Percent || 0).toFixed(1)}%`,
          `${Number(p.Safety_Score || 0).toFixed(1)}%`,
          `${Number(p.Budget_Used_Percent || 0).toFixed(1)}%`
        ]);
      });

      addZebraTable(s, rows);
    }

    // ---------- Slide 6: SAFETY (Divider) ----------
    addDividerSlide(pptx, 'Safety Analysis', logos);

    // ---------- Slide 7: SAFETY CONTENT ----------
    {
      const s = pptx.addSlide();
      if (logoPertamina) s.addImage({ data: logoPertamina, x: 8.8, y: 0.2, w: 1.1, h: 1.1 });
      if (logoDanan)     s.addImage({ data: logoDanan,     x: 0.2, y: 0.2, w: 1.7, h: 1.1 });

      const lines = [
        `Total Manpower: ${totalManpower.toLocaleString()}`,
        `Safe Man Hours: ${safeManHours.toLocaleString()}`,
        `Total Issues: ${AppState.data.issues.length}`,
        `Total Plans: ${AppState.data.plans.length}`
      ];
      s.addText(lines.join('\n\n'), { x: 1.0, y: 2.0, w: 8.0, h: 3.2, fontSize: 22, color: PPT_THEME.brandDark });
    }

    const fileName = `Pertamina_${type}_Report_${new Date().toISOString().split('T')[0]}.pptx`;
    await pptx.writeFile({ fileName });
    showNotification('PowerPoint report berhasil didownload!', 'success');

  } catch (error) {
    console.error('Error generating PowerPoint:', error);
    showNotification('Error generating PowerPoint: ' + error.message, 'error');
    generateFallbackReport(type, project, startDate, endDate);
  }
}
// ==================== END REPLACED ====================

function generateFallbackReport(type, project, startDate, endDate) {
    const fileName = `Pertamina_${type}_Report_${new Date().toISOString().split('T')[0]}.txt`;
    const reportContent = `
PERTAMINA - Construction ${type.charAt(0).toUpperCase() + type.slice(1)} Report
======================================================

Project: ${project === 'all' ? 'Semua Proyek' : project}
Period: ${startDate} to ${endDate}
Generated: ${new Date().toLocaleString('id-ID')}

EXECUTIVE SUMMARY
-----------------
Total Projects: ${AppState.data.projects.length}
Active Permits: ${AppState.data.permits.filter(p => p.Status === 'Open').length}
Total Issues: ${AppState.data.issues.length}
Total Plans: ${AppState.data.plans.length}

SAFETY SUMMARY
--------------
Total Manpower: ${AppState.data.safety.reduce((sum, item) => sum + (item.Total_Manpower || 0), 0)}
Safe Man Hours: ${AppState.data.safety.reduce((sum, item) => sum + (item.Safe_Man_Hours || 0), 0)}

This is a fallback text report. PowerPoint functionality will be available when all libraries are loaded.
    `;
    const blob = new Blob([reportContent], { type: 'text/plain' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.style.display = 'none'; a.href = url; a.download = fileName; document.body.appendChild(a); a.click();
    window.URL.revokeObjectURL(url); document.body.removeChild(a);
    showNotification('Fallback report berhasil didownload!', 'success');
}

function updateReportsTable() {
    const tbody = document.querySelector('#reportsTable tbody');
    tbody.innerHTML = '';
    (AppState.data.generatedReports || []).forEach(report => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${report.fileName}</td>
            <td>${report.type}</td>
            <td>${report.project}</td>
            <td>${report.period}</td>
            <td>${report.createdDate}</td>
            <td><span class="status status--success">${report.status}</span></td>`;
        tbody.appendChild(row);
    });
}

// Project Filters Setup
function setupProjectFilters() {
    const mainFilter = document.getElementById('projectFilter');
    const schedulerFilter = document.getElementById('schedulerProjectFilter');
    const safetyFilter = document.getElementById('safetyProjectFilter');
    const reportFilter = document.getElementById('reportProject');

    [mainFilter, schedulerFilter, safetyFilter, reportFilter].forEach(filter => {
        if (filter) {
            const firstOption = filter.querySelector('option[value="all"]');
            filter.innerHTML = '';
            if (firstOption) {
                filter.appendChild(firstOption.cloneNode(true));
            } else {
                const allOption = document.createElement('option');
                allOption.value = 'all';
                allOption.textContent = 'Semua Proyek';
                filter.appendChild(allOption);
            }
            const names = [...new Set((AppState.data.projects || []).map(p => p.Project_Name))].sort((a,b)=>a.localeCompare(b));
            names.forEach(n => { const option = document.createElement('option'); option.value = n; option.textContent = n; filter.appendChild(option); });
        }
    });

    if (mainFilter) {
        mainFilter.addEventListener('change', function(e) { AppState.selectedProject = e.target.value; updateOverviewContent(); });
    }
    if (schedulerFilter) {
        schedulerFilter.addEventListener('change', function(e) { AppState.schedulerSelectedProject = e.target.value; updateSchedulerContent(); });
    }
    if (safetyFilter) {
        safetyFilter.addEventListener('change', function(e) { AppState.safetySelectedProject = e.target.value; updatePermitsContent(); });
    }
}

function getFilteredProjects() {
    if (AppState.selectedProject === 'all') return AppState.data.projects;
    return AppState.data.projects.filter(p => p.Project_Name === AppState.selectedProject);
}

// Chatbot (unchanged)
function initializeChatbot() {
    const chatbot = document.getElementById('chatbot');
    chatbot.classList.add('collapsed');
}
function toggleChatbot() { const chatbot = document.getElementById('chatbot'); chatbot.classList.toggle('collapsed'); }
function askQuestion(question) {
    const messagesContainer = document.getElementById('chatMessages');
    const userMessage = document.createElement('div'); userMessage.className = 'message user-message'; userMessage.textContent = question; messagesContainer.appendChild(userMessage);
    setTimeout(() => {
        const botMessage = document.createElement('div'); botMessage.className = 'message bot-message'; botMessage.textContent = getBotResponse(question); messagesContainer.appendChild(botMessage); messagesContainer.scrollTop = messagesContainer.scrollHeight; }, 1000);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}
function handleChatInput(event) { if (event.key === 'Enter') { sendChatMessage(); } }
function sendChatMessage() { const input = document.getElementById('chatInput'); const message = input.value.trim(); if (message) { askQuestion(message); input.value = ''; } }
function getBotResponse(question) {
    const responses = {
        'Cara upload Excel?': 'Upload file Excel dengan 7 sheet (Projects, S_Curve, Safety, Issues, Plans, Documents, Permits) di section "Upload Data Excel". Drag & drop atau klik "Pilih File Excel". Dashboard akan otomatis terupdate!',
        'Status safety terbaru?': `Safety score rata-rata adalah ${AppState.data.projects.length > 0 ? Math.round(AppState.data.projects.reduce((sum, p) => sum + p.Safety_Score, 0) / AppState.data.projects.length) : 0}%. Total manpower: ${AppState.data.safety.reduce((sum, item) => sum + (item.Total_Manpower || 0), 0)}. Safe man hours: ${AppState.data.safety.reduce((sum, item) => sum + (item.Safe_Man_Hours || 0), 0)}.`,
        'Generate report PowerPoint?': 'Klik section "Generator Laporan", pilih tipe laporan, tanggal, dan proyek, lalu klik "Generate PowerPoint Report". File .pptx akan otomatis terdownload dengan data lengkap!'
    };
    return responses[question] || 'Dashboard Pertamina sudah fully integrated: Upload Excel 7 sheet untuk update semua data, AI Scheduler dengan S-Curve Week/Month format, Smart Document Search terintegrasi, PowerPoint reports berfungsi sempurna, dan Safety Monitoring dengan Total Manpower & Safe Man Hours. Semua fitur working tanpa error!';
}

// Notifications & modal events
function showNotification(message, type = 'info') {
    const container = document.getElementById('notifications');
    const notification = document.createElement('div');
    notification.className = `notification ${type}`; notification.textContent = message; container.appendChild(notification);
    setTimeout(() => { notification.remove(); }, 5000);
}

document.addEventListener('click', function(event) {
    const modal = document.getElementById('documentModal');
    if (event.target === modal) { closeDocumentModal(); }
});

document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') { closeDocumentModal(); }
});

  // ---- v2 generator (ringkas): panggil fungsi yang sudah kamu punya ----
  async function generatePowerPointReport_v2(type, project, startDate, endDate) {
    if (typeof PptxGenJS === 'undefined') {
      showNotification('PowerPoint library not loaded. Fallback report dibuat.', 'warning');
      generateFallbackReport(type, project, startDate, endDate);
      return;
    }
    // panggil fungsi baru yang sudah ada di file (kalau namanya sama, pakai yang ini)
    if (typeof generatePowerPointReport === 'function' && generatePowerPointReport !== generatePowerPointReport_v2) {
      // Sudah ada versi baru di atas -> pakai itu
      return await generatePowerPointReport(type, project, startDate, endDate);
    }

    // ====== COPY RINGKASAN INTI DARI GENERATOR V2 (kalau di atas tidak ada) ======
    const ASSETS_BASE = './assets';
    const toDataURL = async (p) => {
      try { const r = await fetch(p, {cache:'no-store'}); const b = await r.blob();
        return await new Promise(res=>{ const fr=new FileReader(); fr.onload=()=>res(fr.result); fr.readAsDataURL(b); });
      } catch { return null; }
    };
    const [cover, lgP, lgD] = await Promise.all([
      toDataURL(`${ASSETS_BASE}/cover_weekly.png`),
      toDataURL(`${ASSETS_BASE}/logo_pertamina.png`),
      toDataURL(`${ASSETS_BASE}/logo_danan.png`)
    ]);

    const pptx = new PptxGenJS(); pptx.layout = 'LAYOUT_16x9';

    const projects = (project==='all' ? AppState.data.projects :
                      AppState.data.projects.filter(p=>p.Project_Name===project));
    const base = projects.length ? projects : AppState.data.projects;
    const total = base.length;
    const avgProg = total ? Math.round(base.reduce((s,p)=>s+(p.Progress_Percent||0),0)/total) : 0;
    const avgSaf  = total ? Math.round(base.reduce((s,p)=>s+(p.Safety_Score||0),0)/total) : 0;

    // Cover
    { const s = pptx.addSlide();
      if (cover) s.addImage({data:cover,x:0,y:0,w:10,h:5.63});
      if (lgP) s.addImage({data:lgP,x:8.8,y:0.2,w:1.1,h:1.1});
      if (lgD) s.addImage({data:lgD,x:0.2,y:0.2,w:1.7,h:1.1});
      s.addShape(pptx.ShapeType.rect,{x:1,y:0.6,w:8,h:2.3,fill:{color:'2E6F86',transparency:35}});
      s.addText('Construction Weekly Report',{x:1.3,y:0.8,w:7.4,h:0.8,fontSize:28,bold:true,color:'FFFFFF'});
      s.addText(project==='all'?'Semua Proyek':project,{x:1.3,y:1.5,w:7.4,h:0.6,fontSize:24,bold:true,color:'FFFFFF'});
      s.addText(`Period: ${startDate} to ${endDate}`,{x:1.3,y:2.1,w:7.4,h:0.6,fontSize:22,bold:true,color:'FFFFFF'});
    }
    // Divider
    const divider=(title)=>{const s=pptx.addSlide({bkgd:'BFBFBF'});
      if (lgP) s.addImage({data:lgP,x:8.8,y:0.2,w:1.1,h:1.1});
      if (lgD) s.addImage({data:lgD,x:0.2,y:0.2,w:1.7,h:1.1});
      s.addShape(pptx.ShapeType.rect,{x:2,y:2.2,w:6,h:1.2,fill:'FFFFFF'});
      s.addText(title,{x:2,y:2.2,w:6,h:1.2,fontSize:28,align:'center',bold:true,color:'1F4E5C'}); };

    divider('Executive Summary');
    { const s=pptx.addSlide();
      if (lgP) s.addImage({data:lgP,x:8.8,y:0.2,w:1.1,h:1.1});
      if (lgD) s.addImage({data:lgD,x:0.2,y:0.2,w:1.7,h:1.1});
      s.addText([
        `Total Projects: ${total||AppState.data.projects.length}`,
        `Average Progress: ${avgProg}%`,
        `Average Safety Score: ${avgSaf}%`,
        `Active Permits: ${AppState.data.permits.filter(p=>p.Status==='Open').length}`,
        `Generated: ${new Date().toLocaleString('id-ID')}`
      ].join('\n\n'),{x:1,y:2,w:8,h:3.2,fontSize:22,color:'1F4E5C'});
    }

    divider('Progress Project Details');
    { const s=pptx.addSlide();
      if (lgP) s.addImage({data:lgP,x:8.8,y:0.2,w:1.1,h:1.1});
      if (lgD) s.addImage({data:lgD,x:0.2,y:0.2,w:1.7,h:1.1});
      const rows=[[
        {text:'Project Name',options:{bold:true,color:'1F4E5C'}},
        {text:'Progress %',options:{bold:true,color:'1F4E5C'}},
        {text:'Safety Score',options:{bold:true,color:'1F4E5C'}},
        {text:'Budget Used %',options:{bold:true,color:'1F4E5C'}}
      ]];
      base.forEach(p=>rows.push([
        p.Project_Name?.replace(' Substation','')||'—',
        `${Number(p.Progress_Percent||0).toFixed(1)}%`,
        `${Number(p.Safety_Score||0).toFixed(1)}%`,
        `${Number(p.Budget_Used_Percent||0).toFixed(1)}%`
      ]));
      s.addTable(rows,{x:0.5,y:1.5,w:9,h:4.6,header:true,fillHdr:'EDF2F5',border:{type:'solid',color:'D5DCE3',pt:1},alternateRowFill:'FAFBFC',fontSize:12});
    }

    divider('Safety Analysis');
    { const s=pptx.addSlide();
      if (lgP) s.addImage({data:lgP,x:8.8,y:0.2,w:1.1,h:1.1});
      if (lgD) s.addImage({data:lgD,x:0.2,y:0.2,w:1.7,h:1.1});
      const tm=AppState.data.safety.reduce((a,i)=>a+(i.Total_Manpower||0),0);
      const sh=AppState.data.safety.reduce((a,i)=>a+(i.Safe_Man_Hours||0),0);
      s.addText([
        `Total Manpower: ${tm.toLocaleString()}`,
        `Safe Man Hours: ${sh.toLocaleString()}`,
        `Total Issues: ${AppState.data.issues.length}`,
        `Total Plans: ${AppState.data.plans.length}`
      ].join('\n\n'),{x:1,y:2,w:8,h:3.2,fontSize:22,color:'1F4E5C'});
    }

    await pptx.writeFile({ fileName: `Pertamina_${type}_Report_${new Date().toISOString().split('T')[0]}.pptx` });
    showNotification('PowerPoint (v2) berhasil didownload!', 'success');
  }

  // Paksa tombol "Generate" pakai v2
  const wire = () => {
    const btn = document.getElementById('generateReport');
    if (!btn) return;
    btn.onclick = async () => {
      const type = document.getElementById('reportType').value;
      const startDate = document.getElementById('startDate').value;
      const endDate = document.getElementById('endDate').value;
      const project = document.getElementById('reportProject').value;
      showNotification('Generating PowerPoint (v2)...', 'info');
      await generatePowerPointReport_v2(type, project, startDate, endDate);
    };
    console.log('[PPT] Button wired to v2');
  };
  if (document.readyState === 'complete' || document.readyState === 'interactive') wire();
  else document.addEventListener('DOMContentLoaded', wire);

  // Expose untuk debug
  window.__ppt_v2 = generatePowerPointReport_v2;
})();
