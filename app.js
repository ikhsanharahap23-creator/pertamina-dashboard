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
const initialData = { /* ... tetap sama dengan data awal kamu ... */ };

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

// ... [semua fungsi lama tetap sama: overview, excel management, scheduler, permits, documents, dll] ...

// =========================
// Reports / PPTX
// =========================
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
    generatePowerPointReport(type, project, startDate, endDate);

    const reportRecord = {
        fileName: `Pertamina_${type}_Report_${startDate}_${endDate}.pptx`,
        type: type,
        project: project === 'all' ? 'Semua Proyek' : project,
        period: `${startDate} to ${endDate}`,
        createdDate: new Date().toLocaleString('id-ID'),
        status: 'Generated'
    };
    AppState.data.generatedReports.unshift(reportRecord);
    updateReportsTable();
}

function generatePowerPointReport(type, project, startDate, endDate) {
    try {
        if (typeof PptxGenJS === 'undefined') {
            showNotification('PowerPoint library not loaded. Generating fallback report...', 'warning');
            generateFallbackReport(type, project, startDate, endDate);
            return;
        }

        const pptx = new PptxGenJS();
        pptx.layout = '16x9';
        pptx.author = 'Pertamina Construction Dashboard';

        // === Cover Slide ===
        const cover = pptx.addSlide();
        cover.addImage({ path: './assets/cover_weekly.png', x:0, y:0, w:'100%', h:'100%' });
        cover.addImage({ path: './assets/logo_pertamina.png', x:0.5, y:0.5, w:2, h:0.8 });
        cover.addImage({ path: './assets/logo_danan.png', x:7, y:0.5, w:2, h:0.8 });
        cover.addText(`Construction ${type} Report`, { x:1, y:3, w:8, h:1, fontSize:32, bold:true, color:'FFFFFF', align:'center' });
        cover.addText(project==='all'?'Semua Proyek':project, { x:1, y:4, w:8, h:0.8, fontSize:24, color:'FFFFFF', align:'center' });
        cover.addText(`Period: ${startDate} to ${endDate}`, { x:1, y:5, w:8, h:0.6, fontSize:18, color:'FFFFFF', align:'center' });

        // === Executive Summary ===
        const summary = pptx.addSlide();
        summary.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:'100%', h:0.8, fill:'1F4E5C' });
        summary.addText('Executive Summary', { x:0, y:0, w:'100%', h:0.8, fontSize:24, bold:true, color:'FFFFFF', align:'center' });
        const totalProjects = AppState.data.projects.length;
        const avgProgress = totalProjects>0?Math.round(AppState.data.projects.reduce((s,p)=>s+p.Progress_Percent,0)/totalProjects):0;
        const avgSafety = totalProjects>0?Math.round(AppState.data.projects.reduce((s,p)=>s+p.Safety_Score,0)/totalProjects):0;
        const activePermits = AppState.data.permits.filter(p=>p.Status==='Open').length;
        summary.addText(`Total Projects: ${totalProjects}\nAverage Progress: ${avgProgress}%\nAverage Safety: ${avgSafety}%\nActive Permits: ${activePermits}`, { x:1, y:1.5, w:8, h:3, fontSize:18, color:'1F4E5C' });

        // === Project Details ===
        const proj = pptx.addSlide();
        proj.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:'100%', h:0.8, fill:'1F4E5C' });
        proj.addText('Project Details', { x:0, y:0, w:'100%', h:0.8, fontSize:24, bold:true, color:'FFFFFF', align:'center' });
        const tableData = [[
            { text:'Project Name', options:{ bold:true } },
            { text:'Progress %', options:{ bold:true } },
            { text:'Safety Score', options:{ bold:true } },
            { text:'Budget Used %', options:{ bold:true } }
        ]];
        AppState.data.projects.slice(0,6).forEach(p=>{
            tableData.push([ p.Project_Name, `${p.Progress_Percent}%`, `${p.Safety_Score}%`, `${p.Budget_Used_Percent}%` ]);
        });
        proj.addTable(tableData, { x:0.5, y:1.2, w:9, border:{ type:'solid', color:'1F4E5C', pt:1 } });

        // === Safety Analysis ===
        const safety = pptx.addSlide();
        safety.addShape(pptx.shapes.RECTANGLE, { x:0, y:0, w:'100%', h:0.8, fill:'1F4E5C' });
        safety.addText('Safety Analysis', { x:0, y:0, w:'100%', h:0.8, fontSize:24, bold:true, color:'FFFFFF', align:'center' });
        const totalManpower = AppState.data.safety.reduce((s,i)=>s+(i.Total_Manpower||0),0);
        const safeManHours = AppState.data.safety.reduce((s,i)=>s+(i.Safe_Man_Hours||0),0);
        safety.addText(`Total Manpower: ${totalManpower}\nSafe Man Hours: ${safeManHours}\nTotal Issues: ${AppState.data.issues.length}\nTotal Plans: ${AppState.data.plans.length}`, { x:1, y:1.5, w:8, h:3, fontSize:18, color:'1F4E5C' });

        const fileName = `Pertamina_${type}_Report_${new Date().toISOString().split('T')[0]}.pptx`;
        pptx.writeFile({ fileName });
        showNotification('PowerPoint report berhasil didownload!', 'success');
    } catch(err){
        console.error('Error generating PPT:', err);
        generateFallbackReport(type, project, startDate, endDate);
    }
}

// =========================
// Notifications & chatbot (tetap sama)
// =========================

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
