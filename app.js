// ========================================================
// Global application state
// ========================================================
const AppState = {
  currentSection: "overview",
  selectedProject: "all",
  schedulerSelectedProject: "all",
  safetySelectedProject: "all",
  data: {
    projects: [],
    sCurveData: [],
    safety: [],
    issues: [],
    plans: [],
    documents: [],
    permits: [],
    uploadHistory: [],
    generatedReports: [],
  },
  charts: {},
};

// Mapping proyek default
const PROJECT_ID_MAPPING = {
  PROJ001: "Petani Substation",
  PROJ002: "Menggala Substation",
  PROJ003: "Nella Substation",
  PROJ004: "Bangko Substation",
  PROJ005: "Balam SS",
  PROJ006: "Sintong SS",
  PROJ007: "OKB Substation",
};

// ========================================================
// Mapping helpers
// ========================================================
let _idToName = {};
let _nameToId = {};

function buildProjectMaps() {
  _idToName = { ...PROJECT_ID_MAPPING };
  _nameToId = Object.fromEntries(
    Object.entries(_idToName).map(([id, name]) => [name, id])
  );

  (AppState?.data?.projects || []).forEach((p) => {
    const id = String(p.Project_ID || "").trim();
    const nm = String(p.Project_Name || "").trim();
    if (id && nm) {
      _idToName[id] = nm;
      _nameToId[nm] = id;
    }
  });
}

function getProjectName(projectId) {
  return _idToName[projectId] || projectId || "Unknown Project";
}
function getProjectId(projectName) {
  if (!projectName || projectName === "all") return null;
  return _nameToId[projectName] || null;
}
function resolveProjectId(row) {
  return (
    row.Project_ID ||
    getProjectId(row.Project_Name) ||
    "PROJ_UNKNOWN"
  );
}

// Status CSS
function getStatusClass(status) {
  if (!status) return "info";
  const s = status.toLowerCase();
  if (s.includes("open") || s.includes("planned")) return "info";
  if (s.includes("progress")) return "warning";
  if (s.includes("completed") || s.includes("closed")) return "success";
  return "info";
}

// ========================================================
// Initial Sample Data
// ========================================================
const initialData = {
  projects: [
    {
      Date: "2024-01-15",
      Project_ID: "PROJ001",
      Project_Name: "Petani Substation",
      Progress_Percent: 75.5,
      Budget_Used_Percent: 72.3,
      Budget_Total: 15000000,
      Labor_Hours: 1200,
      Equipment_Hours: 450,
      Safety_Score: 95,
      Quality_Score: 88,
      Weather_Impact: "Minor",
      Delay_Days: 2,
    },
    {
      Date: "2024-01-15",
      Project_ID: "PROJ002",
      Project_Name: "Menggala Substation",
      Progress_Percent: 68.2,
      Budget_Used_Percent: 65.1,
      Budget_Total: 12500000,
      Labor_Hours: 1500,
      Equipment_Hours: 500,
      Safety_Score: 92,
      Quality_Score: 90,
      Weather_Impact: "None",
      Delay_Days: 0,
    },
  ],
  sCurveData: [],
  safety: [],
  issues: [],
  plans: [],
  documents: [],
  permits: [],
};

// ========================================================
// Init App
// ========================================================
document.addEventListener("DOMContentLoaded", () => {
  initializeApp();
});

function initializeApp() {
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
  showNotification(
    "Dashboard Pertamina berhasil dimuat! Upload file Excel untuk data terlengkap.",
    "success"
  );
}

// ========================================================
// Navigation
// ========================================================
function setupNavigation() {
  const navLinks = document.querySelectorAll(".nav-link");
  navLinks.forEach((link) => {
    link.addEventListener("click", function (e) {
      e.preventDefault();
      const targetSection = this.getAttribute("data-section");
      switchSection(targetSection);
      navLinks.forEach((nav) => nav.classList.remove("active"));
      this.classList.add("active");
    });
  });
}

function switchSection(sectionName) {
  document
    .querySelectorAll(".content-section")
    .forEach((s) => s.classList.remove("active"));
  const target = document.getElementById(sectionName);
  if (target) {
    target.classList.add("active");
    AppState.currentSection = sectionName;
    if (sectionName === "overview") updateOverviewContent();
    if (sectionName === "scheduler") updateSchedulerContent();
    if (sectionName === "documents") renderDocuments();
    if (sectionName === "permits") updatePermitsContent();
  }
}

// ========================================================
// Overview Section
// ========================================================
function initializeOverview() {
  updateOverviewContent();
}
function updateOverviewContent() {
  updateKPICards();
}
function updateKPICards() {
  const allProjects = AppState.data.projects;
  document.getElementById("totalProjects").textContent = allProjects.length;
}

// ========================================================
// Excel Upload (ringkas untuk contoh)
// ========================================================
function initializeExcelManagement() {
  setupFileUpload();
}
function setupFileUpload() {
  const fileInput = document.getElementById("fileInput");
  fileInput.addEventListener("change", function (e) {
    if (e.target.files.length > 0) handleFileUpload(e.target.files[0]);
  });
}
function handleFileUpload(file) {
  showNotification(`File ${file.name} berhasil diupload`, "success");
}

// ========================================================
// Scheduler
// ========================================================
function initializeScheduler() {
  updateSchedulerContent();
}
function updateSchedulerContent() {
  // contoh update tabel
  updateIssuesTable();
}
function updateIssuesTable() {
  const tbody = document.querySelector("#issuesTable tbody");
  tbody.innerHTML =
    "<tr><td colspan='7'>Tidak ada data issue tersedia</td></tr>";
}

// ========================================================
// Permits & Safety
// ========================================================
function initializePermits() {
  updatePermitsContent();
}
function updatePermitsContent() {
  document.getElementById("totalManpower").textContent = "0";
}

// ========================================================
// Documents
// ========================================================
function initializeDocuments() {
  renderDocuments();
}
function renderDocuments() {
  document.getElementById("documentsGrid").innerHTML =
    "<div class='no-data-message'>Tidak ada dokumen</div>";
}

// ========================================================
// Reports / PPTX
// ========================================================
function initializeReports() {
  setupReportGeneration();
  updateReportsTable();
}
function setupReportGeneration() {
  const btn = document.getElementById("generateReport");
  if (btn) btn.addEventListener("click", generateReport);
}
function updateReportsTable() {
  const tbody = document.querySelector("#reportsTable tbody");
  tbody.innerHTML = "";
}

// ==========================================
// REPORT GENERATOR FIXED
// ==========================================
const ASSETS = {
  cover: "assets/cover_weekly.png",
  logoDanan: "assets/logo_danan.png",
  logoPertamina: "assets/logo_pertamina.png",
};
function absPath(p) { return new URL(p, document.baseURI).href; }
async function loadAsDataUrl(url) {
  const res = await fetch(url, { cache: "no-cache" });
  const blob = await res.blob();
  return new Promise((resolve) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.readAsDataURL(blob);
  });
}

async function generateReport() {
  const type = document.getElementById("reportType").value;
  const startDate = document.getElementById("startDate").value;
  const endDate = document.getElementById("endDate").value;
  const project = document.getElementById("reportProject").value;

  await generatePowerPointReport(type, project, startDate, endDate);
}

async function generatePowerPointReport(type, project, startDate, endDate) {
  try {
    if (typeof PptxGenJS === "undefined") {
      showNotification("Library PPTX tidak ada, fallback digunakan", "warning");
      generateFallbackReport(type, project, startDate, endDate);
      return;
    }

    const [coverImg, logoDanan, logoPertamina] = await Promise.all([
      loadAsDataUrl(absPath(ASSETS.cover)),
      loadAsDataUrl(absPath(ASSETS.logoDanan)),
      loadAsDataUrl(absPath(ASSETS.logoPertamina)),
    ]);

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Pertamina Construction Dashboard";

    const reportTitle = `Construction ${type} Report`;

    // Cover
    {
      const s = pptx.addSlide();
      s.addImage({ data: coverImg, x: 0, y: 0, w: 13.33, h: 7.5 });
      s.addImage({ data: logoDanan, x: 1, y: 0.7, w: 3, h: 1 });
      s.addImage({ data: logoPertamina, x: 9.8, y: 0.7, w: 2.5, h: 1 });
      s.addText(reportTitle, {
        x: 1.3, y: 2, w: 10, h: 0.8,
        fontSize: 32, bold: true, color: "FFFFFF",
      });
    }

    // Executive Summary
    {
      const s = pptx.addSlide();
      s.addImage({ data: logoDanan, x: 0.6, y: 0.4, w: 2.5, h: 0.8 });
      s.addImage({ data: logoPertamina, x: 10.2, y: 0.4, w: 2.2, h: 0.8 });
      s.addText("Executive Summary - (dummy data)", { x: 1, y: 2, w: 10, h: 1 });
    }

    await pptx.writeFile({
      fileName: `Pertamina_${type}_Report_${startDate}_${endDate}.pptx`,
    });
    showNotification("Report berhasil diunduh!", "success");
  } catch (err) {
    console.error("PPT error:", err);
    generateFallbackReport(type, project, startDate, endDate);
  }
}

function generateFallbackReport(type, project, startDate, endDate) {
  const txt = `Pertamina ${type} Report (${startDate} - ${endDate})`;
  const blob = new Blob([txt], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "fallback_report.txt";
  a.click();
}

// ========================================================
// Chatbot
// ========================================================
function initializeChatbot() {
  const chatbot = document.getElementById("chatbot");
  if (chatbot) chatbot.classList.add("collapsed");
}

// ========================================================
// Notifications
// ========================================================
function showNotification(msg, type = "info") {
  const box = document.getElementById("notifications");
  const n = document.createElement("div");
  n.className = `notification ${type}`;
  n.textContent = msg;
  box.appendChild(n);
  setTimeout(() => n.remove(), 4000);
}
