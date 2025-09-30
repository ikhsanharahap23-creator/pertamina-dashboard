// ========================================================
// Global Application State
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

// Static Project ID Mapping
const PROJECT_ID_MAPPING = {
  PROJ001: "Petani Substation",
  PROJ002: "Menggala Substation",
  PROJ003: "Nella Substation",
  PROJ004: "Bangko Substation",
  PROJ005: "Balam SS",
  PROJ006: "Sintong SS",
  PROJ007: "OKB Substation",
};

// Dynamic maps (rebuilt from sheet Projects)
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
function getProjectName(id) {
  return _idToName[id] || id || "Unknown Project";
}
function getProjectId(name) {
  if (!name || name === "all") return null;
  return _nameToId[name] || null;
}
function resolveProjectId(row) {
  return (
    row.Project_ID ||
    getProjectId(row.Project_Name) ||
    "PROJ_UNKNOWN"
  );
}
function getStatusClass(status) {
  if (!status) return "info";
  const s = status.toLowerCase();
  if (s.includes("open") || s.includes("planned")) return "info";
  if (s.includes("progress")) return "warning";
  if (s.includes("completed") || s.includes("closed")) return "success";
  return "info";
}

// ========================================================
// Initialize Application
// ========================================================
document.addEventListener("DOMContentLoaded", initializeApp);
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
  showNotification("Dashboard berhasil dimuat!", "success");
}

// ========================================================
// Navigation
// ========================================================
function setupNavigation() {
  const navLinks = document.querySelectorAll(".nav-link");
  navLinks.forEach((link) => {
    link.addEventListener("click", function (e) {
      e.preventDefault();
      const target = this.getAttribute("data-section");
      switchSection(target);
      navLinks.forEach((n) => n.classList.remove("active"));
      this.classList.add("active");
    });
  });
}
function switchSection(section) {
  document.querySelectorAll(".content-section").forEach((s) =>
    s.classList.remove("active")
  );
  const target = document.getElementById(section);
  if (target) {
    target.classList.add("active");
    AppState.currentSection = section;
    if (section === "overview") updateOverviewContent();
    if (section === "scheduler") updateSchedulerContent();
    if (section === "documents") renderDocuments();
    if (section === "permits") updatePermitsContent();
  }
}

// ========================================================
// Overview (KPI & Charts)
// ========================================================
function initializeOverview() {
  updateOverviewContent();
}
function updateOverviewContent() {
  const projects = AppState.data.projects;
  document.getElementById("totalProjects").textContent = projects.length;
  if (projects.length > 0) {
    const avgSafety = Math.round(
      projects.reduce((s, p) => s + (p.Safety_Score || 0), 0) / projects.length
    );
    document.getElementById("safetyScore").textContent = avgSafety + "%";
  }
}

// ========================================================
// Excel Upload (Simplified)
// ========================================================
function initializeExcelManagement() {
  const fileInput = document.getElementById("fileInput");
  if (fileInput) {
    fileInput.addEventListener("change", (e) => {
      if (e.target.files.length > 0) {
        handleFileUpload(e.target.files[0]);
      }
    });
  }
}
function handleFileUpload(file) {
  showNotification(`File ${file.name} berhasil diupload`, "success");
}

// ========================================================
// Scheduler (S-Curve & Issues/Plans)
// ========================================================
function initializeScheduler() {
  updateSchedulerContent();
}
function updateSchedulerContent() {
  updateIssuesTable();
}
function updateIssuesTable() {
  const tbody = document.querySelector("#issuesTable tbody");
  tbody.innerHTML =
    "<tr><td colspan='7' class='table-loading'>Tidak ada data issue</td></tr>";
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
// Reports (PowerPoint Generator)
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

const ASSETS = {
  cover: "assets/cover_weekly.png",
  logoDanan: "assets/logo_danan.png",
  logoPertamina: "assets/logo_pertamina.png",
};
function absPath(p) {
  return new URL(p, document.baseURI).href;
}
async function loadAsDataUrl(url) {
  const res = await fetch(url, { cache: "no-cache" });
  const blob = await res.blob();
  return new Promise((resolve) => {
    const fr = new FileReader();
    fr.onload = () => resolve(fr.result);
    fr.readAsDataURL(blob);
  });
}

// === Main Generate Report ===
async function generateReport() {
  const type = document.getElementById("reportType").value;
  const start = document.getElementById("startDate").value;
  const end = document.getElementById("endDate").value;
  const project = document.getElementById("reportProject").value;
  await generatePowerPointReport(type, project, start, end);
}
async function generatePowerPointReport(type, project, startDate, endDate) {
  try {
    if (typeof PptxGenJS === "undefined") {
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

    // Slide 1: Cover
    const s1 = pptx.addSlide();
    s1.addImage({ data: coverImg, x: 0, y: 0, w: 13.33, h: 7.5 });
    s1.addImage({ data: logoDanan, x: 1, y: 0.7, w: 3, h: 1 });
    s1.addImage({ data: logoPertamina, x: 9.8, y: 0.7, w: 2.5, h: 1 });
    s1.addText(`Construction ${type} Report`, {
      x: 1.3,
      y: 2,
      w: 10,
      h: 0.8,
      fontSize: 32,
      bold: true,
      color: "FFFFFF",
    });
    s1.addText(project === "all" ? "Semua Proyek" : project, {
      x: 1.3,
      y: 3,
      w: 10,
      h: 0.6,
      fontSize: 24,
      bold: true,
      color: "FFFFFF",
    });
    s1.addText(`Period: ${startDate} - ${endDate}`, {
      x: 1.3,
      y: 3.8,
      w: 10,
      h: 0.5,
      fontSize: 20,
      color: "FFFFFF",
    });

    // Slide 2: Executive Summary
    const s2 = pptx.addSlide();
    s2.addText("Executive Summary", {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 0.8,
      fontSize: 28,
      bold: true,
      color: "184E5C",
    });
    const totalProjects = AppState.data.projects.length;
    const activePermits = AppState.data.permits.filter(
      (p) => p.Status === "Open"
    ).length;
    s2.addText(`Total Projects: ${totalProjects}`, {
      x: 0.8,
      y: 1.5,
      fontSize: 20,
    });
    s2.addText(`Active Permits: ${activePermits}`, {
      x: 0.8,
      y: 2.2,
      fontSize: 20,
    });

    // Slide 3: Project Progress Table
    const s3 = pptx.addSlide();
    s3.addText("Project Progress", {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 0.8,
      fontSize: 28,
      bold: true,
      color: "184E5C",
    });
    const rows = [
      [
        { text: "Project Name", options: { bold: true, color: "FFFFFF" } },
        { text: "Progress %", options: { bold: true, color: "FFFFFF" } },
        { text: "Safety Score", options: { bold: true, color: "FFFFFF" } },
      ],
    ];
    (AppState.data.projects || []).forEach((p) => {
      rows.push([
        p.Project_Name,
        `${p.Progress_Percent || 0}%`,
        `${p.Safety_Score || 0}%`,
      ]);
    });
    s3.addTable(rows, {
      x: 0.5,
      y: 1.5,
      w: 12,
      border: { type: "solid", color: "184E5C", pt: 1 },
      fill: "F2F2F2",
      fontSize: 14,
      colW: [6, 3, 3],
    });

    // Slide 4: Safety Overview
    const s4 = pptx.addSlide();
    s4.addText("Safety Overview", {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 0.8,
      fontSize: 28,
      bold: true,
      color: "184E5C",
    });
    const totalManpower = AppState.data.safety.reduce(
      (sum, item) => sum + (item.Total_Manpower || 0),
      0
    );
    s4.addText(`Total Manpower: ${totalManpower}`, {
      x: 0.8,
      y: 1.5,
      fontSize: 20,
    });
    s4.addText(`Safe Man Hours: ${
      AppState.data.safety.reduce(
        (sum, item) => sum + (item.Safe_Man_Hours || 0),
        0
      )
    }`, {
      x: 0.8,
      y: 2.2,
      fontSize: 20,
    });

    await pptx.writeFile({
      fileName: `Pertamina_${type}_Report_${startDate}_${endDate}.pptx`,
    });
    showNotification("Report berhasil diunduh!", "success");
  } catch (err) {
    console.error("PPT Error:", err);
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
// Project Filters
// ========================================================
function setupProjectFilters() {
  const reportFilter = document.getElementById("reportProject");
  if (reportFilter) {
    reportFilter.innerHTML = "<option value='all'>Semua Proyek</option>";
    (AppState.data.projects || []).forEach((p) => {
      const o = document.createElement("option");
      o.value = p.Project_Name;
      o.textContent = p.Project_Name;
      reportFilter.appendChild(o);
    });
  }
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
function showNotification(message, type = "info") {
  const container = document.getElementById("notifications");
  const n = document.createElement("div");
  n.className = `notification ${type}`;
  n.textContent = message;
  container.appendChild(n);
  setTimeout(() => n.remove(), 4000);
}
