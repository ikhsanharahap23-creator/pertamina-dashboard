// ====== Pertamina Dashboard - PPT Template Integration (app.fixed.js) ======
// Drop-in script to be loaded AFTER your existing app.js, or replace the same
// helper/function names in your current app.js. This file adds a Slide Master
// that mimics the uploaded 'Laporan PPT.pptx' corporate style and overrides
// generatePowerPointReport() to always use that style.

// ---- Brand configuration (tweak if needed) ----
const BRAND = {
  name: "Pertamina",
  // Colors derived from template (adjust hex if your PPT uses slightly different shades)
  primary: "003366",   // Dark blue header/footer bar
  secondary: "21808D", // Teal accent used in dashboard
  accent: "FFC185",    // Soft orange accent line
  text: "2B2B2B",
  lightText: "FFFFFF",
  gray: "ECEFF1",
  // Path to logo image (put logo.png at same folder as index.html)
  logoPath: "logo.png"
};

// ---- Master definitions to mimic the PPT template ----
function definePertaminaMasters(pptx) {
  // COVER MASTER: Blue top block with right-aligned logo (to match the template look)
  pptx.defineSlideMaster({
    title: "PERTAMINA_COVER",
    background: { color: "FFFFFF" },
    objects: [
      // Top color block
      { rect: { x: 0, y: 0, w: "100%", h: 3.2, fill: BRAND.primary } },
      // Accent thin strip
      { rect: { x: 0.8, y: 3.1, w: 4.5, h: 0.12, fill: BRAND.accent } },
      // Logo on right
      { image: { path: BRAND.logoPath, x: 9.2, y: 0.5, w: 1.2, h: 1.2 } },
    ]
  });

  // CONTENT MASTER: Header bar with logo (left), footer bar with product name and slide number
  pptx.defineSlideMaster({
    title: "PERTAMINA_MASTER",
    background: { color: "FFFFFF" },
    objects: [
      // Header bar
      { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: BRAND.primary } },
      // Logo on header (left)
      { image: { path: BRAND.logoPath, x: 0.25, y: 0.08, w: 0.8, h: 0.8 } },
      // Footer bar
      { rect: { x: 0, y: 6.8, w: "100%", h: 0.4, fill: BRAND.primary } },
      // Footer text (left)
      { text: { text: "Pertamina Project Dashboard", options: { x: 0.6, y: 6.82, fontSize: 11, color: BRAND.lightText } } },
      // Footer date (right)
      { text: { text: () => new Date().toLocaleDateString("id-ID"), options: { x: 9.4, y: 6.82, fontSize: 11, color: BRAND.lightText, align: "right" } } },
    ],
    slideNumber: { x: 10.35, y: 6.82, color: BRAND.lightText, fontFace: "Arial", fontSize: 11 }
  });
}

// ---- Small helpers ----
async function addLogo(slide, x, y, w, h) {
  try { slide.addImage({ path: BRAND.logoPath, x, y, w, h }); } catch (e) {}
}

function addSectionTitle(slide, title, subtitle, y = 0.9) {
  slide.addText(title, { x: 0.85, y, fontSize: 22, bold: true, color: BRAND.primary, fontFace: "Arial" });
  if (subtitle) slide.addText(subtitle, { x: 0.85, y: y + 0.55, fontSize: 12, color: "666666", fontFace: "Arial" });
}

function kpiChip(slide, x, y, label, value) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w: 3.15, h: 1.0, fill: "FFFFFF", line: { color: BRAND.gray },
    rectRadius: 10, shadow: { type: "outer", blur: 2, color: "000000", opacity: 0.12 }
  });
  slide.addText(label, { x: x + 0.28, y: y + 0.16, fontSize: 12, color: "#6B7280", fontFace: "Arial" });
  slide.addText(String(value), { x: x + 0.28, y: y + 0.48, fontSize: 20, bold: true, color: BRAND.secondary, fontFace: "Arial" });
}

function tableTheme() {
  return {
    x: 0.85, y: 1.85, w: 9.1,
    fontFace: "Arial",
    border: { pt: 0.75, color: "DDDDDD" },
    fill: "FFFFFF",
    valign: "middle",
    autoPage: true,
    autoPageRepeatHeader: true,
    rowH: 0.46,
    margin: 0.05,
    color: "#374151",
    master: "PERTAMINA_MASTER",
    header: true,
    fillHdr: BRAND.primary,
    colorHdr: BRAND.lightText,
    fontSizeHdr: 12,
    bold: true
  };
}

// ---- MAIN: Override generator to use the template masters ----
function generatePowerPointReport(type, project, startDate, endDate) {
  try {
    if (typeof PptxGenJS === "undefined") {
      showNotification("PowerPoint library not loaded. Generating fallback report...", "warning");
      generateFallbackReport(type, project, startDate, endDate);
      return;
    }

    // Init & masters
    window.pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE"; // 16:9
    definePertaminaMasters(pptx);

    pptx.author = "Pertamina Construction Dashboard";
    pptx.company = "Pertamina";
    pptx.title = `Laporan ${type.charAt(0).toUpperCase() + type.slice(1)} - Pertamina`;

    // ===== DATA PREP =====
    const filteredProjects = project === "all"
      ? AppState.data.projects
      : AppState.data.projects.filter(p => p.Project_Name === project);

    const avgProgress = filteredProjects.length
      ? Math.round(filteredProjects.reduce((s, p) => s + (p.Progress_Percent || 0), 0) / filteredProjects.length)
      : 0;
    const avgSafety = filteredProjects.length
      ? Math.round(filteredProjects.reduce((s, p) => s + (p.Safety_Score || 0), 0) / filteredProjects.length)
      : 0;
    const activePermits = AppState.data.permits.filter(p => p.Status === "Open").length;

    // ===== SLIDE 1 — COVER =====
    let cover = pptx.addSlide({ masterName: "PERTAMINA_COVER" });
    cover.addText(`LAPORAN ${type.toUpperCase()}`, { x: 0.85, y: 1.8, fontSize: 34, bold: true, color: BRAND.lightText, fontFace: "Arial" });
    cover.addText(project === "all" ? "Semua Proyek" : project, { x: 0.85, y: 2.6, fontSize: 20, color: BRAND.lightText, fontFace: "Arial" });
    cover.addText(`Periode: ${startDate} — ${endDate}`, { x: 0.85, y: 3.2, fontSize: 14, color: BRAND.lightText, fontFace: "Arial" });
    cover.addText(`Generated: ${new Date().toLocaleString("id-ID")}`, { x: 0.85, y: 3.7, fontSize: 12, color: BRAND.lightText, fontFace: "Arial" });

    // ===== SLIDE 2 — EXECUTIVE SUMMARY =====
    let sum = pptx.addSlide({ masterName: "PERTAMINA_MASTER" });
    addSectionTitle(sum, "Executive Summary", new Date().toLocaleDateString("id-ID"));

    kpiChip(sum, 0.85, 1.55, "Total Proyek", filteredProjects.length);
    kpiChip(sum, 4.05, 1.55, "Permit Aktif", activePermits);
    kpiChip(sum, 7.25, 1.55, "Rata-rata Safety", `${avgSafety}%`);

    kpiChip(sum, 0.85, 2.75, "Rata-rata Progress", `${avgProgress}%`);
    kpiChip(sum, 4.05, 2.75, "Total Issues", AppState.data.issues.length);
    kpiChip(sum, 7.25, 2.75, "Total Plans", AppState.data.plans.length);

    const narrative = `Rata-rata progress ${avgProgress}% dengan skor safety ${avgSafety}%. ` +
      `Terdapat ${activePermits} permit aktif, ${AppState.data.issues.length} issue, ` +
      `dan ${AppState.data.plans.length} rencana aktif.`;
    sum.addText(narrative, { x: 0.85, y: 4.05, w: 9.6, h: 1.4, fontSize: 12, color: "#374151", fontFace: "Arial" });

    // ===== SLIDE 3 — PROJECT DETAILS TABLE =====
    let sProjects = pptx.addSlide({ masterName: "PERTAMINA_MASTER" });
    addSectionTitle(sProjects, "Project Details", "Ringkasan per proyek");

    const tableHeader = [
      { text: "Project Name", options: { bold: true, color: "FFFFFF", fontFace: "Arial" } },
      { text: "Progress %", options: { bold: true, color: "FFFFFF", fontFace: "Arial" } },
      { text: "Safety Score", options: { bold: true, color: "FFFFFF", fontFace: "Arial" } },
      { text: "Budget Used %", options: { bold: true, color: "FFFFFF", fontFace: "Arial" } }
    ];

    const tableRows = filteredProjects.slice(0, 40).map(p => ([
      p.Project_Name?.replace(" Substation", "") || "N/A",
      `${p.Progress_Percent ?? 0}%`,
      `${p.Safety_Score ?? 0}%`,
      `${p.Budget_Used_Percent ?? 0}%`
    ]));

    sProjects.addTable([tableHeader, ...tableRows], tableTheme());

    // ===== SLIDE 4 — SAFETY SUMMARY =====
    let sSafety = pptx.addSlide({ masterName: "PERTAMINA_MASTER" });
    addSectionTitle(sSafety, "Safety Summary", "Manpower, Safe Hours, dan Insiden");

    const totalManpower = AppState.data.safety.reduce((s, i) => s + (i.Total_Manpower || 0), 0);
    const safeManHours = AppState.data.safety.reduce((s, i) => s + (i.Safe_Man_Hours || 0), 0);
    const incidents = AppState.data.safety.reduce((s, i) =>
      s + (i.Fatal_Accidents || 0) + (i.Lost_Time_Injuries || 0) + (i.Medical_Treatment_Cases || 0) + (i.First_Aid_Cases || 0), 0);

    kpiChip(sSafety, 0.85, 1.55, "Total Manpower", totalManpower.toLocaleString("id-ID"));
    kpiChip(sSafety, 4.05, 1.55, "Safe Man Hours", safeManHours.toLocaleString("id-ID"));
    kpiChip(sSafety, 7.25, 1.55, "Total Insiden", incidents);

    sSafety.addText("Catatan:", { x: 0.85, y: 2.85, fontSize: 12, color: "#6B7280", fontFace: "Arial" });
    sSafety.addText("Pantau tren Near Miss, First Aid, dan LTI pada dashboard untuk area perbaikan prioritas.",
      { x: 0.85, y: 3.2, w: 9.6, h: 1.2, fontSize: 12, color: "#374151", fontFace: "Arial" });

    // ===== SAVE =====
    const fileName = `Pertamina_${type}_Report_${new Date().toISOString().slice(0,10)}.pptx`;
    pptx.writeFile({ fileName });
    showNotification("PowerPoint report berhasil didownload dengan template!", "success");
  } catch (error) {
    console.error("Error generating PowerPoint:", error);
    showNotification("Error generating PowerPoint: " + error.message, "error");
    // Fallback
    if (typeof generateFallbackReport === "function") {
      generateFallbackReport(type, project, startDate, endDate);
    }
  }
}

// ====== End of app.fixed.js ======
