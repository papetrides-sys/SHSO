/**
 * app.js
 * ------
 * Navigation + UI wiring for all views:
 *   Home → Statistics | Update Tickets → Home
 */

import { extractIds } from "./filter.js";

// ── Palette for charts ───────────────────────────────────────────────────────
const COLORS = ["#3b82d4", "#7c5cd8", "#22c55e", "#f59e0b", "#ef4444", "#06b6d4", "#ec4899", "#84cc16"];

// ── View switching ───────────────────────────────────────────────────────────
function showView(id) {
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  document.getElementById(id).classList.add("active");
}

// ── Bootstrap after DOM ready ────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {

  // ── Shared file state ──────────────────────────────────────────────────────
  // rawRows: all rows as parsed (including header if present)
  // sharedRows: data rows after optional header skip
  let rawRows    = [];
  let sharedRows = [];
  let sharedFileName = "";

  const homeDropZone      = document.getElementById("homeDropZone");
  const homeFileInput     = document.getElementById("homeFileInput");
  const homeHeaderChk     = document.getElementById("homeHeaderCheckbox");
  const homeFileName      = document.getElementById("homeFileName");
  const homeErrorBox      = document.getElementById("homeErrorBox");
  const homeTagline       = document.getElementById("homeTagline");
  const tileGrid          = document.getElementById("tileGrid");

  // ── Baseline Comparison file state (completely separate from sharedRows) ──
  let bcRawRows    = [];
  let bcRows       = [];      // after optional header skip
  let bcFileName   = "";

  const bcDropZone       = document.getElementById("bcDropZone");
  const bcFileInput      = document.getElementById("bcFileInput");
  const bcHeaderChk      = document.getElementById("bcHeaderCheckbox");
  const bcFileNameEl     = document.getElementById("bcFileName");
  const bcErrorBox       = document.getElementById("bcErrorBox");
  const bcReadyMsg       = document.getElementById("bcReadyMsg");
  const bcCurrentFileLbl = document.getElementById("bcCurrentFileLabel");

  function parseSharedFile(file) {
    homeErrorBox.hidden = true; homeErrorBox.textContent = "";

    if (!file.name.endsWith(".xlsx")) {
      homeErrorBox.textContent = "Only .xlsx files are supported.";
      homeErrorBox.hidden = false;
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        sharedFileName = file.name;
        applyHeaderSkip();
        homeFileName.textContent = file.name;
        homeTagline.textContent  = `${file.name} — select a tool`;
        // Enable tiles
        tileGrid.style.opacity      = "1";
        tileGrid.style.pointerEvents = "auto";
      } catch (err) {
        homeErrorBox.textContent = "Failed to parse the file: " + err.message;
        homeErrorBox.hidden = false;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function applyHeaderSkip() {
    sharedRows = (homeHeaderChk.checked && rawRows.length) ? rawRows.slice(1) : rawRows;
  }

  homeFileInput.addEventListener("change", () => {
    if (homeFileInput.files.length) parseSharedFile(homeFileInput.files[0]);
  });
  homeDropZone.addEventListener("dragover",  e => { e.preventDefault(); homeDropZone.classList.add("drag-over"); });
  homeDropZone.addEventListener("dragleave", ()  => homeDropZone.classList.remove("drag-over"));
  homeDropZone.addEventListener("drop", e => {
    e.preventDefault();
    homeDropZone.classList.remove("drag-over");
    if (e.dataTransfer.files[0]) parseSharedFile(e.dataTransfer.files[0]);
  });
  homeDropZone.addEventListener("click", () => homeFileInput.click());
  homeHeaderChk.addEventListener("change", () => {
    if (!rawRows.length) return;
    applyHeaderSkip();
    // Re-initialise whichever sub-app is currently open
    initStats();
    initUpdate();
  });

  // ── Baseline Comparison file handling ─────────────────────────────────────

  // DOM refs for baseline dashboard sections
  const bcFiltersEl       = document.getElementById("bcFilters");
  const bcCurrentSection  = document.getElementById("bcCurrentSection");
  const bcBaselineSection = document.getElementById("bcBaselineSection");
  const bcKpiCurrent      = document.getElementById("bcKpiCurrent");
  const bcKpiBaseline     = document.getElementById("bcKpiBaseline");
  const bcCurrentFileTag  = document.getElementById("bcCurrentFileTag");
  const bcBaselineFileTag = document.getElementById("bcBaselineFileTag");
  const bcProgressSection  = document.getElementById("bcProgressSection");
  const bcProgressGrid     = document.getElementById("bcProgressGrid");
  const bcDeliveredSection = document.getElementById("bcDeliveredSection");
  const bcDeliveredCount   = document.getElementById("bcDeliveredCount");
  const bcDeliveredTable   = document.getElementById("bcDeliveredTable");
  const bcReturnedSection  = document.getElementById("bcReturnedSection");
  const bcReturnedCount    = document.getElementById("bcReturnedCount");
  const bcReturnedTable    = document.getElementById("bcReturnedTable");
  const bcResolvedSection  = document.getElementById("bcResolvedSection");
  const bcResolvedCount    = document.getElementById("bcResolvedCount");
  const bcResolvedTable    = document.getElementById("bcResolvedTable");
  const bcRespSection      = document.getElementById("bcRespSection");
  const bcRespContent      = document.getElementById("bcRespContent");
  const bcRespFileTag      = document.getElementById("bcRespFileTag");
  const bcBaseRespSection  = document.getElementById("bcBaseRespSection");
  const bcBaseRespContent  = document.getElementById("bcBaseRespContent");
  const bcBaseRespFileTag  = document.getElementById("bcBaseRespFileTag");
  const bcHint             = document.getElementById("bcHint");

  // Baseline-scoped filter selections (union of both files' values)
  let bcSelectedGroups      = new Set();
  let bcSelectedTypes       = new Set();
  let bcSelectedSeverities  = new Set();
  let bcSelectedFixVersions = new Set();

  /** Union unique values from both files for a given column index */
  function bcUnionValues(colIndex) {
    const s = new Set();
    for (const row of [...sharedRows, ...bcRows]) {
      if (row.length > colIndex && row[colIndex] != null)
        s.add(String(row[colIndex]).trim());
    }
    return s;
  }

  function bcUnionSeverities() {
    const s = new Set();
    for (const row of [...sharedRows, ...bcRows]) {
      if (row.length > 11) s.add(severityLabel(row[11]));
    }
    return s;
  }

  function bcUnionLastCol() {
    const s = new Set();
    for (const row of [...sharedRows, ...bcRows]) {
      if (row.length > 0 && row[row.length - 1] != null)
        s.add(String(row[row.length - 1]).trim());
    }
    return s;
  }

  /** Filter a set of rows using the BC filter selections */
  function bcApplyFilters(rows) {
    return rows.filter(row => {
      const group      = row.length > 4 && row[4] != null ? String(row[4]).trim() : "(blank)";
      const type       = row.length > 5 && row[5] != null ? String(row[5]).trim() : "(blank)";
      const severity   = severityLabel(row.length > 11 ? row[11] : null);
      const fixVersion = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]).trim() : "(blank)";
      return bcSelectedGroups.has(group) && bcSelectedTypes.has(type)
          && bcSelectedSeverities.has(severity) && bcSelectedFixVersions.has(fixVersion);
    });
  }

  function buildBcFilters() {
    bcFiltersEl.innerHTML = "";
    if (!sharedRows.length && !bcRows.length) return;

    const SEVERITY_ORDER = ["Critical", "Major", "Medium", "Low"];

    function addChipGroup(label, selectedSet, values, setterFn) {
      bcFiltersEl.appendChild(chipGroup(
        label, selectedSet, [...values].sort(),
        val => {
          selectedSet.has(val) ? selectedSet.delete(val) : selectedSet.add(val);
          refreshBcKpis();
        },
        () => {
          if (selectedSet.size === values.size) { selectedSet.clear(); }
          else { setterFn(new Set(values)); }
          buildBcFilters();
          refreshBcKpis();
        }
      ));
    }

    const allGroups   = bcUnionValues(4);
    const allTypes    = bcUnionValues(5);
    const allSevs     = bcUnionSeverities();
    const allFix      = bcUnionLastCol();
    const sevOrdered  = SEVERITY_ORDER.filter(l => allSevs.has(l));
    for (const l of allSevs) { if (!SEVERITY_ORDER.includes(l)) sevOrdered.push(l); }

    addChipGroup("Group",       bcSelectedGroups,      allGroups, v => { bcSelectedGroups      = v; });
    addChipGroup("Ticket Type", bcSelectedTypes,       allTypes,  v => { bcSelectedTypes       = v; });
    bcFiltersEl.appendChild(chipGroup(
      "Severity", bcSelectedSeverities, sevOrdered,
      val => { bcSelectedSeverities.has(val) ? bcSelectedSeverities.delete(val) : bcSelectedSeverities.add(val); refreshBcKpis(); },
      () => {
        if (bcSelectedSeverities.size === allSevs.size) { bcSelectedSeverities.clear(); }
        else { bcSelectedSeverities = new Set(allSevs); }
        buildBcFilters(); refreshBcKpis();
      }
    ));
    addChipGroup("Fix Version", bcSelectedFixVersions, allFix,   v => { bcSelectedFixVersions  = v; });
  }

  function refreshBcKpis() {
    const hasCurrentFile  = sharedRows.length > 0;
    const hasBaselineFile = bcRows.length > 0;
    const hasBoth         = hasCurrentFile && hasBaselineFile;

    bcHint.hidden = hasCurrentFile || hasBaselineFile;
    bcCurrentSection.hidden  = !hasCurrentFile;
    bcBaselineSection.hidden = !hasBaselineFile;
    bcProgressSection.hidden  = !hasBoth;
    bcDeliveredSection.hidden = !hasBoth;
    bcReturnedSection.hidden  = !hasBoth;
    bcResolvedSection.hidden  = !hasBoth;
    bcRespSection.style.display  = (hasCurrentFile || hasBaselineFile) ? "" : "none";
    bcRespSection.hidden         = false;
    bcBaseRespSection.hidden     = !hasBaselineFile;

    let currentKpis  = null;
    let baselineKpis = null;
    let currentTesting  = 0;
    let baselineTesting = 0;
    let filteredCurrent  = [];
    let filteredBaseline = [];

    if (hasCurrentFile) {
      filteredCurrent = bcApplyFilters(sharedRows);
      currentKpis     = computeKpis(filteredCurrent);
      currentTesting  = filteredCurrent.filter(r => (r[2] != null ? String(r[2]).trim().toUpperCase() : "") === "TESTING").length;
      bcKpiCurrent.innerHTML = "";
      bcKpiCurrent.appendChild(renderKpiStrip(currentKpis, filteredCurrent));
      bcCurrentFileTag.textContent = `— ${sharedFileName}`;

      // Current Tickets Responsibility pie (exclude resolved)
      const currentUnresolved = filteredCurrent.filter(r => { const s = r[2] != null ? String(r[2]).trim().toUpperCase() : ""; return s !== "DONE" && s !== "CLOSED"; });
      bcRespContent.innerHTML = "";
      bcRespContent.appendChild(pieBlock(
        "Current Tickets Responsibility",
        groupedTally(currentUnresolved, respGroupLabel),
        currentUnresolved,
        respGroupLabel,
        lbl => RESP_GROUP_COLORS[lbl] ?? COLORS[0]
      ));
      bcRespFileTag.textContent = `— ${sharedFileName}`;
    }
    if (hasBaselineFile) {
      filteredBaseline = bcApplyFilters(bcRows);
      baselineKpis     = computeKpis(filteredBaseline);
      baselineTesting  = filteredBaseline.filter(r => (r[2] != null ? String(r[2]).trim().toUpperCase() : "") === "TESTING").length;
      bcKpiBaseline.innerHTML = "";
      bcKpiBaseline.appendChild(renderKpiStrip(baselineKpis, filteredBaseline));
      bcBaselineFileTag.textContent = `— ${bcFileName}`;

      // Baseline Tickets Responsibility pie (exclude resolved)
      const baselineUnresolved = filteredBaseline.filter(r => { const s = r[2] != null ? String(r[2]).trim().toUpperCase() : ""; return s !== "DONE" && s !== "CLOSED"; });
      bcBaseRespContent.innerHTML = "";
      bcBaseRespContent.appendChild(pieBlock(
        "Baseline Tickets Responsibility",
        groupedTally(baselineUnresolved, respGroupLabel),
        baselineUnresolved,
        respGroupLabel,
        lbl => RESP_GROUP_COLORS[lbl] ?? COLORS[0]
      ));
      bcBaseRespFileTag.textContent = `— ${bcFileName}`;
    }
    if (hasBoth) {
      renderBcProgress(currentKpis, baselineKpis, currentTesting, baselineTesting);
      renderBcDelivered(filteredCurrent, filteredBaseline);
      renderBcReturned(filteredCurrent, filteredBaseline);
      renderBcResolved(filteredCurrent, filteredBaseline);
    }
  }

  /**
   * Render the Progress comparison grid.
   * For each KPI, show current value, baseline value, delta and a
   * colour-coded badge indicating whether the change is better/worse/neutral.
   *
   * "Better" direction per KPI:
   *   total          → neutral (pure size, no good/bad)
   *   resolved       → higher is better
   *   resolutionRate → higher is better
   *   throughputLast30 → higher is better
   *   openOver30     → lower is better
   */
  function renderBcProgress(cur, base, currentTesting, baselineTesting) {
    bcProgressGrid.innerHTML = "";

    const grid = document.createElement("div");
    grid.className = "bc-progress-grid";

    const defs = [
      {
        label: "Total Tickets",
        cur:   cur.total,
        base:  base.total,
        fmt:   v => String(v),
        direction: "neutral",
      },
      {
        label: "Resolved",
        cur:   cur.resolved,
        base:  base.resolved,
        fmt:   v => String(v),
        direction: "higher-better",
      },
      {
        label: "Resolution Rate",
        cur:   cur.resolutionRate,
        base:  base.resolutionRate,
        fmt:   v => v.toFixed(1) + "%",
        direction: "higher-better",
      },
      {
        label: "Resolved (Last 30 d)",
        cur:   cur.throughputLast30,
        base:  base.throughputLast30,
        fmt:   v => String(v),
        direction: "higher-better",
      },
      {
        label: "Open > 30 Days",
        cur:   cur.openOver30,
        base:  base.openOver30,
        fmt:   v => String(v),
        direction: "lower-better",
      },
      {
        label: "In Testing",
        cur:   currentTesting,
        base:  baselineTesting,
        fmt:   v => String(v),
        direction: "lower-better",
      },
      {
        label: "Tickets on OTE",
        cur:   cur.ticketsOnOte,
        base:  base.ticketsOnOte,
        fmt:   v => String(v),
        direction: "lower-better",
      },
      {
        label: "On SHSO",
        cur:   cur.ticketsOnShso,
        base:  base.ticketsOnShso,
        fmt:   v => String(v),
        direction: "lower-better",
      },
    ];

    for (const def of defs) {
      const delta     = def.cur - def.base;
      const absDelta  = Math.abs(delta);
      const isNum     = def.label !== "Resolution Rate";

      // Format the delta display
      const deltaFmt  = def.label === "Resolution Rate"
        ? (delta >= 0 ? "+" : "") + delta.toFixed(1) + "%"
        : (delta >= 0 ? "+" : "") + String(Math.round(delta));

      // Determine badge class
      let badgeClass = "bc-delta--neutral";
      let arrow = "→";
      if (delta !== 0) {
        const improvement =
          (def.direction === "higher-better" && delta > 0) ||
          (def.direction === "lower-better"  && delta < 0);
        if (improvement) {
          badgeClass = "bc-delta--better";
          arrow = "▲";
        } else if (def.direction !== "neutral") {
          badgeClass = "bc-delta--worse";
          arrow = "▼";
        }
      }

      const card = document.createElement("div");
      card.className = "bc-progress-card";
      if (badgeClass === "bc-delta--better") card.style.background = "#f0fdf4";
      else if (badgeClass === "bc-delta--worse") card.style.background = "#fff0f0";
      card.innerHTML =
        `<div class="bc-progress-kpi-label">${def.label}</div>
         <div class="bc-progress-row">
           <span>Current</span>
           <span class="bc-progress-val">${def.fmt(def.cur)}</span>
         </div>
         <div class="bc-progress-row">
           <span>Baseline</span>
           <span class="bc-progress-val" style="color:#57606a;font-size:0.9rem">${def.fmt(def.base)}</span>
         </div>
         <div class="bc-delta ${badgeClass}">${arrow} ${delta === 0 ? "No change" : deltaFmt}</div>`;

      grid.appendChild(card);
    }

    bcProgressGrid.appendChild(grid);
  }

  /**
   * Render the "Delivered for Testing" section.
   * Items qualify when:
   *   - They appear in the baseline with status TODO or INPROG
   *   - They appear in the current file (matched by ID, col 0) with status TESTING or WAIT-INFO
   * The table shows: ID, Description, Status, Severity, Fix Version.
   * Sort order: Fix Version (ascending) then Severity numeric (1 < 2 < 3 < 4).
   */
  function renderBcDelivered(currentRows, baselineRows) {
    // Build a set of baseline IDs that were TODO or INPROG
    const BASE_STATUSES = new Set(["TODO", "INPROG"]);
    const baselineIds = new Set();
    for (const row of baselineRows) {
      const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      if (BASE_STATUSES.has(st) && row[0] != null) {
        baselineIds.add(String(row[0]).trim());
      }
    }

    // Find current rows whose ID was in baseline TODO/INPROG and are now TESTING or WAIT-INFO
    const CUR_STATUSES = new Set(["TESTING", "WAIT-INFO"]);
    const SEV_ORDER = { "1": 1, "2": 2, "3": 3, "4": 4 };

    let delivered = currentRows.filter(row => {
      const id = row[0] != null ? String(row[0]).trim() : "";
      const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      return id !== "" && baselineIds.has(id) && CUR_STATUSES.has(st);
    });

    // Sort: Fix Version asc (last col), then severity numeric asc
    delivered.sort((a, b) => {
      const fa = a.length > 0 && a[a.length - 1] != null ? String(a[a.length - 1]) : "";
      const fb = b.length > 0 && b[b.length - 1] != null ? String(b[b.length - 1]) : "";
      const fCmp = fa.localeCompare(fb, undefined, { numeric: true, sensitivity: "base" });
      if (fCmp !== 0) return fCmp;
      const sa = SEV_ORDER[String(a.length > 11 && a[11] != null ? a[11] : "").trim()] ?? 99;
      const sb = SEV_ORDER[String(b.length > 11 && b[11] != null ? b[11] : "").trim()] ?? 99;
      return sa - sb;
    });

    // Count badge
    bcDeliveredCount.textContent = delivered.length === 0
      ? "No items moved into testing since baseline."
      : `${delivered.length} item${delivered.length === 1 ? "" : "s"} delivered for testing`;

    // Build table
    bcDeliveredTable.innerHTML =
      `<thead><tr>
        <th>ID</th>
        <th>Description</th>
        <th>Status</th>
        <th>Severity</th>
        <th>Fix Version</th>
      </tr></thead>`;

    if (delivered.length === 0) {
      bcDeliveredTable.innerHTML = "";
      return;
    }

    const tbody = document.createElement("tbody");
    for (const row of delivered) {
      const tr = document.createElement("tr");

      // ID — col 0
      const tdId = document.createElement("td");
      tdId.textContent = row[0] != null ? String(row[0]) : "";
      tr.appendChild(tdId);

      // Description — col 1
      const tdDesc = document.createElement("td");
      tdDesc.className = "bc-dft-td-desc";
      tdDesc.textContent = row[1] != null ? String(row[1]) : "";
      tr.appendChild(tdDesc);

      // Status — col 2
      const tdStatus = document.createElement("td");
      tdStatus.className = "bc-dft-td-status";
      tdStatus.textContent = row[2] != null ? String(row[2]).trim() : "";
      tr.appendChild(tdStatus);

      // Severity — col 11
      const tdSev = document.createElement("td");
      tdSev.className = "bc-dft-td-sev";
      tdSev.textContent = severityLabel(row.length > 11 ? row[11] : null);
      tr.appendChild(tdSev);

      // Fix Version — last column
      const tdFix = document.createElement("td");
      tdFix.className = "bc-dft-td-fix";
      tdFix.textContent = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]) : "";
      tr.appendChild(tdFix);

      tbody.appendChild(tr);
    }
    bcDeliveredTable.appendChild(tbody);
  }

  /**
   * Render the "Returned Tickets" section.
   * Items qualify when:
   *   - They appear in the baseline with status TESTING or WAIT-INFO
   *   - They appear in the current file (matched by ID, col 0) with status TODO or INPROG
   * The table shows: ID, Description, Status, Severity, Fix Version.
   * Sort order: Fix Version (ascending) then Severity numeric (1 < 2 < 3 < 4).
   */
  function renderBcReturned(currentRows, baselineRows) {
    // Build a set of baseline IDs that were TESTING or WAIT-INFO
    const BASE_STATUSES = new Set(["TESTING", "WAIT-INFO"]);
    const baselineIds = new Set();
    for (const row of baselineRows) {
      const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      if (BASE_STATUSES.has(st) && row[0] != null) {
        baselineIds.add(String(row[0]).trim());
      }
    }

    // Find current rows whose ID was in baseline TESTING/WAIT-INFO and are now TODO or INPROG
    const CUR_STATUSES = new Set(["TODO", "INPROG"]);
    const SEV_ORDER = { "1": 1, "2": 2, "3": 3, "4": 4 };

    let returned = currentRows.filter(row => {
      const id = row[0] != null ? String(row[0]).trim() : "";
      const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      return id !== "" && baselineIds.has(id) && CUR_STATUSES.has(st);
    });

    // Sort: Fix Version asc (last col), then severity numeric asc
    returned.sort((a, b) => {
      const fa = a.length > 0 && a[a.length - 1] != null ? String(a[a.length - 1]) : "";
      const fb = b.length > 0 && b[b.length - 1] != null ? String(b[b.length - 1]) : "";
      const fCmp = fa.localeCompare(fb, undefined, { numeric: true, sensitivity: "base" });
      if (fCmp !== 0) return fCmp;
      const sa = SEV_ORDER[String(a.length > 11 && a[11] != null ? a[11] : "").trim()] ?? 99;
      const sb = SEV_ORDER[String(b.length > 11 && b[11] != null ? b[11] : "").trim()] ?? 99;
      return sa - sb;
    });

    // Count line
    bcReturnedCount.textContent = returned.length === 0
      ? "No items returned to development since baseline."
      : `${returned.length} item${returned.length === 1 ? "" : "s"} returned to development`;

    // Build table
    if (returned.length === 0) {
      bcReturnedTable.innerHTML = "";
      return;
    }

    bcReturnedTable.innerHTML =
      `<thead><tr>
        <th>ID</th>
        <th>Description</th>
        <th>Status</th>
        <th>Severity</th>
        <th>Fix Version</th>
      </tr></thead>`;

    const tbody = document.createElement("tbody");
    for (const row of returned) {
      const tr = document.createElement("tr");

      // ID — col 0
      const tdId = document.createElement("td");
      tdId.textContent = row[0] != null ? String(row[0]) : "";
      tr.appendChild(tdId);

      // Description — col 1
      const tdDesc = document.createElement("td");
      tdDesc.className = "bc-dft-td-desc";
      tdDesc.textContent = row[1] != null ? String(row[1]) : "";
      tr.appendChild(tdDesc);

      // Status — col 2
      const tdStatus = document.createElement("td");
      tdStatus.className = "bc-dft-td-status";
      tdStatus.textContent = row[2] != null ? String(row[2]).trim() : "";
      tr.appendChild(tdStatus);

      // Severity — col 11
      const tdSev = document.createElement("td");
      tdSev.className = "bc-dft-td-sev";
      tdSev.textContent = severityLabel(row.length > 11 ? row[11] : null);
      tr.appendChild(tdSev);

      // Fix Version — last column
      const tdFix = document.createElement("td");
      tdFix.className = "bc-dft-td-fix";
      tdFix.textContent = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]) : "";
      tr.appendChild(tdFix);

      tbody.appendChild(tr);
    }
    bcReturnedTable.appendChild(tbody);
  }

  /**
   * Render the "Resolved Tickets" section.
   * Items qualify when:
   *   - They appear in the baseline with status NOT DONE or CLOSED
   *   - They appear in the current file (matched by ID, col 0) with status DONE or CLOSED
   * The table shows: ID, Description, Status, Severity, Fix Version.
   * Sort order: Fix Version (ascending) then Severity numeric (1 < 2 < 3 < 4).
   */
  function renderBcResolved(currentRows, baselineRows) {
    // Build a set of baseline IDs that were NOT resolved
    const RESOLVED = new Set(["DONE", "CLOSED"]);
    const baselineIds = new Set();
    for (const row of baselineRows) {
      const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      if (!RESOLVED.has(st) && row[0] != null) {
        baselineIds.add(String(row[0]).trim());
      }
    }

    // Find current rows whose ID was unresolved in baseline and are now DONE or CLOSED
    const SEV_ORDER = { "1": 1, "2": 2, "3": 3, "4": 4 };

    let resolved = currentRows.filter(row => {
      const id = row[0] != null ? String(row[0]).trim() : "";
      const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      return id !== "" && baselineIds.has(id) && RESOLVED.has(st);
    });

    // Sort: Fix Version asc (last col), then severity numeric asc
    resolved.sort((a, b) => {
      const fa = a.length > 0 && a[a.length - 1] != null ? String(a[a.length - 1]) : "";
      const fb = b.length > 0 && b[b.length - 1] != null ? String(b[b.length - 1]) : "";
      const fCmp = fa.localeCompare(fb, undefined, { numeric: true, sensitivity: "base" });
      if (fCmp !== 0) return fCmp;
      const sa = SEV_ORDER[String(a.length > 11 && a[11] != null ? a[11] : "").trim()] ?? 99;
      const sb = SEV_ORDER[String(b.length > 11 && b[11] != null ? b[11] : "").trim()] ?? 99;
      return sa - sb;
    });

    // Count line
    bcResolvedCount.textContent = resolved.length === 0
      ? "No items resolved since baseline."
      : `${resolved.length} item${resolved.length === 1 ? "" : "s"} resolved since baseline`;

    // Build table
    if (resolved.length === 0) {
      bcResolvedTable.innerHTML = "";
      return;
    }

    bcResolvedTable.innerHTML =
      `<thead><tr>
        <th>ID</th>
        <th>Description</th>
        <th>Status</th>
        <th>Severity</th>
        <th>Fix Version</th>
      </tr></thead>`;

    const tbody = document.createElement("tbody");
    for (const row of resolved) {
      const tr = document.createElement("tr");

      // ID — col 0
      const tdId = document.createElement("td");
      tdId.textContent = row[0] != null ? String(row[0]) : "";
      tr.appendChild(tdId);

      // Description — col 1
      const tdDesc = document.createElement("td");
      tdDesc.className = "bc-dft-td-desc";
      tdDesc.textContent = row[1] != null ? String(row[1]) : "";
      tr.appendChild(tdDesc);

      // Status — col 2
      const tdStatus = document.createElement("td");
      tdStatus.className = "bc-dft-td-status";
      tdStatus.textContent = row[2] != null ? String(row[2]).trim() : "";
      tr.appendChild(tdStatus);

      // Severity — col 11
      const tdSev = document.createElement("td");
      tdSev.className = "bc-dft-td-sev";
      tdSev.textContent = severityLabel(row.length > 11 ? row[11] : null);
      tr.appendChild(tdSev);

      // Fix Version — last column
      const tdFix = document.createElement("td");
      tdFix.className = "bc-dft-td-fix";
      tdFix.textContent = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]) : "";
      tr.appendChild(tdFix);

      tbody.appendChild(tr);
    }
    bcResolvedTable.appendChild(tbody);
  }

  function parseBcFile(file) {
    bcErrorBox.hidden = true; bcErrorBox.textContent = "";
    bcReadyMsg.hidden = true;

    if (!file.name.endsWith(".xlsx")) {
      bcErrorBox.textContent = "Only .xlsx files are supported.";
      bcErrorBox.hidden = false;
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        bcRawRows  = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        bcFileName = file.name;
        applyBcHeaderSkip();
        bcFileNameEl.textContent = file.name;
        bcReadyMsg.textContent   = `✓ ${file.name} — ${bcRows.length} row${bcRows.length === 1 ? "" : "s"}`;
        bcReadyMsg.hidden        = false;
        // Re-init filters from the union of both files, then refresh KPIs
        bcSelectedGroups      = bcUnionValues(4);
        bcSelectedTypes       = bcUnionValues(5);
        bcSelectedSeverities  = bcUnionSeverities();
        bcSelectedFixVersions = bcUnionLastCol();
        buildBcFilters();
        refreshBcKpis();
      } catch (err) {
        bcErrorBox.textContent = "Failed to parse the file: " + err.message;
        bcErrorBox.hidden = false;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function applyBcHeaderSkip() {
    bcRows = (bcHeaderChk.checked && bcRawRows.length) ? bcRawRows.slice(1) : bcRawRows;
  }

  bcFileInput.addEventListener("change", () => {
    if (bcFileInput.files.length) parseBcFile(bcFileInput.files[0]);
  });
  bcDropZone.addEventListener("dragover",  e => { e.preventDefault(); bcDropZone.classList.add("drag-over"); });
  bcDropZone.addEventListener("dragleave", ()  => bcDropZone.classList.remove("drag-over"));
  bcDropZone.addEventListener("drop", e => {
    e.preventDefault();
    bcDropZone.classList.remove("drag-over");
    if (e.dataTransfer.files[0]) parseBcFile(e.dataTransfer.files[0]);
  });
  bcDropZone.addEventListener("click", () => bcFileInput.click());
  bcHeaderChk.addEventListener("change", () => {
    if (!bcRawRows.length) return;
    applyBcHeaderSkip();
    bcReadyMsg.textContent = `✓ ${bcFileName} — ${bcRows.length} row${bcRows.length === 1 ? "" : "s"}`;
    bcSelectedGroups      = bcUnionValues(4);
    bcSelectedTypes       = bcUnionValues(5);
    bcSelectedSeverities  = bcUnionSeverities();
    bcSelectedFixVersions = bcUnionLastCol();
    buildBcFilters();
    refreshBcKpis();
  });

  function initBaseline() {
    bcCurrentFileLbl.textContent = sharedFileName ? sharedFileName : "";
    bcErrorBox.hidden = true; bcErrorBox.textContent = "";
    // Refresh filters and KPIs in case the current file changed since last visit
    if (sharedRows.length || bcRows.length) {
      bcSelectedGroups      = bcUnionValues(4);
      bcSelectedTypes       = bcUnionValues(5);
      bcSelectedSeverities  = bcUnionSeverities();
      bcSelectedFixVersions = bcUnionLastCol();
      buildBcFilters();
    }
    refreshBcKpis();
  }

  // ── Home tile navigation ───────────────────────────────────────────────────
  document.getElementById("tileStats").addEventListener("click",  () => { showView("viewStats");  initStats();  });
  document.getElementById("tileUpdate").addEventListener("click", () => { showView("viewUpdate"); initUpdate(); });
  document.getElementById("tileBaseline").addEventListener("click", () => { showView("viewBaseline"); initBaseline(); });
  document.getElementById("tileStats").addEventListener("keydown",  e => { if (e.key === "Enter" || e.key === " ") { showView("viewStats");  initStats();  } });
  document.getElementById("tileUpdate").addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { showView("viewUpdate"); initUpdate(); } });
  document.getElementById("tileBaseline").addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { showView("viewBaseline"); initBaseline(); } });

  // Back buttons
  document.getElementById("backFromStats").addEventListener("click",    () => showView("viewHome"));
  document.getElementById("backFromUpdate").addEventListener("click",   () => showView("viewHome"));
  document.getElementById("backFromBaseline").addEventListener("click", () => showView("viewHome"));

  // ── Collapsible BC sections ─────────────────────────────────────────────────
  document.querySelectorAll(".bc-section-header").forEach(header => {
    header.addEventListener("click", () => {
      const section = header.closest(".bc-kpi-section");
      if (section) section.classList.toggle("bc-collapsed");
    });
  });

  // ── Update Tickets ─────────────────────────────────────────────────────────
  const updateDateControls = document.getElementById("updateDateControls");
  const updateFromDate     = document.getElementById("updateFromDate");
  const updateGroupSelect  = document.getElementById("updateGroupSelect");
  const updateTypeSelect   = document.getElementById("updateTypeSelect");
  const updateFetch        = document.getElementById("updateFetch");
  const errorBox           = document.getElementById("errorBox");
  const updateFileLabel    = document.getElementById("updateFileLabel");

  const resolvedSection  = document.getElementById("resolvedSection");
  const resolvedCount    = document.getElementById("resolvedCount");
  const resolvedTable    = document.getElementById("resolvedTable");
  const returnedSection  = document.getElementById("returnedSection");
  const returnedCount    = document.getElementById("returnedCount");
  const returnedTable    = document.getElementById("returnedTable");
  const newItemsSection  = document.getElementById("newItemsSection");
  const newItemsCount    = document.getElementById("newItemsCount");
  const newItemsTable    = document.getElementById("newItemsTable");

  function initUpdate() {
    if (!sharedRows.length) return;
    updateFileLabel.textContent = `File: ${sharedFileName}`;
    errorBox.hidden = true; errorBox.textContent = "";
    [resolvedSection, returnedSection, newItemsSection].forEach(s => s.hidden = true);
    updateDateControls.hidden = true;

    // Default status date to today
    if (!updateFromDate.value) {
      const today = new Date();
      updateFromDate.value =
        `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
    }

    // Populate Group select from unique col-5 values (index 4)
    const groups = [...new Set(
      sharedRows.map(r => r[4] != null ? String(r[4]).trim() : null).filter(v => v)
    )].sort();

    updateGroupSelect.innerHTML = `<option value="">All groups</option>`;
    for (const g of groups) {
      const opt = document.createElement("option");
      opt.value = g; opt.textContent = g;
      updateGroupSelect.appendChild(opt);
    }

    // Populate Type select from unique col-6 values (index 5)
    const types = [...new Set(
      sharedRows.map(r => r[5] != null ? String(r[5]).trim() : null).filter(v => v)
    )].sort();

    updateTypeSelect.innerHTML = "";
    for (const t of types) {
      const opt = document.createElement("option");
      opt.value = t; opt.textContent = t;
      updateTypeSelect.appendChild(opt);
    }
    updateDateControls.hidden = false;
  }

  updateFetch.addEventListener("click", () => {
    errorBox.hidden = true; errorBox.textContent = "";
    [resolvedSection, returnedSection, newItemsSection].forEach(s => s.hidden = true);

    if (!updateFromDate.value) {
      errorBox.textContent = "Please select a status date.";
      errorBox.hidden = false;
      return;
    }

    const fromDate      = new Date(updateFromDate.value + "T00:00:00Z");
    const selectedGroup = updateGroupSelect.value;   // "" = All groups
    const selectedType  = updateTypeSelect.value;
    const DONE_SET      = new Set(["DONE", "CLOSED"]);

    // Shared group predicate — passes when no group is selected or group matches
    const matchesGroup = row => {
      if (!selectedGroup) return true;
      const g = row[4] != null ? String(row[4]).trim() : "";
      return g === selectedGroup;
    };

    // ── Section 1: Resolved — status DONE/CLOSED, col 4 (index 3) >= fromDate ─
    const resolved = sharedRows.filter(row => {
      const status = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      const type   = row[5] != null ? String(row[5]).trim() : "";
      if (!DONE_SET.has(status)) return false;
      if (type !== selectedType)  return false;
      if (!matchesGroup(row))     return false;
      const d = parseDateVal(row[3]);
      return d !== null && d >= fromDate;
    });

    // ── Section 2: Returned — status TODO, col 4 >= fromDate, col 11 < fromDate
    const returned = sharedRows.filter(row => {
      const status = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      const type   = row[5] != null ? String(row[5]).trim() : "";
      if (status !== "TODO")     return false;
      if (type !== selectedType) return false;
      if (!matchesGroup(row))    return false;
      const statusDate  = parseDateVal(row[3]);
      const createdDate = parseDateVal(row[10]);
      return statusDate !== null && statusDate >= fromDate &&
             createdDate !== null && createdDate < fromDate;
    });

    // ── Section 3: New Items — status TODO, col 11 (index 10) >= fromDate ──────
    const newItems = sharedRows.filter(row => {
      const status = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      const type   = row[5] != null ? String(row[5]).trim() : "";
      if (status !== "TODO")     return false;
      if (type !== selectedType) return false;
      if (!matchesGroup(row))    return false;
      const createdDate = parseDateVal(row[10]);
      return createdDate !== null && createdDate >= fromDate;
    });

    renderSection(resolvedSection,  resolvedCount,  resolvedTable,  resolved,  "resolved item");
    renderSection(returnedSection,  returnedCount,  returnedTable,  returned,  "returned item");
    renderSection(newItemsSection,  newItemsCount,  newItemsTable,  newItems,  "new item");
  });

  /** Parse an Excel date serial or string into a UTC midnight Date, or null */
  function parseDateVal(val) {
    if (val == null) return null;
    if (val instanceof Date) {
      return isNaN(val) ? null : new Date(Date.UTC(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate()));
    }
    if (typeof val === "number") {
      const ms = (val - 25569) * 86400000;
      const d  = new Date(ms);
      return isNaN(d) ? null : new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    }
    const s = String(val).trim();
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d) ? null : new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }

  /** Format a date cell value as DD/MM/YYYY, or blank if unparseable */
  function fmtDateCell(val) {
    const d = parseDateVal(val);
    if (!d) return "";
    return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;
  }

  /** Render a result section as a table: Ticket No · Summary · Status Date · Severity · Fix Version */
  function renderSection(section, countEl, tableEl, rows, label) {
    tableEl.innerHTML = "";
    countEl.textContent = rows.length === 0
      ? `No ${label}s found.`
      : `${rows.length} ${label}${rows.length === 1 ? "" : "s"}`;

    if (rows.length === 0) { section.hidden = false; return; }

    // Header
    tableEl.innerHTML =
      `<thead><tr>
        <th>Ticket Number</th>
        <th>Summary</th>
        <th>Status Date</th>
        <th>Severity</th>
        <th>Fix Version</th>
      </tr></thead>`;

    const tbody = document.createElement("tbody");
    for (const row of rows) {
      const tr = document.createElement("tr");

      // Ticket Number — col 1 (index 0)
      const tdId = document.createElement("td");
      tdId.textContent = row[0] != null ? String(row[0]) : "";
      tr.appendChild(tdId);

      // Summary — col 2 (index 1)
      const tdSum = document.createElement("td");
      tdSum.className = "update-td-summary";
      tdSum.textContent = row[1] != null ? String(row[1]) : "";
      tr.appendChild(tdSum);

      // Status Date — col 4 (index 3)
      const tdDate = document.createElement("td");
      tdDate.className = "update-td-date";
      tdDate.textContent = fmtDateCell(row[3]);
      tr.appendChild(tdDate);

      // Severity — col 12 (index 11), mapped via severityLabel()
      const tdSev = document.createElement("td");
      tdSev.className = "update-td-sev";
      tdSev.textContent = severityLabel(row.length > 11 ? row[11] : null);
      tr.appendChild(tdSev);

      // Fix Version — last column
      const tdFix = document.createElement("td");
      tdFix.className = "update-td-fix";
      tdFix.textContent = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]) : "";
      tr.appendChild(tdFix);

      tbody.appendChild(tr);
    }
    tableEl.appendChild(tbody);
    section.hidden = false;
  }

  // ── Statistics wiring ──────────────────────────────────────────────────────
  const statsErrorBox  = document.getElementById("statsErrorBox");
  const statsFilters   = document.getElementById("statsFilters");
  const statsFileLabel = document.getElementById("statsFileLabel");
  const chartsArea     = document.getElementById("chartsArea");   // kept for compat
  const dashKpiRow     = document.getElementById("dashKpiRow");
  const dashGrid       = document.getElementById("dashGrid");
  const dashFull       = document.getElementById("dashFull");
  const reopenSection  = document.getElementById("reopenSection");
  const reopenCount    = document.getElementById("reopenCount");
  const reopenTable    = document.getElementById("reopenTable");

  // ── Sidebar collapse toggle ────────────────────────────────────────────────
  const dashSidebar       = document.getElementById("dashSidebar");
  const dashSidebarToggle = document.getElementById("dashSidebarToggle");
  dashSidebarToggle.addEventListener("click", () => {
    const collapsed = dashSidebar.classList.toggle("collapsed");
    dashSidebarToggle.textContent = (collapsed ? "\u25B6" : "\u25BC") + " Filters";
    dashSidebarToggle.setAttribute("aria-expanded", String(!collapsed));
  });

  // ── Baseline sidebar collapse toggle ──────────────────────────────────────
  const bcSidebar       = document.getElementById("bcSidebar");
  const bcSidebarToggle = document.getElementById("bcSidebarToggle");
  bcSidebarToggle.addEventListener("click", () => {
    const collapsed = bcSidebar.classList.toggle("collapsed");
    bcSidebarToggle.textContent = (collapsed ? "\u25B6" : "\u25BC") + " Filters";
    bcSidebarToggle.setAttribute("aria-expanded", String(!collapsed));
  });

  // Raw data rows (set once per file load; never mutated)
  let allRows = [];
  // Active filter sets — initialised to all values on file load
  let selectedGroups      = new Set(); // col 5  (index 4)
  let selectedTypes       = new Set(); // col 6  (index 5)
  let selectedSeverities  = new Set(); // col 12 (index 11) — mapped to label
  let selectedFixVersions = new Set(); // last column

  const SEVERITY_LABEL = { "1": "Critical", "2": "Major", "3": "Medium", "4": "Low" };

  /** Normalise a raw col-12 value to its display label */
  function severityLabel(raw) {
    const s = raw != null ? String(raw).trim() : "";
    return SEVERITY_LABEL[s] ?? (s || "(blank)");
  }

  /** Unique severity labels present in the data */
  function uniqueSeverityLabels(rows) {
    const s = new Set();
    for (const row of rows) {
      if (row.length > 11) s.add(severityLabel(row[11]));
    }
    return s;
  }

  function initStats() {
    if (!sharedRows.length) return;
    statsFileLabel.textContent = `File: ${sharedFileName}`;
    statsErrorBox.hidden = true; statsErrorBox.textContent = "";
    statsFilters.hidden = true; statsFilters.innerHTML = "";
    chartsArea.hidden = true;   chartsArea.innerHTML = "";
    dashKpiRow.innerHTML = ""; dashKpiRow.hidden = true;
    dashFull.hidden = true;
    reopenSection.hidden = true; reopenTable.innerHTML = "";
    [bugsCreationSection, enhCreationSection, opCreationSection].forEach(s => {
      s.hidden = true; s.innerHTML = "";
    });
    // Clear chart nodes from dashGrid but keep the two vertical section nodes
    const KEEP = new Set(["col16Section", "col16EnhSection"]);
    Array.from(dashGrid.children).forEach(el => {
      if (!KEEP.has(el.id)) el.remove();
    });
    dashGrid.hidden = true;
    // Reset both vertical sections so a new file re-initialises selections
    for (const sec of Object.values(verticalSections)) {
      sec.selected    = new Set();
      sec.initialised = false;
      sec.activeKey   = null;
    }

    allRows = sharedRows;

    // Initialise selections to "all values selected"
    selectedGroups      = uniqueValues(allRows, 4);
    selectedTypes       = uniqueValues(allRows, 5);
    selectedSeverities  = uniqueSeverityLabels(allRows);
    selectedFixVersions = uniqueLastCol(allRows);

    buildFilters();
    refreshCharts();
    showAgingSection();
  }

  // ── Filter helpers ─────────────────────────────────────────────────────────

  /** Returns a Set of trimmed, non-null unique string values in a column */
  function uniqueValues(rows, colIndex) {
    const s = new Set();
    for (const row of rows) {
      if (row.length > colIndex && row[colIndex] != null) {
        s.add(String(row[colIndex]).trim());
      }
    }
    return s;
  }

  /** Returns a Set of unique values from the last cell of each row */
  function uniqueLastCol(rows) {
    const s = new Set();
    for (const row of rows) {
      if (row.length > 0 && row[row.length - 1] != null) {
        s.add(String(row[row.length - 1]).trim());
      }
    }
    return s;
  }

  /** Returns rows that pass all active filter selections */
  function filteredRows() {
    return allRows.filter(row => {
      const group      = row.length > 4 && row[4] != null ? String(row[4]).trim() : "(blank)";
      const type       = row.length > 5 && row[5] != null ? String(row[5]).trim() : "(blank)";
      const severity   = severityLabel(row.length > 11 ? row[11] : null);
      const fixVersion = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]).trim() : "(blank)";
      return selectedGroups.has(group) && selectedTypes.has(type)
          && selectedSeverities.has(severity) && selectedFixVersions.has(fixVersion);
    });
  }

  /** Builds all chip-filter groups and attaches change listeners */
  function buildFilters() {
    statsFilters.innerHTML = "";

    statsFilters.appendChild(chipGroup(
      "Group", selectedGroups,
      [...uniqueValues(allRows, 4)].sort(),
      val => {
        selectedGroups.has(val) ? selectedGroups.delete(val) : selectedGroups.add(val);
        refreshCharts();
      },
      () => {
        const all = uniqueValues(allRows, 4);
        if (selectedGroups.size === all.size) { selectedGroups.clear(); }
        else { selectedGroups = new Set(all); }
        buildFilters();
        refreshCharts();
      }
    ));

    statsFilters.appendChild(chipGroup(
      "Ticket Type", selectedTypes,
      [...uniqueValues(allRows, 5)].sort(),
      val => {
        selectedTypes.has(val) ? selectedTypes.delete(val) : selectedTypes.add(val);
        refreshCharts();
      },
      () => {
        const all = uniqueValues(allRows, 5);
        if (selectedTypes.size === all.size) { selectedTypes.clear(); }
        else { selectedTypes = new Set(all); }
        buildFilters();
        refreshCharts();
      }
    ));

    const SEVERITY_ORDER = ["Critical", "Major", "Medium", "Low"];
    const severityValues = SEVERITY_ORDER.filter(l => uniqueSeverityLabels(allRows).has(l));
    // append any unexpected labels not in the known order
    for (const l of uniqueSeverityLabels(allRows)) {
      if (!SEVERITY_ORDER.includes(l)) severityValues.push(l);
    }

    statsFilters.appendChild(chipGroup(
      "Severity", selectedSeverities,
      severityValues,
      val => {
        selectedSeverities.has(val) ? selectedSeverities.delete(val) : selectedSeverities.add(val);
        refreshCharts();
      },
      () => {
        const all = uniqueSeverityLabels(allRows);
        if (selectedSeverities.size === all.size) { selectedSeverities.clear(); }
        else { selectedSeverities = new Set(all); }
        buildFilters();
        refreshCharts();
      }
    ));

    statsFilters.appendChild(chipGroup(
      "Fix Version", selectedFixVersions,
      [...uniqueLastCol(allRows)].sort(),
      val => {
        selectedFixVersions.has(val) ? selectedFixVersions.delete(val) : selectedFixVersions.add(val);
        refreshCharts();
      },
      () => {
        const all = uniqueLastCol(allRows);
        if (selectedFixVersions.size === all.size) { selectedFixVersions.clear(); }
        else { selectedFixVersions = new Set(all); }
        buildFilters();
        refreshCharts();
      }
    ));

    statsFilters.hidden = false;
  }

  /** Creates one filter-group DOM node */
  function chipGroup(label, selectedSet, values, onToggle, onToggleAll) {
    const group = document.createElement("div");
    group.className = "filter-group";

    const allSelected = values.every(v => selectedSet.has(v));

    group.innerHTML = `
      <div class="filter-group-header">
        <span class="filter-group-label">${label}</span>
        <button class="filter-toggle-all">${allSelected ? "Deselect all" : "Select all"}</button>
      </div>
      <div class="filter-chips"></div>`;

    group.querySelector(".filter-toggle-all").addEventListener("click", onToggleAll);

    const chipsEl = group.querySelector(".filter-chips");
    for (const val of values) {
      const chip = document.createElement("label");
      chip.className = "filter-chip" + (selectedSet.has(val) ? " selected" : "");
      chip.title = val;
      chip.innerHTML = `<input type="checkbox" ${selectedSet.has(val) ? "checked" : ""}><span>${val}</span>`;
      chip.querySelector("input").addEventListener("change", () => {
        chip.classList.toggle("selected");
        onToggle(val);
      });
      chipsEl.appendChild(chip);
    }

    return group;
  }

  function refreshCharts() {
    chartsArea.innerHTML = "";
    // Clear chart panels from grid but keep both vertical section nodes
    const KEEP = new Set(["col16Section", "col16EnhSection"]);
    Array.from(dashGrid.children).forEach(el => {
      if (!KEEP.has(el.id)) el.remove();
    });
    dashKpiRow.innerHTML = "";
    renderCharts(filteredRows());
    buildCol16Section(filteredRows());
    buildReopenSection();
    // keep aging chart in sync when filters change
    refreshAging();
  }

  function buildReopenSection() {
    reopenTable.innerHTML = "";
    reopenSection.hidden = true;

    // Today at UTC midnight — independent of sidebar filters
    const now = new Date();
    const today = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));

    const rows = allRows.filter(row => {
      const status = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      if (status !== "TODO") return false;
      const statusDate  = parseDateVal(row[3]);
      const createdDate = parseDateVal(row[10]);
      return statusDate  !== null && statusDate  >= today &&
             createdDate !== null && createdDate <  today;
    });

    reopenCount.textContent = rows.length === 0
      ? "No reopen items found."
      : `${rows.length} item${rows.length === 1 ? "" : "s"}`;

    if (rows.length === 0) { reopenSection.hidden = false; return; }

    reopenTable.innerHTML =
      `<thead><tr>
        <th>Ticket Number</th>
        <th>Summary</th>
        <th>Status Date</th>
        <th>Severity</th>
        <th>Fix Version</th>
      </tr></thead>`;

    const tbody = document.createElement("tbody");
    for (const row of rows) {
      const tr = document.createElement("tr");

      const tdId = document.createElement("td");
      tdId.textContent = row[0] != null ? String(row[0]) : "";
      tr.appendChild(tdId);

      const tdSum = document.createElement("td");
      tdSum.className = "update-td-summary";
      tdSum.textContent = row[1] != null ? String(row[1]) : "";
      tr.appendChild(tdSum);

      const tdDate = document.createElement("td");
      tdDate.className = "update-td-date";
      tdDate.textContent = fmtDateCell(row[3]);
      tr.appendChild(tdDate);

      const tdSev = document.createElement("td");
      tdSev.className = "update-td-sev";
      tdSev.textContent = severityLabel(row.length > 11 ? row[11] : null);
      tr.appendChild(tdSev);

      const tdFix = document.createElement("td");
      tdFix.className = "update-td-fix";
      tdFix.textContent = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]) : "";
      tr.appendChild(tdFix);

      tbody.appendChild(tr);
    }
    reopenTable.appendChild(tbody);
    reopenSection.hidden = false;
  }

  // ── Aging ──────────────────────────────────────────────────────────────────
  const agingSection    = document.getElementById("agingSection");
  const agingFrom       = document.getElementById("agingFrom");
  const agingTo         = document.getElementById("agingTo");
  const agingFetch      = document.getElementById("agingFetch");
  const agingError      = document.getElementById("agingError");
  const agingChartWrap      = document.getElementById("agingChartWrap");
  const bugsCreationSection = document.getElementById("bugsCreationSection");
  const enhCreationSection  = document.getElementById("enhCreationSection");
  const opCreationSection   = document.getElementById("opCreationSection");

  /** Show the aging section, pre-fill default dates, and auto-fetch */
  function showAgingSection() {
    dashFull.hidden = false;
    agingSection.hidden = false;

    // Default From: 2026-03-01  |  Default To: today
    if (!agingFrom.value) agingFrom.value = "2026-03-01";
    if (!agingTo.value) {
      const today = new Date();
      agingTo.value = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
    }

    runAging();
  }

  /** Parse a cell value from column 4 (index 3) into a JS Date (midnight UTC).
   *  Accepts JS Date serial numbers (Excel) or ISO / common string formats. */
  function parseDate(val) {
    if (val == null) return null;
    if (val instanceof Date) {
      return isNaN(val) ? null : new Date(Date.UTC(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate()));
    }
    // Excel numeric serial date
    if (typeof val === "number") {
      // Excel epoch is 1900-01-01, but has a leap-year bug so offset is 25569 days to Unix epoch
      const ms = (val - 25569) * 86400000;
      const d = new Date(ms);
      return isNaN(d) ? null : new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    }
    const s = String(val).trim();
    if (!s) return null;
    // Try ISO and common formats
    const d = new Date(s);
    if (!isNaN(d)) return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    return null;
  }

  function refreshAging() {
    // Only re-run if dates are already set and chart was previously shown
    if (agingChartWrap.hidden) return;
    runAging();
  }

  agingFetch.addEventListener("click", runAging);

  function runAging() {
    agingError.hidden = true;
    agingChartWrap.hidden = true;
    agingChartWrap.innerHTML = "";
    [bugsCreationSection, enhCreationSection, opCreationSection].forEach(s => {
      s.hidden = true; s.innerHTML = "";
    });

    const fromVal = agingFrom.value;
    const toVal   = agingTo.value;

    if (!fromVal || !toVal) {
      agingError.textContent = "Please select both From and To dates.";
      agingError.hidden = false;
      return;
    }

    const fromDate = new Date(fromVal + "T00:00:00Z");
    const toDate   = new Date(toVal   + "T00:00:00Z");

    if (fromDate >= toDate) {
      agingError.textContent = "From date must be before To date.";
      agingError.hidden = false;
      return;
    }

    // Build calendar-week checkpoints (each point = Sunday end-of-week)
    // Advance fromDate to its Monday, then each checkpoint is the following Sunday
    const MS_DAY = 86400000;

    /** Return the Monday on or before a given UTC date */
    function weekMonday(d) {
      const dow = d.getUTCDay(); // 0=Sun … 6=Sat
      const diff = (dow === 0) ? -6 : 1 - dow; // days back to Monday
      return new Date(d.getTime() + diff * MS_DAY);
    }

    /** Return the Sunday on or after a given UTC date */
    function weekSunday(d) {
      const dow = d.getUTCDay();
      const diff = (dow === 0) ? 0 : 7 - dow;
      return new Date(d.getTime() + diff * MS_DAY);
    }

    /** ISO week number (Mon–Sun, week 1 = first week with a Thursday) */
    function isoWeek(d) {
      const thu = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      const dow = d.getUTCDay() || 7; // convert Sun=0 → 7
      thu.setUTCDate(thu.getUTCDate() + 4 - dow); // nearest Thursday
      const yearStart = new Date(Date.UTC(thu.getUTCFullYear(), 0, 1));
      return Math.ceil(((thu - yearStart) / MS_DAY + 1) / 7);
    }

    function fmtDate(d) {
      return `${String(d.getUTCMonth() + 1).padStart(2, "0")}/${String(d.getUTCDate()).padStart(2, "0")}`;
    }

    const checkpoints = [];
    const labels      = [];

    let cur = weekSunday(fromDate); // first Sunday on-or-after fromDate
    while (cur <= toDate) {
      checkpoints.push(new Date(cur));
      const mon = weekMonday(cur);
      labels.push(`W${isoWeek(cur)} ${fmtDate(mon)}–${fmtDate(cur)}`);
      cur = new Date(cur.getTime() + 7 * MS_DAY);
    }
    // Always include the to-date's week if not already covered
    if (checkpoints.length === 0 || checkpoints[checkpoints.length - 1] < toDate) {
      const lastSun = weekSunday(toDate);
      checkpoints.push(lastSun);
      const mon = weekMonday(lastSun);
      labels.push(`W${isoWeek(lastSun)} ${fmtDate(mon)}–${fmtDate(lastSun)}`);
    }

    // For each checkpoint count tickets whose col-4 date <= checkpoint (end of week)
    const rows = filteredRows();
    const TODO_SET    = new Set(["TODO"]);
    const INPROG_SET  = new Set(["INPROG"]);
    const TESTING_SET = new Set(["TESTING"]);
    const DONE_SET    = new Set(["DONE", "CLOSED"]);

    const todoSeries    = [];
    const inprogSeries  = [];
    const testingSeries = [];
    const doneSeries    = [];

    for (const cp of checkpoints) {
      let todo = 0, inprog = 0, testing = 0, done = 0;
      for (const row of rows) {
        const status = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
        const d = parseDate(row[3]);
        if (!d || d > cp) continue;
        if (TODO_SET.has(status))        todo++;
        else if (INPROG_SET.has(status)) inprog++;
        else if (TESTING_SET.has(status)) testing++;
        else if (DONE_SET.has(status))   done++;
      }
      todoSeries.push(todo);
      inprogSeries.push(inprog);
      testingSeries.push(testing);
      doneSeries.push(done);
    }

    agingChartWrap.appendChild(buildLineChart(checkpoints, labels, todoSeries, inprogSeries, testingSeries, doneSeries));
    agingChartWrap.hidden = false;

    // ── Type-specific creation charts (independent of sidebar filters) ────────
    // Uses allRows so sidebar filters do not affect these sections.
    const totalDays = Math.max(1, Math.round((toDate.getTime() - fromDate.getTime()) / MS_DAY) + 1);

    function buildTypeCreationSeries(typeValue) {
      return checkpoints.map(cp => {
        const mon = weekMonday(cp);
        let count = 0;
        for (const row of allRows) {
          const t = row[5] != null ? String(row[5]).trim() : "";
          if (t !== typeValue) continue;
          const d = parseDate(row[10]);
          if (d && d >= mon && d <= cp) count++;
        }
        return count;
      });
    }

    const configs = [
      { el: bugsCreationSection, type: "Defect (Bug)",      title: "Bugs Created per Week",          color: "#ef4444", h: 198 },
      { el: enhCreationSection,  type: "Enhancement (SR)",  title: "Enhancements Created per Week",   color: "#7c5cd8", h: 198 },
      { el: opCreationSection,   type: "Operational",       title: "Operational Created per Week",    color: "#f59e0b" },
    ];

    for (const cfg of configs) {
      const series = buildTypeCreationSeries(cfg.type);
      const total  = series.reduce((a, b) => a + b, 0);
      const avg    = total / totalDays;
      cfg.el.appendChild(buildCreationChart(checkpoints, labels, series, avg, fromDate, toDate, cfg.title, cfg.color, cfg.h));
      cfg.el.hidden = false;
    }
  }

  function buildLineChart(checkpoints, labels, todoSeries, inprogSeries, testingSeries, doneSeries) {
    const NS = "http://www.w3.org/2000/svg";
    const W = 620, H = 216;
    const PAD = { top: 16, right: 20, bottom: 52, left: 44 };
    const chartW = W - PAD.left - PAD.right;
    const chartH = H - PAD.top  - PAD.bottom;
    const n = checkpoints.length;

    const allVals = [...todoSeries, ...inprogSeries, ...testingSeries, ...doneSeries];
    const maxVal  = Math.max(...allVals, 1);

    const xOf = i => PAD.left + (n === 1 ? chartW / 2 : (i / (n - 1)) * chartW);
    const yOf = v => PAD.top  + chartH - (v / maxVal) * chartH;

    const fmt = (_, i) => labels[i] ?? "";

    const svg = document.createElementNS(NS, "svg");
    svg.setAttribute("viewBox", `0 0 ${W} ${H}`);
    svg.setAttribute("width", "100%");
    svg.removeAttribute("height");
    svg.classList.add("aging-svg");

    // Grid lines + Y-axis labels
    const steps = 5;
    for (let s = 0; s <= steps; s++) {
      const v = Math.round((maxVal / steps) * s);
      const y = yOf(v);
      const line = document.createElementNS(NS, "line");
      line.setAttribute("x1", PAD.left); line.setAttribute("x2", PAD.left + chartW);
      line.setAttribute("y1", y);        line.setAttribute("y2", y);
      line.classList.add("aging-grid-line");
      svg.appendChild(line);

      const lbl = document.createElementNS(NS, "text");
      lbl.setAttribute("x", PAD.left - 4);
      lbl.setAttribute("y", y + 3);
      lbl.setAttribute("text-anchor", "end");
      lbl.classList.add("aging-axis-label");
      lbl.textContent = v;
      svg.appendChild(lbl);
    }

    // X-axis base line
    const xAxis = document.createElementNS(NS, "line");
    xAxis.setAttribute("x1", PAD.left); xAxis.setAttribute("x2", PAD.left + chartW);
    xAxis.setAttribute("y1", PAD.top + chartH); xAxis.setAttribute("y2", PAD.top + chartH);
    xAxis.classList.add("aging-axis-line");
    svg.appendChild(xAxis);

    // X-axis labels (max 12, rotated when dense)
    const labelStep = Math.ceil(n / 12);
    checkpoints.forEach((cp, i) => {
      if (i % labelStep !== 0 && i !== n - 1) return;
      const x = xOf(i);
      const lbl = document.createElementNS(NS, "text");
      lbl.setAttribute("x", x);
      lbl.setAttribute("y", PAD.top + chartH + 14);
      lbl.setAttribute("text-anchor", "middle");
      lbl.classList.add("aging-axis-label");
      lbl.textContent = fmt(cp, i);
      if (n > 8) lbl.setAttribute("transform", `rotate(-35, ${x}, ${PAD.top + chartH + 14})`);
      svg.appendChild(lbl);
    });

    // Draw polyline + dots (no individual <title> — tooltip handled by overlay)
    function drawSeries(series, dotClass, lineClass) {
      if (n === 0) return;
      const points = series.map((v, i) => `${xOf(i).toFixed(1)},${yOf(v).toFixed(1)}`).join(" ");
      const poly = document.createElementNS(NS, "polyline");
      poly.setAttribute("points", points);
      poly.classList.add(lineClass);
      svg.appendChild(poly);

      series.forEach((v, i) => {
        const dot = document.createElementNS(NS, "circle");
        dot.setAttribute("cx", xOf(i));
        dot.setAttribute("cy", yOf(v));
        dot.setAttribute("r", 3);
        dot.classList.add(dotClass);
        svg.appendChild(dot);
      });
    }

    drawSeries(inprogSeries,  "aging-dot-inprog",   "aging-series-inprog");
    drawSeries(testingSeries, "aging-dot-testing",  "aging-series-testing");
    drawSeries(doneSeries,    "aging-dot-done",     "aging-series-done");
    drawSeries(todoSeries,    "aging-dot-todo",     "aging-series-todo");

    // ── Hover tooltip overlay ────────────────────────────────────────────────
    // Invisible full-chart hit area that tracks the nearest checkpoint
    const overlay = document.createElementNS(NS, "rect");
    overlay.setAttribute("x", PAD.left);
    overlay.setAttribute("y", PAD.top);
    overlay.setAttribute("width",  chartW);
    overlay.setAttribute("height", chartH);
    overlay.setAttribute("fill", "transparent");
    svg.appendChild(overlay);

    // Vertical crosshair line (hidden by default)
    const crosshair = document.createElementNS(NS, "line");
    crosshair.setAttribute("y1", PAD.top);
    crosshair.setAttribute("y2", PAD.top + chartH);
    crosshair.setAttribute("stroke", "#c0c8d0");
    crosshair.setAttribute("stroke-width", "1");
    crosshair.setAttribute("stroke-dasharray", "3 2");
    crosshair.setAttribute("display", "none");
    svg.appendChild(crosshair);

    // Wrap in a position:relative container so the tooltip can be absolute
    const wrap = document.createElement("div");
    wrap.style.position = "relative";
    wrap.appendChild(svg);

    const tooltip = document.createElement("div");
    tooltip.className = "aging-tooltip";
    wrap.appendChild(tooltip);

    // Mouse move on the SVG overlay
    overlay.addEventListener("mousemove", e => {
      const rect = svg.getBoundingClientRect();
      const mouseX = e.clientX - rect.left - PAD.left;
      const frac   = Math.max(0, Math.min(1, mouseX / chartW));
      const idx    = Math.round(frac * (n - 1));

      const cx = xOf(idx);
      crosshair.setAttribute("x1", cx);
      crosshair.setAttribute("x2", cx);
      crosshair.setAttribute("display", "");

      tooltip.innerHTML =
        `<strong>${labels[idx]}</strong><br>` +
        `<span style="color:#3b82d4">■</span> TODO: ${todoSeries[idx]}<br>` +
        `<span style="color:#f59e0b">■</span> In Progress: ${inprogSeries[idx]}<br>` +
        `<span style="color:#7c5cd8">■</span> Testing: ${testingSeries[idx]}<br>` +
        `<span style="color:#22c55e">■</span> DONE: ${doneSeries[idx]}`;

      // Position tooltip: prefer right of cursor, flip left near the edge
      const svgLeft  = rect.left - wrap.getBoundingClientRect().left;
      const tipX     = svgLeft + cx + 10;
      const tipRight = tipX + 140; // estimated tooltip width
      tooltip.style.left = (tipRight > wrap.offsetWidth ? svgLeft + cx - 145 : tipX) + "px";
      tooltip.style.top  = (PAD.top + 8) + "px";
      tooltip.style.display = "block";
    });

    overlay.addEventListener("mouseleave", () => {
      crosshair.setAttribute("display", "none");
      tooltip.style.display = "none";
    });

    // Legend
    const legend = document.createElement("div");
    legend.className = "aging-legend";
    legend.innerHTML = `
      <div class="aging-legend-item">
        <span class="aging-legend-swatch" style="background:#3b82d4"></span> TODO
      </div>
      <div class="aging-legend-item">
        <span class="aging-legend-swatch" style="background:#f59e0b"></span> In Progress
      </div>
      <div class="aging-legend-item">
        <span class="aging-legend-swatch" style="background:#7c5cd8"></span> Testing
      </div>
      <div class="aging-legend-item">
        <span class="aging-legend-swatch" style="background:#22c55e"></span> DONE / CLOSED
      </div>`;
    wrap.appendChild(legend);
    return wrap;
  }

  /**
   * Build the "Items Created per Week" bar chart with a stat header.
   * @param {Date[]}   checkpoints  — Sunday end-of-week dates (one per bar)
   * @param {string[]} labels       — week labels matching checkpoints
   * @param {number[]} createdSeries — items created within each week
   * @param {number}   avgPerDay    — average items created per calendar day across the range
   * @param {Date}     fromDate     — range start
   * @param {Date}     toDate       — range end
   */
  function buildCreationChart(checkpoints, labels, createdSeries, avgPerDay, fromDate, toDate, title, barColor, chartHeight) {
    const NS = "http://www.w3.org/2000/svg";
    const W = 620, H = chartHeight ?? 220;
    const PAD = { top: 20, right: 20, bottom: 52, left: 44 };
    const chartW = W - PAD.left - PAD.right;
    const chartH = H - PAD.top  - PAD.bottom;
    const n = checkpoints.length;
    const BAR_COLOR = barColor ?? "#3b82d4";

    const section = document.createElement("div");

    // ── Trend: compare avg of first half vs second half of weeks ─────────────
    const mid  = Math.floor(n / 2);
    const firstHalf  = createdSeries.slice(0, mid);
    const secondHalf = createdSeries.slice(n - mid);
    const avgFirst  = firstHalf.length  ? firstHalf.reduce((a, b)  => a + b, 0) / firstHalf.length  : 0;
    const avgSecond = secondHalf.length ? secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length : 0;
    const delta = avgSecond - avgFirst;
    const THRESHOLD = 0.1; // treat < 0.1 items/week difference as flat

    let trendClass, trendIcon, trendText;
    if (Math.abs(delta) < THRESHOLD) {
      trendClass = "flat"; trendIcon = "→"; trendText = "Stable";
    } else if (delta > 0) {
      trendClass = "up";   trendIcon = "▲"; trendText = `+${delta.toFixed(1)}/wk vs first half`;
    } else {
      trendClass = "down"; trendIcon = "▼"; trendText = `${delta.toFixed(1)}/wk vs first half`;
    }

    // ── Stat bar ──────────────────────────────────────────────────────────────
    const statBar = document.createElement("div");
    statBar.className = "creation-stat-bar";
    statBar.innerHTML =
      `<div class="creation-stat-item">
         <span class="creation-stat-value">${createdSeries.reduce((a, b) => a + b, 0)}</span>
         <span class="creation-stat-label">Total Created</span>
       </div>
       <div class="creation-stat-item">
         <span class="creation-stat-value">${avgPerDay.toFixed(2)}</span>
         <span class="creation-stat-label">Avg per Day</span>
       </div>
       <div class="creation-stat-item">
         <span class="creation-stat-value">${n > 0 ? (createdSeries.reduce((a, b) => a + b, 0) / n).toFixed(1) : "—"}</span>
         <span class="creation-stat-label">Avg per Week</span>
       </div>
       <span class="creation-trend ${trendClass}">${trendIcon} ${trendText}</span>`;

    // ── SVG bar chart ─────────────────────────────────────────────────────────
    const maxVal = Math.max(...createdSeries, 1);
    const xOf    = i => PAD.left + (n <= 1 ? chartW / 2 : (i / (n - 1)) * chartW);
    const yOf    = v => PAD.top + chartH - (v / maxVal) * chartH;
    const barW   = n <= 1 ? 40 : Math.max(4, Math.min(36, (chartW / n) * 0.65));

    const svg = document.createElementNS(NS, "svg");
    svg.setAttribute("viewBox", `0 0 ${W} ${H}`);
    svg.setAttribute("width", "100%");
    svg.removeAttribute("height");
    svg.classList.add("aging-svg");

    // Grid lines + Y-axis labels
    const steps = 4;
    for (let s = 0; s <= steps; s++) {
      const v = Math.round((maxVal / steps) * s);
      const y = yOf(v);
      const gl = document.createElementNS(NS, "line");
      gl.setAttribute("x1", PAD.left); gl.setAttribute("x2", PAD.left + chartW);
      gl.setAttribute("y1", y);        gl.setAttribute("y2", y);
      gl.classList.add("aging-grid-line");
      svg.appendChild(gl);

      const lbl = document.createElementNS(NS, "text");
      lbl.setAttribute("x", PAD.left - 4);
      lbl.setAttribute("y", y + 3);
      lbl.setAttribute("text-anchor", "end");
      lbl.classList.add("aging-axis-label");
      lbl.textContent = v;
      svg.appendChild(lbl);
    }

    // X-axis base line
    const xAxis = document.createElementNS(NS, "line");
    xAxis.setAttribute("x1", PAD.left); xAxis.setAttribute("x2", PAD.left + chartW);
    xAxis.setAttribute("y1", PAD.top + chartH); xAxis.setAttribute("y2", PAD.top + chartH);
    xAxis.classList.add("aging-axis-line");
    svg.appendChild(xAxis);

    // Avg-per-week reference line (dashed)
    const avgPerWeek = n > 0 ? createdSeries.reduce((a, b) => a + b, 0) / n : 0;
    const avgY = yOf(avgPerWeek);
    const avgLine = document.createElementNS(NS, "line");
    avgLine.setAttribute("x1", PAD.left); avgLine.setAttribute("x2", PAD.left + chartW);
    avgLine.setAttribute("y1", avgY);     avgLine.setAttribute("y2", avgY);
    avgLine.setAttribute("stroke", "#f59e0b");
    avgLine.setAttribute("stroke-width", "1.5");
    avgLine.setAttribute("stroke-dasharray", "5 3");
    svg.appendChild(avgLine);

    // X-axis labels (rotated when dense)
    const labelStep = Math.ceil(n / 12);
    checkpoints.forEach((cp, i) => {
      if (i % labelStep !== 0 && i !== n - 1) return;
      const x = xOf(i);
      const lbl = document.createElementNS(NS, "text");
      lbl.setAttribute("x", x);
      lbl.setAttribute("y", PAD.top + chartH + 14);
      lbl.setAttribute("text-anchor", "middle");
      lbl.classList.add("aging-axis-label");
      lbl.textContent = labels[i] ?? "";
      if (n > 8) lbl.setAttribute("transform", `rotate(-35, ${x}, ${PAD.top + chartH + 14})`);
      svg.appendChild(lbl);
    });

    // Bars
    createdSeries.forEach((v, i) => {
      const x   = xOf(i);
      const barH = Math.max(0, chartH - (yOf(v) - PAD.top));
      const rect = document.createElementNS(NS, "rect");
      rect.setAttribute("x",      (x - barW / 2).toFixed(1));
      rect.setAttribute("y",      yOf(v).toFixed(1));
      rect.setAttribute("width",  barW);
      rect.setAttribute("height", barH.toFixed(1));
      rect.setAttribute("rx",     "2");
      rect.setAttribute("fill",   BAR_COLOR);
      rect.classList.add("creation-bar");
      rect.setAttribute("data-idx", i);
      svg.appendChild(rect);
    });

    // Tooltip overlay
    const overlay = document.createElementNS(NS, "rect");
    overlay.setAttribute("x", PAD.left);
    overlay.setAttribute("y", PAD.top);
    overlay.setAttribute("width",  chartW);
    overlay.setAttribute("height", chartH);
    overlay.setAttribute("fill", "transparent");
    svg.appendChild(overlay);

    const crosshair = document.createElementNS(NS, "line");
    crosshair.setAttribute("y1", PAD.top);
    crosshair.setAttribute("y2", PAD.top + chartH);
    crosshair.setAttribute("stroke", "#c0c8d0");
    crosshair.setAttribute("stroke-width", "1");
    crosshair.setAttribute("stroke-dasharray", "3 2");
    crosshair.setAttribute("display", "none");
    svg.appendChild(crosshair);

    const wrap = document.createElement("div");
    wrap.style.position = "relative";
    wrap.appendChild(svg);

    const tooltip = document.createElement("div");
    tooltip.className = "aging-tooltip";
    wrap.appendChild(tooltip);

    overlay.addEventListener("mousemove", e => {
      const rect2 = svg.getBoundingClientRect();
      const mouseX = e.clientX - rect2.left - PAD.left;
      const frac   = Math.max(0, Math.min(1, mouseX / chartW));
      const idx    = Math.round(frac * (n - 1));

      const cx = xOf(idx);
      crosshair.setAttribute("x1", cx);
      crosshair.setAttribute("x2", cx);
      crosshair.setAttribute("display", "");

      tooltip.innerHTML =
        `<strong>${labels[idx]}</strong><br>` +
        `<span style="color:${BAR_COLOR}">■</span> Created: ${createdSeries[idx]}`;

      const svgLeft = rect2.left - wrap.getBoundingClientRect().left;
      const tipX    = svgLeft + cx + 10;
      const tipRight = tipX + 130;
      tooltip.style.left = (tipRight > wrap.offsetWidth ? svgLeft + cx - 135 : tipX) + "px";
      tooltip.style.top  = (PAD.top + 8) + "px";
      tooltip.style.display = "block";
    });

    overlay.addEventListener("mouseleave", () => {
      crosshair.setAttribute("display", "none");
      tooltip.style.display = "none";
    });

    // Legend
    const legend = document.createElement("div");
    legend.className = "aging-legend";
    legend.innerHTML =
      `<div class="aging-legend-item">
         <span class="aging-legend-swatch" style="background:${BAR_COLOR}"></span> Created this week
       </div>
       <div class="aging-legend-item">
         <span class="aging-legend-swatch" style="background:#f59e0b"></span> Avg per week
       </div>`;
    wrap.appendChild(legend);

    // ── Panel title ───────────────────────────────────────────────────────────
    const titleEl = document.createElement("p");
    titleEl.className = "dash-panel-title";
    titleEl.textContent = title ?? "Items Created per Week";

    section.appendChild(titleEl);
    section.appendChild(statBar);
    section.appendChild(wrap);
    return section;
  }



  // ── Col-16 stacked-bar sections ────────────────────────────────────────────

  // Status colour map (normalised)
  const STATUS_COLORS = {
    DONE:      "#22c55e",
    INPROG:    "#f59e0b",
    TODO:      "#3b82d4",
    BLOCKED:   "#ef4444",
    REJECTED:  "#7c5cd8",
    TESTING:   "#67e8f9",
    "WAIT-INFO": "#1f2328",
  };
  const STATUS_ORDER  = ["DONE", "INPROG", "TODO"];
  function statusColor(s) { return STATUS_COLORS[s] ?? "#c0c8d0"; }
  function normaliseStatus(raw) {
    const s = raw != null ? String(raw).trim().toUpperCase() : "";
    return s === "CLOSED" ? "DONE" : s;
  }

  // Severity order for grouping
  const SEV_ORDER = ["Critical", "Major", "Medium", "Low"];

  // Per-section state (bugs / enhancements each have independent selections)
  const verticalSections = {
    bugs: {
      sectionEl:    document.getElementById("col16Section"),
      filtersEl:    document.getElementById("col16Filters"),
      chartAreaEl:  document.getElementById("col16ChartArea"),
      drilldownEl:  document.getElementById("col16Drilldown"),
      drillTitleEl: document.getElementById("col16DrillTitle"),
      drillTableEl: document.getElementById("col16DrillTable"),
      typeFilter:   r => (r[5] != null ? String(r[5]).trim() : "") === "Defect (Bug)",
      selected:     new Set(),
      initialised:  false,
      activeKey:    null,
      showPie:      true,
      pieLabel:     "Bug",
    },
    enhancements: {
      sectionEl:    document.getElementById("col16EnhSection"),
      filtersEl:    document.getElementById("col16EnhFilters"),
      chartAreaEl:  document.getElementById("col16EnhChartArea"),
      drilldownEl:  document.getElementById("col16EnhDrilldown"),
      drillTitleEl: document.getElementById("col16EnhDrillTitle"),
      drillTableEl: document.getElementById("col16EnhDrillTable"),
      typeFilter:   r => (r[5] != null ? String(r[5]).trim() : "") === "Enhancement (SR)",
      selected:     new Set(),
      initialised:  false,
      activeKey:    null,
      showPie:      true,
      pieLabel:     "Enhancement",
    },
  };

  /**
   * Generic vertical stacked-bar section builder.
   * @param {any[][]} rows      — already sidebar-filtered rows
   * @param {object}  sec       — one of the verticalSections entries
   */
  function buildVerticalSection(rows, sec) {
    sec.filtersEl.innerHTML = "";
    sec.chartAreaEl.innerHTML = "";
    sec.drilldownEl.hidden = true;
    sec.activeKey = null;

    // Apply the section's type filter (bugs vs enhancements)
    const typeFiltered = rows.filter(sec.typeFilter);

    // Unique col-16 values (index 15) within the type-filtered set
    const allCol16 = [...new Set(
      typeFiltered.map(r => r.length > 15 && r[15] != null ? String(r[15]).trim() : null).filter(v => v)
    )].sort();

    if (allCol16.length === 0) { sec.sectionEl.hidden = true; return; }
    sec.sectionEl.hidden = false;

    // Initialise selection to all on first build for this file
    if (!sec.initialised) {
      sec.selected    = new Set(allCol16);
      sec.initialised = true;
    }

    // ── Optional: Vertical distribution pie (TODO + INPROG only) ─────────────
    if (sec.showPie) {
      const OPEN_STATUSES = new Set(["TODO", "INPROG"]);
      // Count open items per vertical (col 16, index 15)
      const vertMap = new Map();
      for (const row of typeFiltered) {
        const st = normaliseStatus(row[2]);
        if (!OPEN_STATUSES.has(st)) continue;
        const vert = row.length > 15 && row[15] != null ? String(row[15]).trim() : "(blank)";
        vertMap.set(vert, (vertMap.get(vert) ?? 0) + 1);
      }

      const pieEntries = [...vertMap.entries()].sort((a, b) => b[1] - a[1]);
      const pieTotal   = pieEntries.reduce((s, [, c]) => s + c, 0);

      if (pieTotal > 0) {
        const NS2 = "http://www.w3.org/2000/svg";
        const R2 = 70, cx2 = 80, cy2 = 80;

        const pieSvg = document.createElementNS(NS2, "svg");
        pieSvg.setAttribute("width", "160");
        pieSvg.setAttribute("height", "160");
        pieSvg.setAttribute("viewBox", "0 0 160 160");

        let sa = -Math.PI / 2;
        pieEntries.forEach(([, count], i) => {
          const ang   = (count / pieTotal) * 2 * Math.PI;
          const x1    = cx2 + R2 * Math.cos(sa);
          const y1    = cy2 + R2 * Math.sin(sa);
          const x2    = cx2 + R2 * Math.cos(sa + ang);
          const y2    = cy2 + R2 * Math.sin(sa + ang);
          const large = ang > Math.PI ? 1 : 0;
          const color = COLORS[i % COLORS.length];

          const path2 = document.createElementNS(NS2, "path");
          path2.setAttribute("d", `M${cx2},${cy2} L${x1.toFixed(2)},${y1.toFixed(2)} A${R2},${R2} 0 ${large} 1 ${x2.toFixed(2)},${y2.toFixed(2)} Z`);
          path2.setAttribute("fill", color);
          path2.classList.add("pie-slice");
          pieSvg.appendChild(path2);

          // Number label at slice midpoint
          const midAng = sa + ang / 2;
          const labelR = R2 * 0.65;
          const lx = cx2 + labelR * Math.cos(midAng);
          const ly = cy2 + labelR * Math.sin(midAng);
          const txt = document.createElementNS(NS2, "text");
          txt.setAttribute("x", lx.toFixed(1));
          txt.setAttribute("y", (ly + 4).toFixed(1));
          txt.setAttribute("text-anchor", "middle");
          txt.setAttribute("font-size", "11");
          txt.setAttribute("font-weight", "600");
          txt.setAttribute("fill", "#ffffff");
          txt.textContent = count;
          pieSvg.appendChild(txt);

          sa += ang;
        });

        // Donut hole + centre total
        const hole2 = document.createElementNS(NS2, "circle");
        hole2.setAttribute("cx", cx2); hole2.setAttribute("cy", cy2);
        hole2.setAttribute("r", "26"); hole2.setAttribute("fill", "#ffffff");
        pieSvg.appendChild(hole2);

        const centreLabel = document.createElementNS(NS2, "text");
        centreLabel.setAttribute("x", cx2); centreLabel.setAttribute("y", cy2 + 5);
        centreLabel.setAttribute("text-anchor", "middle");
        centreLabel.setAttribute("font-size", "12");
        centreLabel.setAttribute("font-weight", "600");
        centreLabel.setAttribute("fill", "#1f2328");
        centreLabel.textContent = pieTotal;
        pieSvg.appendChild(centreLabel);

        // Legend
        const pieLegend = document.createElement("div");
        pieLegend.className = "pie-legend";
        pieEntries.forEach(([vertName, count], i) => {
          const pct  = ((count / pieTotal) * 100).toFixed(1);
          const item = document.createElement("div");
          item.className = "legend-item";
          item.innerHTML =
            `<span class="legend-swatch" style="background:${COLORS[i % COLORS.length]}"></span>` +
            `<span>${vertName} — ${count} (${pct}%)</span>`;
          pieLegend.appendChild(item);
        });

        const pieTitle = document.createElement("p");
        pieTitle.className = "dash-panel-title";
        pieTitle.style.cssText = "margin-bottom:0.5rem";
        pieTitle.textContent = `Open ${sec.pieLabel}s by Vertical (TODO + In Progress)`;

        const pieHint = document.createElement("p");
        pieHint.style.cssText = "font-size:0.78rem;color:#57606a;margin-bottom:0.75rem";
        pieHint.textContent = `${pieTotal} open ${sec.pieLabel.toLowerCase()}${pieTotal === 1 ? "" : "s"} across ${pieEntries.length} vertical${pieEntries.length === 1 ? "" : "s"}`;

        const pieWrap2 = document.createElement("div");
        pieWrap2.className = "pie-wrap";
        pieWrap2.appendChild(pieSvg);
        pieWrap2.appendChild(pieLegend);

        const pieContainer = document.createElement("div");
        pieContainer.style.cssText = "margin-bottom:1.25rem";
        pieContainer.appendChild(pieTitle);
        pieContainer.appendChild(pieHint);
        pieContainer.appendChild(pieWrap2);

        sec.filtersEl.appendChild(pieContainer);
      }
    }

    // ── Vertical filter chips ─────────────────────────────────────────────────
    const chipWrap = document.createElement("div");
    chipWrap.className = "filter-group";

    const allSel = allCol16.every(v => sec.selected.has(v));
    chipWrap.innerHTML = `
      <div class="filter-group-header">
        <span class="filter-group-label">Verticals</span>
        <button class="filter-toggle-all">${allSel ? "Deselect all" : "Select all"}</button>
      </div>
      <div class="filter-chips"></div>`;

    chipWrap.querySelector(".filter-toggle-all").addEventListener("click", () => {
      if (allCol16.every(v => sec.selected.has(v))) { sec.selected.clear(); }
      else { sec.selected = new Set(allCol16); }
      buildVerticalSection(filteredRows(), sec);
    });

    const chipsEl = chipWrap.querySelector(".filter-chips");
    for (const val of allCol16) {
      const chip = document.createElement("label");
      chip.className = "filter-chip" + (sec.selected.has(val) ? " selected" : "");
      chip.title = val;
      chip.innerHTML = `<input type="checkbox" ${sec.selected.has(val) ? "checked" : ""}><span>${val}</span>`;
      chip.querySelector("input").addEventListener("change", () => {
        chip.classList.toggle("selected");
        sec.selected.has(val) ? sec.selected.delete(val) : sec.selected.add(val);
        buildVerticalSection(filteredRows(), sec);
      });
      chipsEl.appendChild(chip);
    }
    sec.filtersEl.appendChild(chipWrap);

    // ── Filter by selected verticals ──────────────────────────────────────────
    const filtered = typeFiltered.filter(r => {
      const v = r.length > 15 && r[15] != null ? String(r[15]).trim() : null;
      return v && sec.selected.has(v);
    });

    if (filtered.length === 0) {
      sec.chartAreaEl.innerHTML = "<p style='color:#57606a;font-size:0.875rem;margin-top:0.5rem'>No rows match the selected verticals.</p>";
      return;
    }

    // ── Build data: { sevLabel → { status → [rows] } } ───────────────────────
    const sevMap = new Map();
    for (const row of filtered) {
      const sev    = severityLabel(row.length > 11 ? row[11] : null);
      const status = normaliseStatus(row[2]);
      if (!sevMap.has(sev)) sevMap.set(sev, new Map());
      const sMap = sevMap.get(sev);
      if (!sMap.has(status)) sMap.set(status, []);
      sMap.get(status).push(row);
    }

    const sevKeys = [...SEV_ORDER.filter(s => sevMap.has(s)),
                     ...[...sevMap.keys()].filter(s => !SEV_ORDER.includes(s))];

    const allStatuses = [...new Set(filtered.map(r => normaliseStatus(r[2])))];
    const statusKeys  = [...STATUS_ORDER.filter(s => allStatuses.includes(s)),
                         ...allStatuses.filter(s => !STATUS_ORDER.includes(s))];

    const maxTotal = Math.max(...sevKeys.map(sev => {
      const sMap = sevMap.get(sev);
      return [...sMap.values()].reduce((a, arr) => a + arr.length, 0);
    }));

    // ── Legend ────────────────────────────────────────────────────────────────
    const legend = document.createElement("div");
    legend.className = "stacked-legend";
    for (const st of statusKeys) {
      legend.innerHTML += `<div class="stacked-legend-item">
        <span class="stacked-legend-swatch" style="background:${statusColor(st)}"></span>${st}
      </div>`;
    }
    sec.chartAreaEl.appendChild(legend);

    // ── Bars ──────────────────────────────────────────────────────────────────
    for (const sev of sevKeys) {
      const sMap  = sevMap.get(sev);
      const total = [...sMap.values()].reduce((a, arr) => a + arr.length, 0);

      const row = document.createElement("div");
      row.className = "stacked-bar-row";

      const labelEl = document.createElement("span");
      labelEl.className = "stacked-bar-label";
      labelEl.title = sev;
      labelEl.textContent = sev;
      row.appendChild(labelEl);

      const track = document.createElement("div");
      track.className = "stacked-bar-track";

      for (const st of statusKeys) {
        const segRows = sMap.get(st) ?? [];
        if (segRows.length === 0) continue;
        const pct = (segRows.length / maxTotal) * 100;
        const key = `${sev}|${st}`;
        const seg = document.createElement("div");
        seg.className = "stacked-seg" + (sec.activeKey === key ? " active" : "");
        seg.style.width      = pct + "%";
        seg.style.background = statusColor(st);
        seg.title = `${sev} / ${st}: ${segRows.length}`;
        seg.addEventListener("click", () => {
          if (sec.activeKey === key) {
            sec.activeKey = null;
            sec.drilldownEl.hidden = true;
            seg.classList.remove("active");
          } else {
            sec.chartAreaEl.querySelectorAll(".stacked-seg.active")
              .forEach(el => el.classList.remove("active"));
            sec.activeKey = key;
            seg.classList.add("active");
            showVerticalDrilldown(sec, sev, st, segRows);
          }
        });
        track.appendChild(seg);
      }
      row.appendChild(track);

      const totalEl = document.createElement("span");
      totalEl.className = "stacked-bar-total";
      totalEl.textContent = total;
      row.appendChild(totalEl);

      const inlineLegend = document.createElement("div");
      inlineLegend.className = "stacked-bar-inline-legend";
      for (const st of statusKeys) {
        const segRows = sMap.get(st) ?? [];
        if (segRows.length === 0) continue;
        const item = document.createElement("span");
        item.className = "stacked-bar-inline-item";
        item.innerHTML =
          `<span class="stacked-bar-inline-swatch" style="background:${statusColor(st)}"></span>` +
          `${st}: ${segRows.length}`;
        inlineLegend.appendChild(item);
      }
      row.appendChild(inlineLegend);

      sec.chartAreaEl.appendChild(row);
    }
  }

  function showVerticalDrilldown(sec, sev, status, rows) {
    sec.drillTitleEl.textContent =
      `${sev} / ${status} — ${rows.length} ticket${rows.length === 1 ? "" : "s"}`;

    sec.drillTableEl.innerHTML =
      `<thead><tr><th>ID</th><th>Description</th><th>Status</th><th>Severity</th><th>Fix Version</th></tr></thead>`;
    const tbody = document.createElement("tbody");

    const sorted = [...rows].sort((a, b) => {
      const fa = a.length > 0 && a[a.length - 1] != null ? String(a[a.length - 1]) : "";
      const fb = b.length > 0 && b[b.length - 1] != null ? String(b[b.length - 1]) : "";
      return fb.localeCompare(fa, undefined, { numeric: true, sensitivity: "base" });
    });

    for (const row of sorted) {
      const tr = document.createElement("tr");
      [row[0], row[1], row[2]].forEach(val => {
        const td = document.createElement("td");
        td.textContent = val != null ? String(val) : "";
        tr.appendChild(td);
      });
      const sevTd = document.createElement("td");
      sevTd.textContent = severityLabel(row.length > 11 ? row[11] : null);
      tr.appendChild(sevTd);
      const fixTd = document.createElement("td");
      fixTd.textContent = row.length > 0 && row[row.length - 1] != null
        ? String(row[row.length - 1]) : "";
      tr.appendChild(fixTd);
      tbody.appendChild(tr);
    }
    sec.drillTableEl.appendChild(tbody);
    sec.drilldownEl.hidden = false;
  }

  // Keep old name as a thin wrapper so any future callers still work
  function buildCol16Section(rows) {
    buildVerticalSection(allRows, verticalSections.bugs);         // independent of filters
    buildVerticalSection(allRows, verticalSections.enhancements); // independent of filters
  }

  // ── Chart rendering ────────────────────────────────────────────────────────
  function tally(rows, colIndex) {
    const map = new Map();
    for (const row of rows) {
      if (row.length <= colIndex) continue;
      let val = row[colIndex] != null ? String(row[colIndex]).trim() : "(blank)";
      // Column 3 (index 2): treat CLOSED as DONE
      if (colIndex === 2 && val === "CLOSED") val = "DONE";
      map.set(val, (map.get(val) ?? 0) + 1);
    }
    return [...map.entries()].sort((a, b) => b[1] - a[1]);
  }

  // Maps a raw col-3 status to a responsibility group label
  const RESP_GROUP_MAP = {
    DONE: "Resolved", CLOSED: "Resolved",
    TODO: "On OTE Team", INPROG: "On OTE Team",
    TESTING: "On SHSO Team", "WAIT-INFO": "On SHSO Team",
    BLOCKED: "Blocked", REJECTED: "Rejected",
  };
  const RESP_GROUP_COLORS = {
    "Resolved":     "#22c55e",   // light green
    "On OTE Team":  "#3b82d4",   // blue
    "On SHSO Team": "#67e8f9",   // light blue
    "Blocked":      "#ef4444",   // red
    "Rejected":     "#7c5cd8",   // purple
  };
  function respGroupLabel(raw) {
    const s = raw != null ? String(raw).trim().toUpperCase() : "";
    return RESP_GROUP_MAP[s] ?? s;
  }

  /** Tally rows by a custom label function on col-3 */
  function groupedTally(rows, labelFn) {
    const map = new Map();
    for (const row of rows) {
      const lbl = labelFn(row[2]);
      map.set(lbl, (map.get(lbl) ?? 0) + 1);
    }
    return [...map.entries()].sort((a, b) => b[1] - a[1]);
  }

  // ── KPI computation ────────────────────────────────────────────────────────

  /**
   * Compute all resolution/velocity KPIs from a set of filtered rows.
   * col 3  (index 2)  = status
   * col 4  (index 3)  = status date  (used as resolution date for DONE/CLOSED)
   * col 11 (index 10) = created date
   */
  function computeKpis(rows) {
    const RESOLVED = new Set(["DONE", "CLOSED"]);
    const EXCLUDED  = new Set(["TESTING", "WAIT-INFO", "BLOCKED", "REJECTED"]);
    const now = Date.now();
    const MS_DAY = 86400000;
    const thirtyDaysAgo = now - 30 * MS_DAY;

    let resolved = 0;
    let openOver30 = 0;
    let throughputLast30 = 0;
    let ticketsOnOte = 0;
    let ticketsOnShso = 0;
    let shsoOver30 = 0;

    for (const row of rows) {
      const statusRaw = row[2] != null ? String(row[2]).trim().toUpperCase() : "";

      if (statusRaw === "TODO" || statusRaw === "INPROG") ticketsOnOte++;
      if (statusRaw === "TESTING" || statusRaw === "WAIT-INFO") {
        ticketsOnShso++;
        const statusDate = parseDate(row[3]);
        if (statusDate && (now - statusDate.getTime()) > 30 * MS_DAY) shsoOver30++;
      }

      if (RESOLVED.has(statusRaw)) {
        resolved++;
        const statusDate = parseDate(row[3]);
        if (statusDate && statusDate.getTime() >= thirtyDaysAgo) {
          throughputLast30++;
        }
      } else if (!EXCLUDED.has(statusRaw)) {
        // open tickets older than 30 days — exclude TESTING/WAIT-INFO/BLOCKED/REJECTED
        const createdDate = parseDate(row[10]);
        if (createdDate && (now - createdDate.getTime()) > 30 * MS_DAY) {
          openOver30++;
        }
      }
    }

    // Type breakdown entries (sorted by count desc)
    const typeMap = new Map();
    for (const row of rows) {
      const t = row[5] != null ? String(row[5]).trim() : "(blank)";
      typeMap.set(t, (typeMap.get(t) ?? 0) + 1);
    }
    const typeEntries = [...typeMap.entries()].sort((a, b) => b[1] - a[1]);

    return {
      total: rows.length,
      resolved,
      resolutionRate: rows.length > 0 ? (resolved / rows.length) * 100 : 0,
      openOver30,
      throughputLast30,
      ticketsOnOte,
      ticketsOnShso,
      shsoOver30,
      typeEntries,
    };
  }

  /**
   * Build and return the KPI strip + drilldown container.
   * Each tile is clickable; clicking toggles a shared drilldown panel below the strip.
   */
  function renderKpiStrip(kpis, allRows) {
    const MS_DAY = 86400000;
    const now = Date.now();
    const thirtyDaysAgo = now - 30 * MS_DAY;
    const RESOLVED = new Set(["DONE", "CLOSED"]);
    const SEVERITY_ORDER_DD = ["Critical", "Major", "Medium", "Low"];

    // ── Container holds strip + drilldown ─────────────────────────────────────
    const container = document.createElement("div");

    // ── Strip: tiles ──────────────────────────────────────────────────────────
    const strip = document.createElement("div");
    strip.className = "kpi-strip";

    // ── Drilldown panel (shared; one at a time) ──────────────────────────────
    const drilldown = document.createElement("div");
    drilldown.className = "kpi-drilldown";
    drilldown.hidden = true;

    const drillTitle = document.createElement("p");
    drillTitle.className = "drilldown-title";

    const drillSevFilter = document.createElement("div");
    drillSevFilter.className = "drilldown-sev-filter";

    const drillTable = document.createElement("table");
    drillTable.className = "drilldown-table";

    drilldown.appendChild(drillTitle);
    drilldown.appendChild(drillSevFilter);
    drilldown.appendChild(drillTable);

    // Track active KPI + tiles
    let activeKpi = null;
    let matchedRows = [];
    let activeSevs = new Set();
    const tiles = [];

    // ── Render the table using matched rows × severity filter ────────────────
    function renderDrillTable(showDaysCol) {
      const visible = matchedRows.filter(row => {
        const sev = severityLabel(row.length > 11 ? row[11] : null);
        return activeSevs.has(sev);
      });

      drillTitle.textContent =
        `${activeKpi} — ${visible.length} ticket${visible.length === 1 ? "" : "s"}` +
        (visible.length < matchedRows.length ? ` (${matchedRows.length} total)` : "");

      // Table header — add "Days to Resolve" if requested
      const headers = ["ID", "Description", "Status", "Severity", "Fix Version"];
      if (showDaysCol) headers.splice(3, 0, "Days to Resolve");

      drillTable.innerHTML = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr></thead>`;
      const tbody = document.createElement("tbody");

      for (const row of visible) {
        const tr = document.createElement("tr");

        // ID, Desc, Status
        [row[0], row[1], row[2]].forEach(val => {
          const td = document.createElement("td");
          td.textContent = val != null ? String(val) : "";
          tr.appendChild(td);
        });

        // Days to Resolve (if enabled)
        if (showDaysCol) {
          const td = document.createElement("td");
          const statusDate = parseDate(row[3]);
          const createdDate = parseDate(row[10]);
          if (statusDate && createdDate && statusDate >= createdDate) {
            td.textContent = Math.round((statusDate.getTime() - createdDate.getTime()) / MS_DAY);
          } else {
            td.textContent = "—";
          }
          tr.appendChild(td);
        }

        // Severity
        const sevTd = document.createElement("td");
        sevTd.textContent = severityLabel(row.length > 11 ? row[11] : null);
        tr.appendChild(sevTd);

        // Fix Version
        const fixTd = document.createElement("td");
        fixTd.textContent = row.length > 0 && row[row.length - 1] != null
          ? String(row[row.length - 1]) : "";
        tr.appendChild(fixTd);

        tbody.appendChild(tr);
      }
      drillTable.appendChild(tbody);
    }

    function buildSevFilter() {
      drillSevFilter.innerHTML = "";
      const present = new Set(matchedRows.map(r => severityLabel(r.length > 11 ? r[11] : null)));
      const ordered = SEVERITY_ORDER_DD.filter(l => present.has(l));
      for (const l of present) { if (!SEVERITY_ORDER_DD.includes(l)) ordered.push(l); }
      if (ordered.length <= 1) return;

      const lbl = document.createElement("span");
      lbl.className = "drilldown-sev-label";
      lbl.textContent = "Severity:";
      drillSevFilter.appendChild(lbl);

      for (const sev of ordered) {
        const chip = document.createElement("label");
        chip.className = "filter-chip" + (activeSevs.has(sev) ? " selected" : "");
        chip.title = sev;
        chip.innerHTML = `<input type="checkbox" ${activeSevs.has(sev) ? "checked" : ""}><span>${sev}</span>`;
        chip.querySelector("input").addEventListener("change", () => {
          chip.classList.toggle("selected");
          activeSevs.has(sev) ? activeSevs.delete(sev) : activeSevs.add(sev);
          renderDrillTable(showDaysCol);
        });
        drillSevFilter.appendChild(chip);
      }
    }

    /** Track if the current KPI's drilldown should show "Days to Resolve" column */
    let showDaysCol = false;

    function activateKpi(kpiName, rowsFilterFn, showDays) {
      // Toggle off if clicking the same KPI
      if (activeKpi === kpiName) {
        activeKpi = null;
        drilldown.hidden = true;
        tiles.forEach(t => t.classList.remove("active"));
        return;
      }

      activeKpi = kpiName;
      showDaysCol = showDays;
      matchedRows = allRows.filter(rowsFilterFn);

      // Sort by fix version descending
      matchedRows.sort((a, b) => {
        const fa = a.length > 0 && a[a.length - 1] != null ? String(a[a.length - 1]) : "";
        const fb = b.length > 0 && b[b.length - 1] != null ? String(b[b.length - 1]) : "";
        return fb.localeCompare(fa, undefined, { numeric: true, sensitivity: "base" });
      });

      activeSevs = new Set(matchedRows.map(r => severityLabel(r.length > 11 ? r[11] : null)));
      buildSevFilter();
      renderDrillTable(showDaysCol);
      drilldown.hidden = false;

      // Update active tile state
      tiles.forEach(t => t.classList.remove("active"));
      const idx = tiles.findIndex(t => t.dataset.kpiName === kpiName);
      if (idx >= 0) tiles[idx].classList.add("active");
    }

    function tile(value, label, sub, accentClass, kpiName, filterFn, showDays) {
      const t = document.createElement("div");
      t.className = `kpi-tile ${accentClass}`;
      t.dataset.kpiName = kpiName;
      t.innerHTML =
        `<div class="kpi-value">${value}</div>` +
        `<div class="kpi-label">${label}</div>` +
        (sub ? `<div class="kpi-sub">${sub}</div>` : "");
      t.addEventListener("click", () => activateKpi(kpiName, filterFn, showDays));
      return t;
    }

    // ── Tile definitions ──────────────────────────────────────────────────────

    // 1. Total tickets
    tiles.push(strip.appendChild(tile(
      kpis.total,
      "Total Tickets",
      null,
      "kpi-accent-blue",
      "Total Tickets",
      () => true, // all rows
      false
    )));

    // 2. Resolved
    tiles.push(strip.appendChild(tile(
      kpis.resolved,
      "Resolved",
      null,
      "kpi-accent-green",
      "Resolved",
      row => RESOLVED.has((row[2] != null ? String(row[2]).trim().toUpperCase() : "")),
      true  // show days column
    )));

    // 2b. Tickets on OTE (TODO + INPROG)
    const OTE_STATUSES = new Set(["TODO", "INPROG"]);
    tiles.push(strip.appendChild(tile(
      kpis.ticketsOnOte,
      "Tickets on OTE",
      "TODO + INPROG",
      "kpi-accent-blue",
      "Tickets on OTE",
      row => OTE_STATUSES.has((row[2] != null ? String(row[2]).trim().toUpperCase() : "")),
      false
    )));

    // 2c. On SHSO (TESTING + WAIT-INFO)
    const SHSO_STATUSES = new Set(["TESTING", "WAIT-INFO"]);
    tiles.push(strip.appendChild(tile(
      kpis.ticketsOnShso,
      "On SHSO",
      "TESTING + WAIT-INFO",
      "kpi-accent-purple",
      "On SHSO",
      row => SHSO_STATUSES.has((row[2] != null ? String(row[2]).trim().toUpperCase() : "")),
      false
    )));

    // 3. Resolution rate (same filter as Resolved)
    tiles.push(strip.appendChild(tile(
      kpis.resolutionRate.toFixed(1) + "%",
      "Resolution Rate",
      `${kpis.resolved} of ${kpis.total}`,
      "kpi-accent-green",
      "Resolution Rate",
      row => RESOLVED.has((row[2] != null ? String(row[2]).trim().toUpperCase() : "")),
      true
    )));

    // 4. Resolved last 30 days
    tiles.push(strip.appendChild(tile(
      kpis.throughputLast30,
      "Resolved (Last 30 d)",
      "by status date",
      "kpi-accent-amber",
      "Resolved (Last 30 d)",
      row => {
        const statusRaw = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
        if (!RESOLVED.has(statusRaw)) return false;
        const statusDate = parseDate(row[3]);
        return statusDate && statusDate.getTime() >= thirtyDaysAgo;
      },
      true
    )));

    // 5. SHSO > 30 days (TESTING + WAIT-INFO with status date older than 30 days)
    const SHSO_OVER30_STATUSES = new Set(["TESTING", "WAIT-INFO"]);
    tiles.push(strip.appendChild(tile(
      kpis.shsoOver30,
      "SHSO > 30 Days",
      "by status date",
      kpis.shsoOver30 > 0 ? "kpi-accent-red" : "kpi-accent-green",
      "SHSO > 30 Days",
      row => {
        const statusRaw = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
        if (!SHSO_OVER30_STATUSES.has(statusRaw)) return false;
        const statusDate = parseDate(row[3]);
        return statusDate && (now - statusDate.getTime()) > 30 * MS_DAY;
      },
      false
    )));

    // 6. Open > 30 days
    tiles.push(strip.appendChild(tile(
      kpis.openOver30,
      "Open > 30 Days",
      "by created date",
      kpis.openOver30 > 0 ? "kpi-accent-red" : "kpi-accent-green",
      "Open > 30 Days",
      row => {
        const statusRaw = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
        if (RESOLVED.has(statusRaw)) return false;
        const EXCLUDED = new Set(["TESTING", "WAIT-INFO", "BLOCKED", "REJECTED"]);
        if (EXCLUDED.has(statusRaw)) return false;
        const createdDate = parseDate(row[10]);
        return createdDate && (now - createdDate.getTime()) > 30 * MS_DAY;
      },
      false  // "days to resolve" doesn't apply to open tickets
    )));

    // ── Shared helper: render only the type breakdown bar chart in the drilldown ─
    function showTypeBreakdownDrilldown(kpiName, tileEl) {
      // Toggle off if already active
      if (activeKpi === kpiName) {
        activeKpi = null;
        drilldown.hidden = true;
        tiles.forEach(t => t.classList.remove("active"));
        return;
      }

      activeKpi = kpiName;
      tiles.forEach(t => t.classList.remove("active"));
      tileEl.classList.add("active");

      drillTitle.textContent = `Type Breakdown — ${kpis.typeEntries.length} type${kpis.typeEntries.length === 1 ? "" : "s"}`;
      drillSevFilter.innerHTML = "";
      drillTable.innerHTML = "";

      const barSection = document.createElement("div");
      barSection.style.cssText = "margin-bottom:0.5rem";

      const maxCount = kpis.typeEntries[0]?.[1] ?? 1;
      kpis.typeEntries.forEach(([typeName, count], i) => {
        const pct   = Math.round((count / maxCount) * 100);
        const color = COLORS[i % COLORS.length];
        const barRow = document.createElement("div");
        barRow.className = "bar-row";
        barRow.innerHTML =
          `<span class="bar-label" title="${typeName}">${typeName}</span>` +
          `<span class="bar-track"><span class="bar-fill" style="width:${pct}%;background:${color}"></span></span>` +
          `<span class="bar-count">${count}</span>`;
        barSection.appendChild(barRow);
      });

      drilldown.innerHTML = "";
      drilldown.appendChild(drillTitle);
      drilldown.appendChild(barSection);
      drilldown.hidden = false;
    }

    // ── Type Breakdown tile ────────────────────────────────────────────────────
    const typeTile = document.createElement("div");
    typeTile.className = "kpi-tile kpi-accent-blue";
    typeTile.dataset.kpiName = "Type Breakdown";
    const topType = kpis.typeEntries[0];
    typeTile.innerHTML =
      `<div class="kpi-value" style="font-size:1.1rem;line-height:1.3;word-break:break-word">${topType ? topType[0] : "—"}</div>` +
      `<div class="kpi-label">Type Breakdown</div>` +
      (topType ? `<div class="kpi-sub">${kpis.typeEntries.length} type${kpis.typeEntries.length === 1 ? "" : "s"} · top: ${topType[1]}</div>` : "");
    tiles.push(typeTile);
    strip.appendChild(typeTile);
    typeTile.addEventListener("click", () => showTypeBreakdownDrilldown("Type Breakdown", typeTile));

    container.appendChild(strip);
    container.appendChild(drilldown);
    return container;
  }

  function renderCharts(rows) {
    const total = rows.length;
    if (total === 0) {
      // Show a "no data" message in the KPI row area
      const msg = document.createElement("p");
      msg.style.cssText = "color:#57606a;font-size:0.875rem";
      msg.textContent = "No rows match the selected filters.";
      dashKpiRow.appendChild(msg);
      dashKpiRow.hidden = false;
      dashGrid.hidden = false;
      return;
    }

    // ── KPI strip ──────────────────────────────────────────────────────────
    dashKpiRow.appendChild(renderKpiStrip(computeKpis(rows), rows));
    dashKpiRow.hidden = false;

    // ── Chart grid: two pies only (Type Breakdown moved to KPI strip) ──────
    const statusLabelFn = raw => {
      const s = raw != null ? String(raw).trim() : "";
      return s === "CLOSED" ? "DONE" : s;
    };
    const col16El = document.getElementById("col16Section");
    // Status pie uses STATUS_COLORS by label name; responsibility pie uses positional COLORS
    dashGrid.insertBefore(pieBlock("Ticket Status Breakdown", tally(rows, 2), rows, statusLabelFn, lbl => statusColor(lbl)), col16El);
    const statsUnresolved = rows.filter(r => { const s = r[2] != null ? String(r[2]).trim().toUpperCase() : ""; return s !== "DONE" && s !== "CLOSED"; });
    dashGrid.insertBefore(pieBlock("Responsibility Breakdown", groupedTally(statsUnresolved, respGroupLabel), statsUnresolved, respGroupLabel, lbl => RESP_GROUP_COLORS[lbl] ?? COLORS[0]), col16El);
    dashGrid.hidden = false;
  }

  function barBlock(title, entries, total) {
    const block = document.createElement("div");
    block.className = "chart-block";
    const maxCount = entries[0]?.[1] ?? 1;
    let html = `<h3>${title}</h3>`;
    entries.forEach(([label, count], i) => {
      const pct = Math.round((count / maxCount) * 100);
      const color = COLORS[i % COLORS.length];
      html += `
        <div class="bar-row">
          <span class="bar-label" title="${label}">${label}</span>
          <span class="bar-track">
            <span class="bar-fill" style="width:${pct}%;background:${color}"></span>
          </span>
          <span class="bar-count">${count}</span>
        </div>`;
    });
    block.innerHTML = html;
    return block;
  }

  function pieBlock(title, entries, rows, labelFn, colorFn) {
    // Default labelFn: identity (use raw col-3 value as-is)
    if (!labelFn) labelFn = raw => raw != null ? String(raw).trim() : "";
    // Default colorFn: positional palette
    if (!colorFn) colorFn = (_, i) => COLORS[i % COLORS.length];
    const block = document.createElement("div");
    block.className = "chart-block";

    const total = entries.reduce((s, [, c]) => s + c, 0);
    const R = 80, cx = 90, cy = 90;
    const NS = "http://www.w3.org/2000/svg";

    // ── build SVG ────────────────────────────────────────────────────────────
    const svg = document.createElementNS(NS, "svg");
    svg.setAttribute("width", "180");
    svg.setAttribute("height", "180");
    svg.setAttribute("viewBox", "0 0 180 180");

    // donut hole + centre count — appended last so they sit on top
    const hole  = document.createElementNS(NS, "circle");
    hole.setAttribute("cx", cx); hole.setAttribute("cy", cy);
    hole.setAttribute("r", "30"); hole.setAttribute("fill", "#ffffff");

    const label = document.createElementNS(NS, "text");
    label.setAttribute("x", cx); label.setAttribute("y", cy + 5);
    label.setAttribute("text-anchor", "middle");
    label.setAttribute("font-size", "13");
    label.setAttribute("font-weight", "600");
    label.setAttribute("fill", "#1f2328");
    label.textContent = total;

    // ── legend container ─────────────────────────────────────────────────────
    const legendEl = document.createElement("div");
    legendEl.className = "pie-legend";

    // ── drilldown panel ──────────────────────────────────────────────────────
    const drilldown = document.createElement("div");
    drilldown.className = "drilldown";
    drilldown.hidden = true;

    const drillTitle   = document.createElement("p");
    drillTitle.className = "drilldown-title";

    const drillSevFilter = document.createElement("div");
    drillSevFilter.className = "drilldown-sev-filter";

    const drillTable = document.createElement("table");
    drillTable.className = "drilldown-table";

    drilldown.appendChild(drillTitle);
    drilldown.appendChild(drillSevFilter);
    drilldown.appendChild(drillTable);

    // Track which slice is currently active
    let activeLabel    = null;
    let matchedRows    = [];          // rows for the active slice (pre-severity filter)
    let activeSevs     = new Set();   // which severity labels are selected in drilldown
    const sliceEls     = [];
    const legendItems  = [];

    const SEVERITY_ORDER_DD = ["Critical", "Major", "Medium", "Low"];

    /** Re-render the table using current matchedRows × activeSevs */
    function renderDrillTable() {
      const visible = matchedRows.filter(row => {
        const sev = severityLabel(row.length > 11 ? row[11] : null);
        return activeSevs.has(sev);
      });

      drillTitle.textContent =
        `${activeLabel} — ${visible.length} ticket${visible.length === 1 ? "" : "s"}` +
        (visible.length < matchedRows.length ? ` (${matchedRows.length} total)` : "");

      drillTable.innerHTML =
        `<thead><tr><th>ID</th><th>Description</th><th>Status</th><th>Severity</th><th>Fix Version</th></tr></thead>`;
      const tbody = document.createElement("tbody");
      for (const row of visible) {
        const tr = document.createElement("tr");
        [row[0], row[1], row[2]].forEach(val => {
          const td = document.createElement("td");
          td.textContent = val != null ? String(val) : "";
          tr.appendChild(td);
        });
        const sevTd = document.createElement("td");
        sevTd.textContent = severityLabel(row.length > 11 ? row[11] : null);
        tr.appendChild(sevTd);
        const fixTd = document.createElement("td");
        fixTd.textContent = row.length > 0 && row[row.length - 1] != null
          ? String(row[row.length - 1]) : "";
        tr.appendChild(fixTd);
        tbody.appendChild(tr);
      }
      drillTable.appendChild(tbody);
    }

    /** Build severity chip row from matched rows */
    function buildSevFilter() {
      drillSevFilter.innerHTML = "";

      // Collect severity labels present in matched rows, in priority order
      const present = new Set(matchedRows.map(r => severityLabel(r.length > 11 ? r[11] : null)));
      const ordered = SEVERITY_ORDER_DD.filter(l => present.has(l));
      for (const l of present) { if (!SEVERITY_ORDER_DD.includes(l)) ordered.push(l); }

      if (ordered.length <= 1) return; // no point showing filter for 0 or 1 values

      const lbl = document.createElement("span");
      lbl.className = "drilldown-sev-label";
      lbl.textContent = "Severity:";
      drillSevFilter.appendChild(lbl);

      for (const sev of ordered) {
        const chip = document.createElement("label");
        chip.className = "filter-chip" + (activeSevs.has(sev) ? " selected" : "");
        chip.title = sev;
        chip.innerHTML = `<input type="checkbox" ${activeSevs.has(sev) ? "checked" : ""}><span>${sev}</span>`;
        chip.querySelector("input").addEventListener("change", () => {
          chip.classList.toggle("selected");
          activeSevs.has(sev) ? activeSevs.delete(sev) : activeSevs.add(sev);
          renderDrillTable();
        });
        drillSevFilter.appendChild(chip);
      }
    }

    function activateSlice(sliceLabel) {
      // Toggle off if clicking the same slice again
      if (activeLabel === sliceLabel) {
        activeLabel = null;
        drilldown.hidden = true;
      } else {
        activeLabel = sliceLabel;

        // Collect matching rows using the chart's own labelFn
        matchedRows = rows.filter(row => row.length >= 1 && labelFn(row[2]) === sliceLabel);

        // Sort descending by Fix Version (last column)
        matchedRows.sort((a, b) => {
          const fa = a.length > 0 && a[a.length - 1] != null ? String(a[a.length - 1]) : "";
          const fb = b.length > 0 && b[b.length - 1] != null ? String(b[b.length - 1]) : "";
          return fb.localeCompare(fa, undefined, { numeric: true, sensitivity: "base" });
        });

        // Initialise severity filter to all present values
        activeSevs = new Set(matchedRows.map(r => severityLabel(r.length > 11 ? r[11] : null)));

        buildSevFilter();
        renderDrillTable();
        drilldown.hidden = false;
      }

      // Update active state on slices and legend items
      sliceEls.forEach((el, i) => {
        el.classList.toggle("active", entries[i][0] === activeLabel);
      });
      legendItems.forEach((el, i) => {
        el.classList.toggle("active", entries[i][0] === activeLabel);
      });
    }

    // ── draw slices ──────────────────────────────────────────────────────────
    let startAngle = -Math.PI / 2;

    entries.forEach(([lbl, count], i) => {
      const angle = (count / total) * 2 * Math.PI;
      const x1 = cx + R * Math.cos(startAngle);
      const y1 = cy + R * Math.sin(startAngle);
      const x2 = cx + R * Math.cos(startAngle + angle);
      const y2 = cy + R * Math.sin(startAngle + angle);
      const large = angle > Math.PI ? 1 : 0;
      const color = colorFn(lbl, i);

      const path = document.createElementNS(NS, "path");
      path.setAttribute("d", `M${cx},${cy} L${x1.toFixed(2)},${y1.toFixed(2)} A${R},${R} 0 ${large} 1 ${x2.toFixed(2)},${y2.toFixed(2)} Z`);
      path.setAttribute("fill", color);
      path.classList.add("pie-slice");
      path.addEventListener("click", () => activateSlice(lbl));
      svg.appendChild(path);
      sliceEls.push(path);

      // Legend row
      const pct  = ((count / total) * 100).toFixed(1);
      const item = document.createElement("div");
      item.className = "legend-item";
      item.innerHTML = `<span class="legend-swatch" style="background:${color}"></span>
        <span>${lbl} — ${count} (${pct}%)</span>`;
      item.addEventListener("click", () => activateSlice(lbl));
      legendEl.appendChild(item);
      legendItems.push(item);

      startAngle += angle;
    });

    svg.appendChild(hole);
    svg.appendChild(label);

    // ── assemble block ───────────────────────────────────────────────────────
    const heading = document.createElement("h3");
    heading.textContent = title;

    const hint = document.createElement("p");
    hint.style.cssText = "font-size:0.78rem;color:#57606a;margin-bottom:0.75rem";
    hint.textContent = "Click a slice or legend item to list its tickets.";

    const pieWrap = document.createElement("div");
    pieWrap.className = "pie-wrap";
    pieWrap.appendChild(svg);
    pieWrap.appendChild(legendEl);

    block.appendChild(heading);
    block.appendChild(hint);
    block.appendChild(pieWrap);
    block.appendChild(drilldown);
    return block;
  }

  // ── Create Presentation ────────────────────────────────────────────────────

  // ── Export Backlog of OTE ──────────────────────────────────────────────────
  document.getElementById("exportOteBtn").addEventListener("click", () => {
    const OTE_STATUSES = new Set(["TODO", "INPROG"]);
    const oteRows = sharedRows.filter(row => {
      const s = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      return OTE_STATUSES.has(s);
    });

    // Prepend header row if the file had one
    const exportRows = (homeHeaderChk.checked && rawRows.length) ? [rawRows[0], ...oteRows] : oteRows;

    const ws = XLSX.utils.aoa_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OTE Backlog");

    const baseName = sharedFileName.replace(/\.xlsx$/i, "");
    XLSX.writeFile(wb, `${baseName}_OTE_Backlog.xlsx`);
  });

  // ── Export SHSO Pending ────────────────────────────────────────────────────
  document.getElementById("exportShsoBtn").addEventListener("click", () => {
    const SHSO_STATUSES = new Set(["TESTING", "WAIT-INFO"]);
    const shsoRows = sharedRows.filter(row => {
      const s = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
      return SHSO_STATUSES.has(s);
    });

    // Prepend header row if the file had one
    const exportRows = (homeHeaderChk.checked && rawRows.length) ? [rawRows[0], ...shsoRows] : shsoRows;

    const ws = XLSX.utils.aoa_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SHSO Pending");

    const baseName = sharedFileName.replace(/\.xlsx$/i, "");
    XLSX.writeFile(wb, `${baseName}_SHSO_Pending.xlsx`);
  });

  const createPptxBtn = document.getElementById("createPptxBtn");

  const PPTX_BTN_HTML =
    `<svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true"><rect x="1" y="2" width="10" height="13" rx="1.5" fill="#fff" stroke="#c2540a" stroke-width="1.2"/><rect x="5" y="0" width="10" height="13" rx="1.5" fill="#fff7f0" stroke="#c2540a" stroke-width="1.2"/><rect x="7" y="4" width="6" height="1.5" rx=".75" fill="#c2540a"/><rect x="7" y="7" width="5" height="1.5" rx=".75" fill="#c2540a"/><rect x="7" y="10" width="4" height="1.5" rx=".75" fill="#c2540a"/></svg> Create Presentation`;

  createPptxBtn.addEventListener("click", () => {
    createPresentation();
  });

  function createPresentation() {
    // Disable button immediately so user knows something is happening
    createPptxBtn.disabled = true;
    createPptxBtn.textContent = "Generating…";

    // Wrap everything in try/catch — any sync error before write() would otherwise
    // silently swallow the failure and leave the button permanently disabled.
    let pres;
    try {
      pres = buildPresentation();
    } catch (err) {
      console.error("PPTX build failed:", err);
      alert("Failed to build presentation: " + err.message);
      createPptxBtn.disabled = false;
      createPptxBtn.innerHTML = PPTX_BTN_HTML;
      return;
    }

    const safeName = (sharedFileName || "presentation").replace(/\.xlsx$/i, "");
    const fileName = `${safeName}_presentation.pptx`;

    pres.write({ outputType: "blob" })
      .then(blob => {
        const url = URL.createObjectURL(blob);
        const a   = document.createElement("a");
        a.href     = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 5000);
      })
      .catch(err => {
        console.error("PPTX write failed:", err);
        alert("Failed to generate presentation: " + err.message);
      })
      .finally(() => {
        createPptxBtn.disabled = false;
        createPptxBtn.innerHTML = PPTX_BTN_HTML;
      });
  }

  function buildPresentation() {
    const rows  = filteredRows();
    const kpis  = computeKpis(rows);

    // ── Palette helpers ──────────────────────────────────────────────────────
    const SLIDE_W = 10, SLIDE_H = 5.63; // inches (16:9)
    const BG      = "F7F8FA";
    const WHITE   = "FFFFFF";
    const TEXT1   = "1F2328";
    const TEXT2   = "57606A";
    const BORDER  = "E5E7EB";
    const ACCENT  = "3B82D4";

    // Map our CSS hex colours to PptxGen hex (no #)
    const h = c => c.replace("#", "");

    // PptxGen colour constants (no #)
    const STATUS_C = {
      DONE: "22C55E", INPROG: "F59E0B", TODO: "3B82D4",
      BLOCKED: "EF4444", REJECTED: "7C5CD8", TESTING: "67E8F9", "WAIT-INFO": "1F2328",
    };
    const RESP_C = {
      "Resolved": "22C55E", "On OTE Team": "3B82D4",
      "On SHSO Team": "67E8F9", "Blocked": "EF4444", "Rejected": "7C5CD8",
    };
    const PAL = ["3B82D4","7C5CD8","22C55E","F59E0B","EF4444","06B6D4","EC4899","84CC16"];

    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_WIDE"; // 16:9

    // ── Shared slide helpers ─────────────────────────────────────────────────

    /** Add a standard slide with a coloured title bar and subtitle */
    function addSlide(titleText, subtitleText) {
      const slide = pres.addSlide();
      // Background
      slide.background = { color: BG };
      // Title bar (top strip)
      slide.addShape(pres.ShapeType.rect, {
        x: 0, y: 0, w: SLIDE_W, h: 0.55,
        fill: { color: WHITE }, line: { color: BORDER, width: 0.5 },
      });
      slide.addText(titleText, {
        x: 0.2, y: 0, w: 7, h: 0.55,
        fontSize: 14, bold: true, color: TEXT1, valign: "middle",
      });
      if (subtitleText) {
        slide.addText(subtitleText, {
          x: 7.2, y: 0, w: 2.6, h: 0.55,
          fontSize: 8, color: TEXT2, valign: "middle", align: "right",
        });
      }
      return slide;
    }

    /** Draw a rounded white card box */
    function addCard(slide, x, y, w, h) {
      slide.addShape(pres.ShapeType.rect, {
        x, y, w, h,
        fill: { color: WHITE },
        line: { color: BORDER, width: 0.75 },
        rectRadius: 0.06,
      });
    }

    /** Draw a section heading (small caps label style) */
    function addSectionLabel(slide, text, x, y, w) {
      slide.addText(text.toUpperCase(), {
        x, y, w, h: 0.22,
        fontSize: 7, bold: true, color: TEXT2, charSpacing: 1,
      });
    }

    /** Build pie chart data + legend entries from entries array [[label,count]…] */
    function buildPieData(entries, colorMap) {
      const total = entries.reduce((s, [, c]) => s + c, 0);
      const labels = entries.map(([l]) => l);
      const values = entries.map(([, c]) => c);
      const colors = entries.map(([l], i) => colorMap?.[l] ?? PAL[i % PAL.length]);
      return { labels, values, colors, total };
    }

    /** Draw a pie chart + legend on a slide.
     *  cx,cy = centre (inches), r = radius (inches) */
    function drawPie(slide, entries, cx, cy, r, colorMap) {
      const { labels, values, colors, total } = buildPieData(entries, colorMap);
      if (total === 0) return;

      const NS = "http://www.w3.org/2000/svg";
      const PX = 200; // SVG canvas px per inch equivalent
      const svgW = r * 2 * PX + 60;  // extra for labels
      const svgH = r * 2 * PX + 60;
      const scx  = svgW / 2;
      const scy  = svgH / 2;
      const sR   = r * PX;

      let paths = "";
      let angle = -Math.PI / 2;
      entries.forEach(([, count], i) => {
        const sweep = (count / total) * 2 * Math.PI;
        const x1 = scx + sR * Math.cos(angle);
        const y1 = scy + sR * Math.sin(angle);
        const x2 = scx + sR * Math.cos(angle + sweep);
        const y2 = scy + sR * Math.sin(angle + sweep);
        const large = sweep > Math.PI ? 1 : 0;
        const col = colors[i];
        paths += `<path d="M${scx},${scy} L${x1.toFixed(1)},${y1.toFixed(1)} A${sR},${sR} 0 ${large} 1 ${x2.toFixed(1)},${y2.toFixed(1)} Z" fill="#${col}" />`;

        // count label at midpoint
        const mid  = angle + sweep / 2;
        const lx   = scx + sR * 0.65 * Math.cos(mid);
        const ly   = scy + sR * 0.65 * Math.sin(mid);
        if (count > 0) {
          paths += `<text x="${lx.toFixed(1)}" y="${(ly + 5).toFixed(1)}" text-anchor="middle" font-size="14" font-weight="bold" fill="white" font-family="sans-serif">${count}</text>`;
        }
        angle += sweep;
      });

      // Donut hole
      paths += `<circle cx="${scx}" cy="${scy}" r="${sR * 0.33}" fill="white"/>`;
      paths += `<text x="${scx}" y="${scy + 5}" text-anchor="middle" font-size="16" font-weight="bold" fill="#${TEXT1}" font-family="sans-serif">${total}</text>`;

      const svgStr = `<svg xmlns="http://www.w3.org/2000/svg" width="${svgW}" height="${svgH}" viewBox="0 0 ${svgW} ${svgH}">${paths}</svg>`;
      const b64 = btoa(unescape(encodeURIComponent(svgStr)));

      slide.addImage({
        data: `data:image/svg+xml;base64,${b64}`,
        x: cx - r - 0.15,
        y: cy - r - 0.15,
        w: r * 2 + 0.3,
        h: r * 2 + 0.3,
      });
    }

    /** Draw a vertical legend to the right of a pie */
    function drawLegend(slide, entries, x, y, colorMap) {
      const ROW_H = 0.22;
      entries.forEach(([label, count], i) => {
        const col   = colorMap?.[label] ?? PAL[i % PAL.length];
        const total = entries.reduce((s, [, c]) => s + c, 0);
        const pct   = total > 0 ? ((count / total) * 100).toFixed(1) : "0.0";
        // swatch
        slide.addShape(pres.ShapeType.rect, {
          x, y: y + i * ROW_H, w: 0.12, h: 0.12,
          fill: { color: col }, line: { color: col, width: 0 },
        });
        // text
        slide.addText(`${label}  ${count} (${pct}%)`, {
          x: x + 0.16, y: y + i * ROW_H - 0.01, w: 2.5, h: ROW_H,
          fontSize: 7.5, color: TEXT1, valign: "middle",
        });
      });
    }

    /** Draw a simple horizontal bar chart */
    function drawBars(slide, entries, x, y, w, bh, barColor, title) {
      if (title) addSectionLabel(slide, title, x, y, w);
      const top    = y + (title ? 0.22 : 0);
      const bH     = Math.min(0.22, (bh - 0.22) / Math.max(entries.length, 1));
      const maxVal = Math.max(...entries.map(([, c]) => c), 1);
      const labelW = 1.1;
      const countW = 0.3;
      const trackW = w - labelW - countW - 0.05;

      entries.forEach(([label, count], i) => {
        const by   = top + i * (bH + 0.04);
        const barW = (count / maxVal) * trackW;
        // label
        slide.addText(label, {
          x, y: by, w: labelW, h: bH,
          fontSize: 7, color: TEXT1, valign: "middle",
          align: "right",
        });
        // track
        slide.addShape(pres.ShapeType.rect, {
          x: x + labelW + 0.04, y: by + bH * 0.15,
          w: trackW, h: bH * 0.7,
          fill: { color: "F0F0F0" }, line: { color: "F0F0F0", width: 0 },
        });
        // fill
        if (barW > 0) {
          slide.addShape(pres.ShapeType.rect, {
            x: x + labelW + 0.04, y: by + bH * 0.15,
            w: barW, h: bH * 0.7,
            fill: { color: h(barColor) }, line: { color: h(barColor), width: 0 },
          });
        }
        // count
        slide.addText(String(count), {
          x: x + labelW + 0.04 + trackW + 0.03, y: by, w: countW, h: bH,
          fontSize: 7, color: TEXT2, valign: "middle",
        });
      });
    }

    /** Draw a simple SVG line chart and embed as image */
    function drawLineChart(slide, checkpoints, labels, series, seriesColors, seriesNames, x, y, w, h) {
      const PX_W = 560, PX_H = 200;
      const PAD = { top: 14, right: 10, bottom: 44, left: 34 };
      const cW = PX_W - PAD.left - PAD.right;
      const cH = PX_H - PAD.top  - PAD.bottom;
      const n  = checkpoints.length;

      const allVals = series.flatMap(s => s);
      const maxV    = Math.max(...allVals, 1);
      const xOf = i => PAD.left + (n <= 1 ? cW / 2 : (i / (n - 1)) * cW);
      const yOf = v => PAD.top  + cH - (v / maxV) * cH;

      let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${PX_W}" height="${PX_H}" viewBox="0 0 ${PX_W} ${PX_H}">`;
      svg += `<rect width="${PX_W}" height="${PX_H}" fill="white"/>`;

      // grid lines
      for (let s = 0; s <= 4; s++) {
        const v = Math.round((maxV / 4) * s);
        const gy = yOf(v);
        svg += `<line x1="${PAD.left}" y1="${gy}" x2="${PAD.left + cW}" y2="${gy}" stroke="#E5E7EB" stroke-width="1"/>`;
        svg += `<text x="${PAD.left - 3}" y="${gy + 3}" text-anchor="end" font-size="8" fill="#57606A" font-family="sans-serif">${v}</text>`;
      }

      // x-axis
      svg += `<line x1="${PAD.left}" y1="${PAD.top + cH}" x2="${PAD.left + cW}" y2="${PAD.top + cH}" stroke="#D0D7DE" stroke-width="1"/>`;

      // x labels (max 10)
      const step = Math.ceil(n / 10);
      for (let i = 0; i < n; i++) {
        if (i % step !== 0 && i !== n - 1) continue;
        const lx = xOf(i);
        const short = labels[i] ? labels[i].substring(0, 8) : "";
        svg += `<text x="${lx}" y="${PAD.top + cH + 12}" text-anchor="middle" font-size="7" fill="#57606A" font-family="sans-serif" transform="rotate(-30,${lx},${PAD.top + cH + 12})">${short}</text>`;
      }

      // series lines + dots
      series.forEach((vals, si) => {
        const col = seriesColors[si];
        const pts = vals.map((v, i) => `${xOf(i).toFixed(1)},${yOf(v).toFixed(1)}`).join(" ");
        svg += `<polyline points="${pts}" fill="none" stroke="${col}" stroke-width="2"/>`;
        vals.forEach((v, i) => {
          svg += `<circle cx="${xOf(i).toFixed(1)}" cy="${yOf(v).toFixed(1)}" r="3" fill="${col}"/>`;
        });
      });

      svg += `</svg>`;
      const b64 = btoa(unescape(encodeURIComponent(svg)));
      slide.addImage({ data: `data:image/svg+xml;base64,${b64}`, x, y, w, h });
    }

    /** Draw an SVG bar chart and embed as image */
    function drawBarChartSvg(slide, series, labels, barColor, avgPerWeek, x, y, w, h) {
      const PX_W = 520, PX_H = 160;
      const PAD = { top: 14, right: 10, bottom: 44, left: 34 };
      const cW = PX_W - PAD.left - PAD.right;
      const cH = PX_H - PAD.top  - PAD.bottom;
      const n  = series.length;
      const maxV = Math.max(...series, 1);
      const xOf  = i => PAD.left + (n <= 1 ? cW / 2 : (i / (n - 1)) * cW);
      const yOf  = v => PAD.top + cH - (v / maxV) * cH;
      const barW = Math.max(4, Math.min(28, (cW / Math.max(n, 1)) * 0.6));

      let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${PX_W}" height="${PX_H}" viewBox="0 0 ${PX_W} ${PX_H}">`;
      svg += `<rect width="${PX_W}" height="${PX_H}" fill="white"/>`;

      for (let s = 0; s <= 4; s++) {
        const v  = Math.round((maxV / 4) * s);
        const gy = yOf(v);
        svg += `<line x1="${PAD.left}" y1="${gy}" x2="${PAD.left + cW}" y2="${gy}" stroke="#E5E7EB" stroke-width="1"/>`;
        svg += `<text x="${PAD.left - 3}" y="${gy + 3}" text-anchor="end" font-size="8" fill="#57606A" font-family="sans-serif">${v}</text>`;
      }
      svg += `<line x1="${PAD.left}" y1="${PAD.top + cH}" x2="${PAD.left + cW}" y2="${PAD.top + cH}" stroke="#D0D7DE" stroke-width="1"/>`;

      // avg dashed line
      const avgY = yOf(avgPerWeek);
      svg += `<line x1="${PAD.left}" y1="${avgY}" x2="${PAD.left + cW}" y2="${avgY}" stroke="#F59E0B" stroke-width="1.5" stroke-dasharray="5 3"/>`;

      // bars
      series.forEach((v, i) => {
        const bx  = xOf(i) - barW / 2;
        const bH2 = Math.max(0, cH - (yOf(v) - PAD.top));
        svg += `<rect x="${bx.toFixed(1)}" y="${yOf(v).toFixed(1)}" width="${barW}" height="${bH2.toFixed(1)}" rx="2" fill="${barColor}"/>`;
      });

      // x labels (max 10)
      const step = Math.ceil(n / 10);
      for (let i = 0; i < n; i++) {
        if (i % step !== 0 && i !== n - 1) continue;
        const lx = xOf(i);
        const short = labels[i] ? labels[i].substring(0, 8) : "";
        svg += `<text x="${lx}" y="${PAD.top + cH + 12}" text-anchor="middle" font-size="7" fill="#57606A" font-family="sans-serif" transform="rotate(-30,${lx},${PAD.top + cH + 12})">${short}</text>`;
      }
      svg += `</svg>`;
      const b64 = btoa(unescape(encodeURIComponent(svg)));
      slide.addImage({ data: `data:image/svg+xml;base64,${b64}`, x, y, w, h });
    }

    // ── Compute shared data ──────────────────────────────────────────────────

    // Status tally
    const statusEntries = tally(rows, 2).map(([raw, c]) => {
      const lbl = raw === "CLOSED" ? "DONE" : raw;
      return [lbl, c];
    });

    // Responsibility tally
    const respEntries = groupedTally(rows, respGroupLabel);

    // Vertical pies
    function vertPieEntries(typeFilterFn) {
      const OPEN = new Set(["TODO", "INPROG"]);
      const map  = new Map();
      for (const row of allRows.filter(typeFilterFn)) {
        const st   = normaliseStatus(row[2]);
        if (!OPEN.has(st)) continue;
        const vert = row.length > 15 && row[15] != null ? String(row[15]).trim() : "(blank)";
        map.set(vert, (map.get(vert) ?? 0) + 1);
      }
      return [...map.entries()].sort((a, b) => b[1] - a[1]);
    }
    const bugVertEntries = vertPieEntries(
      r => (r[5] != null ? String(r[5]).trim() : "") === "Defect (Bug)"
    );
    const enhVertEntries = vertPieEntries(
      r => (r[5] != null ? String(r[5]).trim() : "") === "Enhancement (SR)"
    );

    // Burnout data (recompute from aging dates)
    const fromVal = document.getElementById("agingFrom").value;
    const toVal   = document.getElementById("agingTo").value;
    const fromDate = fromVal ? new Date(fromVal + "T00:00:00Z") : null;
    const toDate   = toVal   ? new Date(toVal   + "T00:00:00Z") : null;
    const MS_DAY   = 86400000;

    function weekMonday2(d) {
      const dow  = d.getUTCDay();
      const diff = (dow === 0) ? -6 : 1 - dow;
      return new Date(d.getTime() + diff * MS_DAY);
    }
    function weekSunday2(d) {
      const dow  = d.getUTCDay();
      const diff = (dow === 0) ? 0 : 7 - dow;
      return new Date(d.getTime() + diff * MS_DAY);
    }
    function isoWeek2(d) {
      const thu = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      const dow = d.getUTCDay() || 7;
      thu.setUTCDate(thu.getUTCDate() + 4 - dow);
      const ys = new Date(Date.UTC(thu.getUTCFullYear(), 0, 1));
      return Math.ceil(((thu - ys) / MS_DAY + 1) / 7);
    }
    function fmtD(d) {
      return `${String(d.getUTCMonth() + 1).padStart(2, "0")}/${String(d.getUTCDate()).padStart(2, "0")}`;
    }

    let checkpoints = [], cpLabels = [];
    if (fromDate && toDate && fromDate < toDate) {
      let cur = weekSunday2(fromDate);
      while (cur <= toDate) {
        checkpoints.push(new Date(cur));
        const mon = weekMonday2(cur);
        cpLabels.push(`W${isoWeek2(cur)} ${fmtD(mon)}`);
        cur = new Date(cur.getTime() + 7 * MS_DAY);
      }
      if (!checkpoints.length || checkpoints[checkpoints.length - 1] < toDate) {
        const ls = weekSunday2(toDate);
        checkpoints.push(ls);
        cpLabels.push(`W${isoWeek2(ls)} ${fmtD(weekMonday2(ls))}`);
      }
    }

    const TODO_SET    = new Set(["TODO"]);
    const INPROG_SET  = new Set(["INPROG"]);
    const TESTING_SET = new Set(["TESTING"]);
    const DONE_SET    = new Set(["DONE", "CLOSED"]);

    const burnoutRows  = filteredRows();
    const todoS = [], inprogS = [], testingS = [], doneS = [];
    for (const cp of checkpoints) {
      let t = 0, ip = 0, te = 0, d = 0;
      for (const row of burnoutRows) {
        const st = row[2] != null ? String(row[2]).trim().toUpperCase() : "";
        const dt = parseDate(row[3]);
        if (!dt || dt > cp) continue;
        if (TODO_SET.has(st))        t++;
        else if (INPROG_SET.has(st)) ip++;
        else if (TESTING_SET.has(st)) te++;
        else if (DONE_SET.has(st))   d++;
      }
      todoS.push(t); inprogS.push(ip); testingS.push(te); doneS.push(d);
    }

    // Creation series
    function creationSeries(typeValue) {
      return checkpoints.map(cp => {
        const mon = weekMonday2(cp);
        let count = 0;
        for (const row of allRows) {
          const t = row[5] != null ? String(row[5]).trim() : "";
          if (t !== typeValue) continue;
          const d = parseDate(row[10]);
          if (d && d >= mon && d <= cp) count++;
        }
        return count;
      });
    }
    const bugCreation  = creationSeries("Defect (Bug)");
    const enhCreation  = creationSeries("Enhancement (SR)");
    const opCreation   = creationSeries("Operational");
    const totalDays    = (fromDate && toDate) ? Math.max(1, Math.round((toDate - fromDate) / MS_DAY) + 1) : 1;
    const bugAvg       = bugCreation.reduce((a, b) => a + b, 0) / totalDays;
    const enhAvg       = enhCreation.reduce((a, b) => a + b, 0) / totalDays;
    const opAvg        = opCreation.reduce((a, b) => a + b, 0) / totalDays;

    // ── SLIDE 1: KPI Tiles ───────────────────────────────────────────────────
    {
      const slide = addSlide("Statistics — KPI Summary", sharedFileName);
      const accentColors = {
        "kpi-accent-blue":   ACCENT,
        "kpi-accent-green":  "22C55E",
        "kpi-accent-amber":  "F59E0B",
        "kpi-accent-red":    "EF4444",
        "kpi-accent-purple": "7C5CD8",
      };

      const tiles = [
        { value: String(kpis.total),                          label: "Total Tickets",       sub: null,                              accent: "3B82D4" },
        { value: String(kpis.resolved),                       label: "Resolved",             sub: null,                              accent: "22C55E" },
        { value: kpis.resolutionRate.toFixed(1) + "%",        label: "Resolution Rate",      sub: `${kpis.resolved} of ${kpis.total}`, accent: "22C55E" },
        { value: String(kpis.throughputLast30),               label: "Resolved (Last 30 d)", sub: "by status date",                  accent: "F59E0B" },
        { value: String(kpis.openOver30),                     label: "Open > 30 Days",       sub: "by created date",                 accent: kpis.openOver30 > 0 ? "EF4444" : "22C55E" },
        { value: kpis.typeEntries[0]?.[0] ?? "—",             label: "Type Breakdown",       sub: `${kpis.typeEntries.length} types · top: ${kpis.typeEntries[0]?.[1] ?? 0}`, accent: "3B82D4" },
      ];

      const COLS = 3;
      const cardW = 2.8, cardH = 1.1;
      const gapX = 0.25, gapY = 0.35;
      const startX = (SLIDE_W - COLS * cardW - (COLS - 1) * gapX) / 2;
      const startY = 0.75;

      tiles.forEach((t, i) => {
        const col  = i % COLS;
        const row2 = Math.floor(i / COLS);
        const cx   = startX + col * (cardW + gapX);
        const cy   = startY + row2 * (cardH + gapY);

        addCard(slide, cx, cy, cardW, cardH);
        // accent top bar
        slide.addShape(pres.ShapeType.rect, {
          x: cx, y: cy, w: cardW, h: 0.06,
          fill: { color: t.accent }, line: { color: t.accent, width: 0 },
        });
        // value
        slide.addText(t.value, {
          x: cx + 0.15, y: cy + 0.1, w: cardW - 0.3, h: 0.5,
          fontSize: 24, bold: true, color: TEXT1,
        });
        // label
        slide.addText(t.label.toUpperCase(), {
          x: cx + 0.15, y: cy + 0.58, w: cardW - 0.3, h: 0.22,
          fontSize: 7, bold: true, color: TEXT2, charSpacing: 0.5,
        });
        // sub
        if (t.sub) {
          slide.addText(t.sub, {
            x: cx + 0.15, y: cy + 0.79, w: cardW - 0.3, h: 0.2,
            fontSize: 7, color: TEXT2,
          });
        }
      });

      // Type breakdown mini-bar below tiles
      if (kpis.typeEntries.length) {
        const bx = 0.5, by = 3.3, bw = 9.0, bh = 1.9;
        addCard(slide, bx, by, bw, bh);
        addSectionLabel(slide, "Ticket Type Breakdown", bx + 0.15, by + 0.12, bw - 0.3);
        drawBars(slide, kpis.typeEntries.slice(0, 7), bx + 0.15, by + 0.12, bw - 0.3, bh - 0.15, "#3b82d4", null);
      }
    }

    // ── SLIDE 2: Status Pie + Responsibility Pie ──────────────────────────────
    {
      const slide = addSlide("Statistics — Ticket Breakdown", sharedFileName);
      const cardH = 4.6, cardW = 4.55;

      // ── Status card (left) ──────────────────────────────────────────────────
      addCard(slide, 0.15, 0.65, cardW, cardH);
      addSectionLabel(slide, "Ticket Status Breakdown", 0.3, 0.72, cardW - 0.3);

      drawPie(slide, statusEntries, 0.15 + 1.15, 0.65 + 1.2 + 0.8, 1.1,
        Object.fromEntries(statusEntries.map(([l]) => [l, STATUS_C[l] ?? "C0C8D0"]))
      );
      drawLegend(slide, statusEntries, 0.15 + 2.55, 0.65 + 0.95 + 0.25,
        Object.fromEntries(statusEntries.map(([l]) => [l, STATUS_C[l] ?? "C0C8D0"]))
      );

      // ── Responsibility card (right) ─────────────────────────────────────────
      addCard(slide, 5.3, 0.65, cardW, cardH);
      addSectionLabel(slide, "Responsibility Breakdown", 5.45, 0.72, cardW - 0.3);

      drawPie(slide, respEntries, 5.3 + 1.15, 0.65 + 1.2 + 0.8, 1.1, RESP_C);
      drawLegend(slide, respEntries, 5.3 + 2.55, 0.65 + 0.95 + 0.25, RESP_C);
    }

    // ── SLIDE 3: Vertical Pies ────────────────────────────────────────────────
    {
      const slide = addSlide("Statistics — Open Tickets by Vertical", sharedFileName);
      const cardW = 4.55, cardH = 4.6;

      // ── Bugs per vertical (left) ─────────────────────────────────────────────
      addCard(slide, 0.15, 0.65, cardW, cardH);
      addSectionLabel(slide, "Open Bugs by Vertical (TODO + In Progress)", 0.3, 0.72, cardW - 0.3);

      if (bugVertEntries.length) {
        const bugColorMap = Object.fromEntries(bugVertEntries.map(([l], i) => [l, PAL[i % PAL.length]]));
        drawPie(slide, bugVertEntries, 0.15 + 1.1, 0.65 + 1.1 + 0.9, 1.05, bugColorMap);
        drawLegend(slide, bugVertEntries, 0.15 + 2.45, 0.65 + 0.95 + 0.25, bugColorMap);
      } else {
        slide.addText("No open bugs in current data", {
          x: 0.3, y: 1.5, w: cardW - 0.3, h: 0.4, fontSize: 9, color: TEXT2,
        });
      }

      // ── Enhancements per vertical (right) ────────────────────────────────────
      addCard(slide, 5.3, 0.65, cardW, cardH);
      addSectionLabel(slide, "Open Enhancements by Vertical (TODO + In Progress)", 5.45, 0.72, cardW - 0.3);

      if (enhVertEntries.length) {
        const enhColorMap = Object.fromEntries(enhVertEntries.map(([l], i) => [l, PAL[i % PAL.length]]));
        drawPie(slide, enhVertEntries, 5.3 + 1.1, 0.65 + 1.1 + 0.9, 1.05, enhColorMap);
        drawLegend(slide, enhVertEntries, 5.3 + 2.45, 0.65 + 0.95 + 0.25, enhColorMap);
      } else {
        slide.addText("No open enhancements in current data", {
          x: 5.45, y: 1.5, w: cardW - 0.3, h: 0.4, fontSize: 9, color: TEXT2,
        });
      }
    }

    // ── SLIDE 4: Burnout + Creation Charts ────────────────────────────────────
    {
      const slide = addSlide("Statistics — Burnout & Creation Trends", sharedFileName);
      const dateRange = (fromVal && toVal) ? `${fromVal} → ${toVal}` : "All dates";

      // Burnout line chart (full width, top)
      addCard(slide, 0.15, 0.65, 9.7, 1.95);
      addSectionLabel(slide, `Burnout  ·  ${dateRange}`, 0.3, 0.72, 6);
      if (checkpoints.length > 0) {
        drawLineChart(
          slide, checkpoints, cpLabels,
          [todoS, inprogS, testingS, doneS],
          ["#3B82D4", "#F59E0B", "#7C5CD8", "#22C55E"],
          ["TODO", "In Progress", "Testing", "Done"],
          0.2, 0.88, 9.6, 1.65
        );
        // Mini legend
        const legendItems = [
          { label: "TODO",        color: "3B82D4" },
          { label: "In Progress", color: "F59E0B" },
          { label: "Testing",     color: "7C5CD8" },
          { label: "Done",        color: "22C55E" },
        ];
        legendItems.forEach((li, i) => {
          slide.addShape(pres.ShapeType.rect, {
            x: 0.3 + i * 1.5, y: 2.43, w: 0.22, h: 0.05,
            fill: { color: li.color }, line: { color: li.color, width: 0 },
          });
          slide.addText(li.label, {
            x: 0.55 + i * 1.5, y: 2.39, w: 1.2, h: 0.16,
            fontSize: 7, color: TEXT2,
          });
        });
      } else {
        slide.addText("No date range selected — run Burnout with dates to include this chart.", {
          x: 0.3, y: 1.3, w: 9.4, h: 0.4, fontSize: 9, color: TEXT2,
        });
      }

      // Three creation charts side by side (bottom row)
      const creationDefs = [
        { label: "Bugs Created / Week",          series: bugCreation, avg: bugAvg, color: "#EF4444" },
        { label: "Enhancements Created / Week",  series: enhCreation, avg: enhAvg, color: "#7C5CD8" },
        { label: "Operational Created / Week",   series: opCreation,  avg: opAvg,  color: "#F59E0B" },
      ];
      const chartW = 3.05, chartH = 2.05, chartY = 2.75;
      creationDefs.forEach((def, i) => {
        const cx = 0.15 + i * (chartW + 0.2);
        addCard(slide, cx, chartY, chartW, chartH);
        // stat numbers
        const total = def.series.reduce((a, b) => a + b, 0);
        const avgWk = def.series.length ? (total / def.series.length).toFixed(1) : "—";
        slide.addText(def.label.toUpperCase(), {
          x: cx + 0.12, y: chartY + 0.1, w: chartW - 0.2, h: 0.18,
          fontSize: 6.5, bold: true, color: TEXT2, charSpacing: 0.5,
        });
        slide.addText(`${total} total  ·  ${avgWk}/wk avg`, {
          x: cx + 0.12, y: chartY + 0.28, w: chartW - 0.2, h: 0.18,
          fontSize: 7.5, color: TEXT1,
        });
        if (checkpoints.length > 0) {
          const avgPW = def.series.length ? total / def.series.length : 0;
          drawBarChartSvg(slide, def.series, cpLabels, def.color, avgPW,
            cx + 0.05, chartY + 0.48, chartW - 0.1, chartH - 0.55);
        }
      });
    }

    // Return the assembled presentation object for the caller to save
    return pres;
  }

});
