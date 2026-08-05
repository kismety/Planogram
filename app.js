
let DATA = [];
let currentCol = "";
let fileName = "";
let currentStore = "";

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, m => ({
  "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"
}[m]));

function show(screen) {
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("show"));
  document.querySelectorAll("nav button").forEach(b => b.classList.remove("active"));
  $(screen).classList.add("show");
  $("nav" + screen.charAt(0).toUpperCase() + screen.slice(1)).classList.add("active");

  const onReset = screen === "reset";
  $("resetTools").style.display = onReset ? "block" : "none";
  $("footer").classList.toggle("show", onReset);

  if (onReset) renderReset();
  if (screen === "database") renderDatabase();
renderSavedPlans();
}

$("navUpload").addEventListener("click", () => show("upload"));
$("navReset").addEventListener("click", () => show("reset"));
$("navDatabase").addEventListener("click", () => show("database"));
$("navPlans").addEventListener("click", () => { show("plans"); renderSavedPlans(); });

function hkey(v) {
  return String(v ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function firstIndex(headers, names) {
  for (const name of names) {
    const i = headers.indexOf(name);
    if (i >= 0) return i;
  }
  return -1;
}

function parseWorkbook(arrayBuffer) {
  if (typeof XLSX === "undefined") throw new Error("Excel reader did not load.");

  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const priority = ["Reset Database", "Extracted Positions", "Import Template"];
  let sheetName = priority.find(n => wb.SheetNames.includes(n)) || wb.SheetNames[0];
  if (!sheetName) throw new Error("Workbook has no worksheets.");

  const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: false
  });

  const headerIndex = matrix.findIndex(row => row.some(v => hkey(v) === "cardname"));
  if (headerIndex < 0) throw new Error("Could not find a Card Name column.");

  const headers = matrix[headerIndex].map(hkey);
  const idx = {
    card: firstIndex(headers, ["cardname","card","productcardname","product"]),
    denom: firstIndex(headers, ["denomination","denom"]),
    category: firstIndex(headers, ["category"]),
    column: firstIndex(headers, ["column","col"]),
    row: firstIndex(headers, ["row"]),
    position: firstIndex(headers, ["position"]),
    planogram: firstIndex(headers, ["planogram"]),
    notes: firstIndex(headers, ["notes"]),
    confidence: firstIndex(headers, ["confidence"])
  };

  if (idx.card < 0 || idx.column < 0) {
    throw new Error("Workbook needs Card Name and Column.");
  }

  const out = [];
  matrix.slice(headerIndex + 1).forEach((r, n) => {
    const card = String(r[idx.card] ?? "").trim();
    if (!card) return;

    const column = idx.column >= 0 ? String(r[idx.column] ?? "").trim() : "";
    const row = idx.row >= 0 ? String(r[idx.row] ?? "").trim() : "";
    let position = idx.position >= 0 ? String(r[idx.position] ?? "").trim() : "";
    if (!position && column) position = column + (row ? "-" + row : "");

    out.push({
      id: `${n}|${card}|${position}`,
      planogram: idx.planogram >= 0 ? String(r[idx.planogram] ?? "").trim() : sheetName,
      card,
      denomination: idx.denom >= 0 ? String(r[idx.denom] ?? "").trim() : "",
      category: idx.category >= 0 ? String(r[idx.category] ?? "").trim() : "",
      column,
      row,
      position,
      notes: idx.notes >= 0 ? String(r[idx.notes] ?? "").trim() : "",
      confidence: idx.confidence >= 0 ? String(r[idx.confidence] ?? "").trim() : ""
    });
  });

  if (!out.length) throw new Error("No planogram rows were found.");
  return { rows: out, sheetName };
}

function statusKey(r) {
  return `pgstatus|${r.planogram || fileName}|${r.position}|${r.card}`;
}
function getStatus(r) {
  return localStorage.getItem(statusKey(r)) || "Not Complete";
}
function setStatus(r, status) {
  localStorage.setItem(statusKey(r), status);
  saveAll(false);
}

function saveImported() {
  localStorage.setItem("planogram_imported_rows", JSON.stringify(DATA));
  localStorage.setItem("planogram_imported_filename", fileName);
}


function saveAll(showMessage=true) {
  try {
    saveImported();
    DATA.forEach(r => {
      const s = getStatus(r);
      localStorage.setItem(statusKey(r), s);
    });
    localStorage.setItem("planogram_saved_at", new Date().toISOString());
    if (showMessage && $("saveStatus")) {
      $("saveStatus").textContent = "✓ Saved on this device at " + new Date().toLocaleTimeString([], {hour:"2-digit", minute:"2-digit"});
      $("saveStatus").classList.add("saved");
    }
    try {
      const pname = $("planogramName") ? $("planogramName").value.trim() : "";
      const sname = $("storeName") ? $("storeName").value.trim() : "";
      if (pname) {
        const match = savedPlansIndex().find(p => p.name.toLowerCase() === pname.toLowerCase() && (!sname || (p.store||"").toLowerCase()===sname.toLowerCase()));
        if (match) updateSavedPlanogram(match.id);
      }
    } catch (_) {}
    return true;
  } catch (e) {
    if ($("saveStatus")) {
      $("saveStatus").textContent = "Could not save: " + e.message;
      $("saveStatus").classList.remove("saved");
    }
    return false;
  }
}

function makeBackupObject() {
  const statuses = {};
  DATA.forEach(r => { statuses[statusKey(r)] = getStatus(r); });
  return {
    version: 1,
    exportedAt: new Date().toISOString(),
    fileName,
    currentStore,
    currentCol,
    currentStore: store,
    data: DATA,
    statuses
  };
}

function downloadBackup() {
  if (!DATA.length) {
    alert("Upload a planogram first.");
    return;
  }
  saveAll(false);
  const blob = new Blob([JSON.stringify(makeBackupObject(), null, 2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const base = (fileName || "planogram").replace(/\.[^.]+$/, "").replace(/[^a-z0-9_-]+/gi, "_");
  a.download = base + "_progress_backup.json";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  if ($("saveStatus")) {
    $("saveStatus").textContent = "✓ Backup downloaded.";
    $("saveStatus").classList.add("saved");
  }
}

async function restoreBackupFromFile(file) {
  const txt = await file.text();
  const backup = JSON.parse(txt);
  if (!backup || !Array.isArray(backup.data)) throw new Error("This is not a valid Planogram backup.");
  DATA = backup.data;
  fileName = backup.fileName || "Restored planogram";
  currentStore = backup.currentStore || "";
  currentCol = backup.currentCol || "";
  if ($("storeName")) $("storeName").value = currentStore;
  localStorage.setItem("planogram_imported_rows", JSON.stringify(DATA));
  localStorage.setItem("planogram_imported_filename", fileName);
  if (backup.statuses && typeof backup.statuses === "object") {
    Object.entries(backup.statuses).forEach(([k,v]) => localStorage.setItem(k,v));
  }
  afterLoad();
  saveAll(false);
}


function savedPlansIndex() {
  try {
    return JSON.parse(localStorage.getItem("saved_planograms_index") || "[]");
  } catch (_) {
    return [];
  }
}

function writeSavedPlansIndex(list) {
  localStorage.setItem("saved_planograms_index", JSON.stringify(list));
}

function safePlanId(name) {
  return "plan_" + name.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") + "_" + Date.now();
}

function collectStatusesForData(rows) {
  const statuses = {};
  rows.forEach(r => { statuses[statusKey(r)] = getStatus(r); });
  return statuses;
}


function exportPlanogramPNG() {
  if (!DATA.length) {
    alert("Upload a planogram first.");
    return;
  }

  const store = ($("storeName")?.value || currentStore || "Store").trim();
  const planName = ($("planogramName")?.value || "Planogram").trim();
  const cols = columns();
  if (!cols.length) {
    alert("No columns found in this planogram.");
    return;
  }

  const grouped = {};
  cols.forEach(c => grouped[c] = DATA.filter(r => r.column === c).sort((a,b) => rowNumber(a.row)-rowNumber(b.row)));

  const colWidth = 300;
  const gap = 18;
  const margin = 40;
  const headerH = 120;
  const rowH = 72;
  const maxRows = Math.max(...cols.map(c => grouped[c].length), 1);

  const canvas = $("exportCanvas");
  const ctx = canvas.getContext("2d");
  canvas.width = margin*2 + cols.length*colWidth + (cols.length-1)*gap;
  canvas.height = headerH + margin + maxRows*rowH + 80;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0,0,canvas.width,canvas.height);

  ctx.fillStyle = "#111111";
  ctx.font = "bold 32px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
  ctx.fillText(store, margin, 44);
  ctx.font = "bold 24px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
  ctx.fillText(planName, margin, 78);
  ctx.font = "16px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
  ctx.fillStyle = "#666666";
  ctx.fillText("Generated from Planogram Reset App", margin, 104);

  cols.forEach((c, ci) => {
    const x = margin + ci*(colWidth+gap);
    ctx.fillStyle = "#111111";
    ctx.fillRect(x, headerH, colWidth, 46);
    ctx.fillStyle = "#ffffff";
    ctx.font = "bold 22px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
    ctx.fillText(c, x+14, headerH+30);

    grouped[c].forEach((r, ri) => {
      const y = headerH + 46 + ri*rowH;
      ctx.fillStyle = ri%2===0 ? "#f7f7f7" : "#ffffff";
      ctx.fillRect(x, y, colWidth, rowH);
      ctx.strokeStyle = "#dddddd";
      ctx.strokeRect(x, y, colWidth, rowH);

      ctx.fillStyle = "#111111";
      ctx.font = "bold 15px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
      ctx.fillText(r.position || (c + "-" + r.row), x+10, y+20);

      ctx.font = "14px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
      const name = r.card || "";
      const maxChars = 34;
      const first = name.length > maxChars ? name.slice(0,maxChars-1)+"…" : name;
      ctx.fillText(first, x+10, y+42);

      if (r.denomination) {
        ctx.fillStyle = "#666666";
        ctx.font = "12px -apple-system, BlinkMacSystemFont, Segoe UI, Arial";
        ctx.fillText(r.denomination, x+10, y+60);
      }
    });
  });

  const link = document.createElement("a");
  const safeStore = store.replace(/[^a-z0-9_-]+/gi,"_");
  const safePlan = planName.replace(/[^a-z0-9_-]+/gi,"_");
  link.download = safeStore + "_" + safePlan + ".png";
  link.href = canvas.toDataURL("image/png");
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function saveAsNamedPlanogram() {
  if (!DATA.length) {
    alert("Upload an Excel planogram first.");
    return;
  }
  const name = $("planogramName").value.trim();
  const store = $("storeName").value.trim();
  currentStore = store;
  if (!store) {
    $("uploadStatus").textContent = "Enter the store name first.";
    return;
  }
  if (!name) {
    $("uploadStatus").textContent = "Enter a name for the planogram first.";
    return;
  }

  const index = savedPlansIndex();
  const existing = index.find(p => p.name.toLowerCase() === name.toLowerCase() && (p.store||"").toLowerCase() === store.toLowerCase());
  const id = existing ? existing.id : safePlanId(store + "_" + name);

  const record = {
    id,
    name,
    store,
    fileName,
    savedAt: new Date().toISOString(),
    currentCol,
    currentStore: store,
    data: DATA,
    statuses: collectStatusesForData(DATA)
  };

  localStorage.setItem("saved_planogram|" + id, JSON.stringify(record));

  if (existing) {
    existing.savedAt = record.savedAt;
    existing.store = store;
    existing.count = DATA.length;
    existing.fileName = fileName;
  } else {
    index.push({
      id,
      name,
      store,
      savedAt: record.savedAt,
      count: DATA.length,
      fileName
    });
  }

  index.sort((a,b) => new Date(b.savedAt) - new Date(a.savedAt));
  writeSavedPlansIndex(index);

  $("uploadStatus").innerHTML = "✓ Saved as planogram <b>" + esc(name) + "</b>.";
  renderSavedPlans();
}

function openSavedPlanogram(id) {
  try {
    const raw = localStorage.getItem("saved_planogram|" + id);
    if (!raw) throw new Error("Saved planogram not found.");
    const record = JSON.parse(raw);

    DATA = record.data || [];
    fileName = record.fileName || record.name || "Saved planogram";
    currentStore = record.store || "";
    if ($("planogramName")) $("planogramName").value = record.name || "";
    if ($("storeName")) $("storeName").value = record.store || "";
    currentCol = record.currentCol || "";

    if (record.statuses) {
      Object.entries(record.statuses).forEach(([k,v]) => localStorage.setItem(k,v));
    }

    saveImported();
    afterLoad();
    show("reset");
  } catch (e) {
    alert("Could not open saved planogram: " + e.message);
  }
}

function deleteSavedPlanogram(id) {
  const index = savedPlansIndex();
  const p = index.find(x => x.id === id);
  if (!p) return;
  if (!confirm('Delete saved planogram "' + p.name + '"?')) return;

  localStorage.removeItem("saved_planogram|" + id);
  writeSavedPlansIndex(index.filter(x => x.id !== id));
  renderSavedPlans();
}

function updateSavedPlanogram(id) {
  const raw = localStorage.getItem("saved_planogram|" + id);
  if (!raw) return;
  const record = JSON.parse(raw);
  record.data = DATA;
  record.fileName = fileName;
  record.currentCol = currentCol;
  record.statuses = collectStatusesForData(DATA);
  record.savedAt = new Date().toISOString();
  localStorage.setItem("saved_planogram|" + id, JSON.stringify(record));

  const index = savedPlansIndex();
  const item = index.find(x => x.id === id);
  if (item) {
    item.savedAt = record.savedAt;
    item.count = DATA.length;
    writeSavedPlansIndex(index);
  }
  renderSavedPlans();
}

function renderSavedPlans() {
  if (!$("plansList")) return;
  const index = savedPlansIndex();
  if (!index.length) {
    $("plansList").innerHTML = '<div class="card">No saved planograms yet. Upload Excel, give it a name, then tap <b>Save Planogram</b>.</div>';
    return;
  }

  $("plansList").innerHTML = index.map(p => `
    <div class="card">
      <div class="planrow">
        <div>
          <div class="dbname">${esc(p.name)}</div>
          <div class="savedstore">${esc(p.store || "")}</div>
          <div class="planmeta">${p.count || 0} positions${p.fileName ? " • " + esc(p.fileName) : ""}</div>
          <div class="planmeta">Saved ${new Date(p.savedAt).toLocaleString()}</div>
        </div>
      </div>
      <div class="planbuttons">
        <button class="primary" data-open-plan="${esc(p.id)}">Open</button>
        <button data-delete-plan="${esc(p.id)}">Delete</button>
      </div>
    </div>`).join("");

  document.querySelectorAll("[data-open-plan]").forEach(b =>
    b.addEventListener("click", () => openSavedPlanogram(b.dataset.openPlan))
  );
  document.querySelectorAll("[data-delete-plan]").forEach(b =>
    b.addEventListener("click", () => deleteSavedPlanogram(b.dataset.deletePlan))
  );
}

function loadSaved() {
  try {
    const saved = localStorage.getItem("planogram_imported_rows");
    if (saved) {
      DATA = JSON.parse(saved);
      fileName = localStorage.getItem("planogram_imported_filename") || "Saved planogram";
      afterLoad();
    }
  } catch (_) {}
}

$("saveAsPlanogram").addEventListener("click", saveAsNamedPlanogram);
$("exportPng").addEventListener("click", exportPlanogramPNG);

$("xlsxFile").addEventListener("change", async e => {
  const f = e.target.files?.[0];
  if (!f) return;
  if ($("planogramName") && !$("planogramName").value.trim()) $("planogramName").value = f.name.replace(/\.[^.]+$/, "");
  if ($("storeName")) currentStore = $("storeName").value.trim();

  $("uploadStatus").textContent = `Reading ${f.name}...`;

  try {
    const buf = await f.arrayBuffer();
    const parsed = parseWorkbook(buf);
    DATA = parsed.rows;
    fileName = f.name;
    saveImported();
    saveAll(false);
    afterLoad();
    $("uploadStatus").innerHTML = `✓ Loaded <b>${DATA.length}</b> positions from <b>${esc(parsed.sheetName)}</b>.`;
    show("reset");
  } catch (err) {
    console.error(err);
    $("uploadStatus").textContent = "Import failed: " + err.message;
  }
});

function columnNumber(c) {
  const n = parseInt(String(c).replace(/\D/g, ""), 10);
  return Number.isFinite(n) ? n : 9999;
}
function rowNumber(r) {
  const n = parseInt(String(r).replace(/\D/g, ""), 10);
  return Number.isFinite(n) ? n : 9999;
}
function columns() {
  return [...new Set(DATA.map(r => r.column).filter(Boolean))]
    .sort((a,b) => columnNumber(a) - columnNumber(b) || a.localeCompare(b));
}

function afterLoad() {
  $("fileLabel").textContent = `${fileName} • ${DATA.length} positions`;
  const cols = columns();
  if (!currentCol || !cols.includes(currentCol)) currentCol = cols[0] || "";

  $("colFilter").innerHTML =
    '<option value="">All columns</option>' +
    cols.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join("");

  renderTabs();
  renderReset();
  renderDatabase();
}

function renderTabs() {
  const cols = columns();
  $("tabs").innerHTML = "";
  cols.forEach(c => {
    const b = document.createElement("button");
    b.textContent = c;
    if (c === currentCol) b.className = "active";
    b.addEventListener("click", () => {
      currentCol = c;
      $("colFilter").value = c;
      $("search").value = "";
      $("statusFilter").value = "";
      renderReset();
      renderTabs();
    });
    $("tabs").appendChild(b);
  });
}

function currentRows() {
  const q = $("search").value.trim().toLowerCase();
  const cf = $("colFilter").value;
  const sf = $("statusFilter").value;

  let rows = DATA.filter(r =>
    (!q || r.card.toLowerCase().includes(q)) &&
    (!cf || r.column === cf) &&
    (!sf || getStatus(r) === sf)
  );

  if (!q && !cf && !sf && currentCol) {
    rows = DATA.filter(r => r.column === currentCol);
  }

  return rows.sort((a,b) =>
    columnNumber(a.column) - columnNumber(b.column) ||
    rowNumber(a.row) - rowNumber(b.row)
  );
}

function cardHtml(r) {
  const status = getStatus(r);
  return `
  <div class="card">
    <div class="line">
      <div class="pos">${esc(r.position || r.column)}</div>
      <div class="name">
        ${esc(r.card)}
        ${r.denomination ? `<div class="meta">Denomination: ${esc(r.denomination)}</div>` : ""}
        ${r.category ? `<div class="meta">Category: ${esc(r.category)}</div>` : ""}
      </div>
    </div>
    <div class="actions">
      <button class="complete ${status==="Complete"?"active":""}" data-id="${esc(r.id)}" data-status="Complete">✓ Complete</button>
      <button class="notcomplete ${status==="Not Complete"?"active":""}" data-id="${esc(r.id)}" data-status="Not Complete">Not Complete</button>
      <button class="missing ${status==="Missing"?"active":""}" data-id="${esc(r.id)}" data-status="Missing">Missing</button>
    </div>
  </div>`;
}

function renderReset() {
  if (!DATA.length) {
    $("resetList").innerHTML = '<div class="card">No planogram loaded. Go to <b>Upload Excel</b>.</div>';
    updateSummary();
    return;
  }

  const rows = currentRows();
  $("resetList").innerHTML = rows.length
    ? rows.map(cardHtml).join("")
    : '<div class="card">No matching positions.</div>';

  document.querySelectorAll("[data-status]").forEach(b => {
    b.addEventListener("click", () => {
      const r = DATA.find(x => x.id === b.dataset.id);
      if (!r) return;
      setStatus(r, b.dataset.status);
      renderReset();
    });
  });

  updateSummary();
}

function updateSummary() {
  const complete = DATA.filter(r => getStatus(r) === "Complete").length;
  const missing = DATA.filter(r => getStatus(r) === "Missing").length;
  const remaining = DATA.length - complete - missing;
  const pct = DATA.length ? Math.round(complete / DATA.length * 100) : 0;

  $("mComplete").textContent = complete;
  $("mRemaining").textContent = remaining;
  $("mMissing").textContent = missing;
  $("fill").style.width = pct + "%";
  $("progressText").textContent = `${pct}% complete • ${complete} of ${DATA.length} positions`;
}

$("search").addEventListener("input", renderReset);
$("statusFilter").addEventListener("change", renderReset);
$("colFilter").addEventListener("change", function() {
  if (this.value) currentCol = this.value;
  renderReset();
  renderTabs();
});

function renderDatabase() {
  const q = $("dbSearch").value.trim().toLowerCase();

  if (!DATA.length) {
    $("dbCount").textContent = "0 imported positions";
    $("dbList").innerHTML = '<div class="card">Upload an Excel planogram first.</div>';
    return;
  }

  const products = new Map();
  DATA.forEach(r => {
    const key = r.card.toLowerCase();
    if (!products.has(key)) {
      products.set(key, { name: r.card, category: r.category, positions: [] });
    }
    products.get(key).positions.push(r.position);
  });

  const list = [...products.values()]
    .filter(p => !q || p.name.toLowerCase().includes(q))
    .sort((a,b) => a.name.localeCompare(b.name));

  $("dbCount").textContent = `${list.length} unique products • ${DATA.length} total positions`;
  $("dbList").innerHTML = list.map(p => `
    <div class="card">
      <div class="dbname">${esc(p.name)}</div>
      ${p.category ? `<span class="pill">${esc(p.category)}</span>` : ""}
      <div class="meta">Positions: ${esc(p.positions.join(", "))}</div>
    </div>`).join("");
}

$("dbSearch").addEventListener("input", renderDatabase);

$("prev").addEventListener("click", () => {
  const cols = columns();
  const i = cols.indexOf(currentCol);
  if (i > 0) {
    currentCol = cols[i-1];
    $("colFilter").value = currentCol;
    $("search").value = "";
    $("statusFilter").value = "";
    renderTabs();
    renderReset();
    window.scrollTo(0,0);
  }
});

$("next").addEventListener("click", () => {
  const cols = columns();
  const i = cols.indexOf(currentCol);
  if (i >= 0 && i < cols.length - 1) {
    currentCol = cols[i+1];
    $("colFilter").value = currentCol;
    $("search").value = "";
    $("statusFilter").value = "";
    renderTabs();
    renderReset();
    window.scrollTo(0,0);
  }
});


$("saveNow").addEventListener("click", () => {
  if (!DATA.length) {
    $("saveStatus").textContent = "Upload a planogram first.";
    return;
  }
  saveAll(true);
});

$("downloadBackup").addEventListener("click", downloadBackup);

$("restoreBackup").addEventListener("click", async () => {
  const f = $("backupFile").files?.[0];
  if (!f) {
    $("saveStatus").textContent = "Choose a backup JSON file first.";
    return;
  }
  try {
    await restoreBackupFromFile(f);
    $("saveStatus").textContent = "✓ Backup restored.";
    $("saveStatus").classList.add("saved");
    show("reset");
  } catch (e) {
    $("saveStatus").textContent = "Restore failed: " + e.message;
    $("saveStatus").classList.remove("saved");
  }
});

window.addEventListener("pagehide", () => {
  if (DATA.length) saveAll(false);
});

$("clearData").addEventListener("click", () => {
  if (!confirm("Clear the imported planogram and saved progress?")) return;
  Object.keys(localStorage)
    .filter(k => k.startsWith("pgstatus|") || k.startsWith("planogram_imported_"))
    .forEach(k => localStorage.removeItem(k));

  DATA = [];
  fileName = "";
  currentCol = "";
  $("fileLabel").textContent = "Upload your converted Excel planogram to begin.";
  $("xlsxFile").value = "";
  $("uploadStatus").textContent = "Saved planogram and progress cleared.";
  renderReset();
  renderDatabase();
});

loadSaved();
try {
  const savedAt = localStorage.getItem("planogram_saved_at");
  if (savedAt && $("saveStatus")) {
    $("saveStatus").textContent = "✓ Saved progress found on this device.";
    $("saveStatus").classList.add("saved");
  }
} catch (_) {}
renderReset();
renderDatabase();
