
let DATA = [];
let currentCol = "";
let fileName = "";

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
}

$("navUpload").addEventListener("click", () => show("upload"));
$("navReset").addEventListener("click", () => show("reset"));
$("navDatabase").addEventListener("click", () => show("database"));

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
}

function saveImported() {
  localStorage.setItem("planogram_imported_rows", JSON.stringify(DATA));
  localStorage.setItem("planogram_imported_filename", fileName);
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

$("xlsxFile").addEventListener("change", async e => {
  const f = e.target.files?.[0];
  if (!f) return;

  $("uploadStatus").textContent = `Reading ${f.name}...`;

  try {
    const buf = await f.arrayBuffer();
    const parsed = parseWorkbook(buf);
    DATA = parsed.rows;
    fileName = f.name;
    saveImported();
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
renderReset();
renderDatabase();
