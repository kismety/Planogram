
let DATA = [];
let currentCol = "";
let fileName = "";
let currentStore = "";
let currentStoreAddress = "";
let currentStoreCity = "";
let currentStoreId = "";
let currentScanImage = "";
let EXTRACTED_SCAN_ROWS = [];

const KNOWN_PRODUCTS = [
  ["Sephora","Retail / Beauty"],["Xbox","Gaming"],["PlayStation","Gaming"],
  ["Nintendo","Gaming"],["Roblox","Gaming"],["Google Play","Gaming / Digital"],
  ["Boston Pizza","Dining"],["The Keg","Dining"],["Keg","Dining"],
  ["Starbucks","Dining"],["Tim Hortons","Dining"],["Subway","Dining"],
  ["DoorDash","Dining"],["Uber Eats","Dining"],["Dairy Queen","Dining"],
  ["McDonald's","Dining"],["Wendy's","Dining"],["Moxies","Dining"],
  ["Earls","Dining"],["Cactus Club Cafe","Dining"],["Chop Steakhouse","Dining"],
  ["Home Depot","Home"],["Home Hardware","Home"],["IKEA","Home"],["RONA","Home"],
  ["Homesense","Retail / Home"],["Urban Barn","Home"],["JYSK","Home"],
  ["PetSmart","Pets"],["Pet Valu","Pets"],["Esso","Fuel / Auto"],
  ["WestJet","Travel"],["Air Canada","Travel"],["Hotels.com","Travel"],["Uber","Travel / Transportation"],
  ["Amazon","Retail / Online"],["Best Buy","Retail / Electronics"],["Staples","Retail / Office"],
  ["Dollarama","Retail"],["Michaels","Retail / Crafts"],["Golf Town","Retail / Sports"],
  ["Cabela's","Retail / Sports"],["Sport Chek","Retail / Sports"],["Sporting Life","Retail / Sports"],
  ["Lululemon","Retail / Fashion"],["Roots","Retail / Fashion"],["Winners","Retail"],
  ["Marshalls","Retail"],["Indigo","Retail / Books"],["Apple","Digital / Technology"],
  ["Visa","Open Loop / Payment"],["Mastercard","Open Loop / Payment"],
  ["One4all","Multi-brand"],["Happy Teen","Multi-brand"],["Happy Kid","Multi-brand"],
  ["Happy Birthday","Multi-brand"],["Happy Moments","Multi-brand"],["Home Sweet Home","Multi-brand"],
  ["Retail Therapy","Multi-brand"],["Fun & Fabulous","Multi-brand"],["Let's Eat","Multi-brand"],
  ["Cheers To You","Multi-brand"],["Game & Grub","Multi-brand"]
];

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
renderStores();
  if (screen === "stores") renderStores();
  if (screen === "final") renderFinalReview();
renderSavedPlans();
}

$("navStores").addEventListener("click", () => { show("stores"); renderStores(); });
$("navUpload").addEventListener("click", () => show("upload"));
$("navReset").addEventListener("click", () => show("reset"));
$("navDatabase").addEventListener("click", () => show("database"));
$("navPlans").addEventListener("click", () => { show("plans"); renderSavedPlans(); });
$("navFinal").addEventListener("click", () => { show("final"); renderFinalReview(); });

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




function normalizeOcrText(s) {
  return String(s || "")
    .replace(/[|_[\]{}]/g," ")
    .replace(/\s+/g," ")
    .trim();
}

function categoryForProduct(name) {
  const low = String(name||"").toLowerCase();
  const sorted = [...KNOWN_PRODUCTS].sort((a,b)=>b[0].length-a[0].length);
  for (const [brand,cat] of sorted) {
    if (low.includes(brand.toLowerCase())) return cat;
  }
  return "";
}

function levenshtein(a,b) {
  a=String(a||"").toLowerCase(); b=String(b||"").toLowerCase();
  const m=a.length,n=b.length;
  const dp=Array.from({length:m+1},()=>Array(n+1).fill(0));
  for(let i=0;i<=m;i++)dp[i][0]=i;
  for(let j=0;j<=n;j++)dp[0][j]=j;
  for(let i=1;i<=m;i++){
    for(let j=1;j<=n;j++){
      dp[i][j]=Math.min(
        dp[i-1][j]+1,
        dp[i][j-1]+1,
        dp[i-1][j-1]+(a[i-1]===b[j-1]?0:1)
      );
    }
  }
  return dp[m][n];
}

function normalizeKnownProduct(raw) {
  const clean = normalizeOcrText(raw);
  if (!clean) return {name:"",category:"",confidence:"Low"};
  let best=null, bestScore=Infinity;
  for (const [brand,cat] of KNOWN_PRODUCTS) {
    const score=levenshtein(clean,brand);
    const norm=score/Math.max(clean.length,brand.length,1);
    if (norm<bestScore) {
      bestScore=norm;
      best={brand,cat};
    }
  }
  if (best && bestScore<=0.28) return {name:best.brand,category:best.cat,confidence:"Medium"};
  const exactCat=categoryForProduct(clean);
  return {name:clean,category:exactCat,confidence:exactCat?"High":"Low"};
}

function splitColumnTextToRows(text, rowCount) {
  let lines = String(text||"")
    .split(/\n+/)
    .map(normalizeOcrText)
    .filter(x=>x.length>=2)
    .filter(x=>!/^[\d$.,%\- ]+$/.test(x));

  // Merge obvious denomination-only fragments into previous line.
  const merged=[];
  for (const line of lines) {
    if (/^\$?\d+(\s*[-–]\s*\$?\d+)?$/.test(line) && merged.length) {
      merged[merged.length-1] += " " + line;
    } else {
      merged.push(line);
    }
  }
  lines=merged;

  const out=[];
  for(let i=0;i<rowCount;i++) {
    const raw = lines[i] || "";
    const norm = normalizeKnownProduct(raw);
    const denomMatch = raw.match(/\$?\d+\s*[-–]\s*\$?\d+|\$?\d+/);
    out.push({
      raw,
      name:norm.name || (raw ? raw : "VERIFY - unreadable"),
      denomination:denomMatch ? denomMatch[0].replace(/\$/g,"") : "",
      category:norm.category || "",
      confidence: raw ? norm.confidence : "Low"
    });
  }
  return out;
}

async function imageToColumnCanvas(img, colIndex, totalCols) {
  const c=document.createElement("canvas");
  const colW=img.naturalWidth/totalCols;
  const scale=Math.max(1,Math.min(2.3,1200/colW));
  c.width=Math.round(colW*scale);
  c.height=Math.round(img.naturalHeight*scale);
  const ctx=c.getContext("2d");
  ctx.filter="grayscale(1) contrast(1.55)";
  ctx.drawImage(
    img,
    colIndex*colW,0,colW,img.naturalHeight,
    0,0,c.width,c.height
  );
  return c;
}

async function extractScanToRows() {
  if (!currentScanImage) throw new Error("Attach a planogram image first.");
  if (typeof Tesseract === "undefined") throw new Error("OCR engine did not load.");

  const cols=Math.max(1,Math.min(30,Number($("scanColumns").value)||1));
  const rows=Math.max(1,Math.min(40,Number($("scanRows").value)||1));

  const img=new Image();
  await new Promise((resolve,reject)=>{
    img.onload=resolve;
    img.onerror=()=>reject(new Error("Could not read image."));
    img.src=currentScanImage;
  });

  const worker=await Tesseract.createWorker("eng",1,{
    logger:m=>{
      if (m.status==="recognizing text") {
        $("extractStatus").textContent="Reading scan… "+Math.round((m.progress||0)*100)+"%";
      }
    }
  });
  await worker.setParameters({tessedit_pageseg_mode:"6", preserve_interword_spaces:"1"});

  const extracted=[];
  for(let ci=0;ci<cols;ci++) {
    $("extractStatus").textContent=`Reading column C${ci+1} of C${cols}…`;
    const canvas=await imageToColumnCanvas(img,ci,cols);
    const result=await worker.recognize(canvas);
    const parsed=splitColumnTextToRows(result.data.text,rows);
    parsed.forEach((item,ri)=>{
      extracted.push({
        id:`scan|${ci}|${ri}|${Date.now()}`,
        planogram:$("planogramName").value.trim() || "Scanned Planogram",
        card:item.name,
        denomination:item.denomination,
        category:item.category,
        column:"C"+(ci+1),
        row:"R"+(ri+1),
        position:"C"+(ci+1)+"-R"+(ri+1),
        notes:item.raw && item.raw!==item.name ? "OCR: "+item.raw : "",
        confidence:item.confidence
      });
    });
  }

  await worker.terminate();
  EXTRACTED_SCAN_ROWS=extracted;
  return extracted;
}

function renderExtractedReview() {
  if (!$("extractedReviewList")) return;
  if (!EXTRACTED_SCAN_ROWS.length) {
    $("extractedReviewCard").style.display="none";
    $("extractedReviewList").innerHTML="";
    return;
  }

  $("extractedReviewCard").style.display="block";
  $("extractedReviewList").innerHTML=EXTRACTED_SCAN_ROWS.map((r,i)=>`
    <div class="extractrow">
      <div>
        <div class="extractpos">${esc(r.position)}</div>
        <div class="extractconf">${esc(r.confidence || "")}</div>
      </div>
      <input data-er-name="${i}" value="${esc(r.card)}" placeholder="Card name">
      <input class="denomfield" data-er-denom="${i}" value="${esc(r.denomination||"")}" placeholder="Denomination">
      <input class="catfield" data-er-cat="${i}" value="${esc(r.category||"")}" placeholder="Category">
    </div>
  `).join("");

  document.querySelectorAll("[data-er-name]").forEach(el=>el.onchange=()=>{
    const i=Number(el.dataset.erName);
    EXTRACTED_SCAN_ROWS[i].card=el.value.trim();
    if (!EXTRACTED_SCAN_ROWS[i].category) {
      EXTRACTED_SCAN_ROWS[i].category=categoryForProduct(el.value.trim());
    }
  });
  document.querySelectorAll("[data-er-denom]").forEach(el=>el.onchange=()=>{
    EXTRACTED_SCAN_ROWS[Number(el.dataset.erDenom)].denomination=el.value.trim();
  });
  document.querySelectorAll("[data-er-cat]").forEach(el=>el.onchange=()=>{
    EXTRACTED_SCAN_ROWS[Number(el.dataset.erCat)].category=el.value.trim();
  });
}

function createPlanogramFromScan() {
  if (!EXTRACTED_SCAN_ROWS.length) {
    alert("Extract the scan first.");
    return;
  }
  DATA=EXTRACTED_SCAN_ROWS.map((r,i)=>({...r,id:`scanrow|${i}|${r.position}|${r.card}`}));
  fileName=($("planogramName").value.trim() || "Scanned Planogram") + ".scan";
  currentCol=columns()[0] || "";
  saveImported();
  afterLoad();
  $("uploadStatus").textContent="✓ Scanned planogram created. Review the Reset tab, then Save Planogram.";
  show("reset");
}

function exportExtractedExcel() {
  if (!EXTRACTED_SCAN_ROWS.length) {
    alert("Extract the scan first.");
    return;
  }
  if (typeof XLSX==="undefined") {
    alert("Excel export engine did not load.");
    return;
  }

  const store=$("storeName").value.trim();
  const address=$("storeAddress").value.trim();
  const city=$("storeCity").value.trim();
  const planName=$("planogramName").value.trim() || "Scanned Planogram";

  const aoa=[["Store Name","Address","City","Planogram Name","Card Name","Denomination","Category","Column","Row","Position","Confidence","Notes"]];
  EXTRACTED_SCAN_ROWS.forEach(r=>aoa.push([
    store,address,city,planName,r.card,r.denomination,r.category,r.column,r.row,r.position,r.confidence,r.notes
  ]));

  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb,ws,"Extracted Positions");
  const safe=(store+"_"+planName).replace(/[^a-z0-9_-]+/gi,"_");
  XLSX.writeFile(wb,safe+"_extracted.xlsx");
}

function storeIndex() {
  try { return JSON.parse(localStorage.getItem("planogram_stores") || "[]"); }
  catch (_) { return []; }
}
function saveStoreIndex(list) {
  localStorage.setItem("planogram_stores", JSON.stringify(list));
}
function storeKey(name,address,city) {
  return [name,address,city].map(x=>String(x||"").trim().toLowerCase()).join("|");
}

function googleStoreSearchUrl(name, city) {
  const q = [name, city].filter(Boolean).join(", ");
  return "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(q);
}

function mapsUrl(address, city) {
  const q = [address, city].filter(Boolean).join(", ");
  return "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(q);
}
function saveStoreRecord(name,address,city,storeId) {
  name = String(name||"").trim();
  address = String(address||"").trim();
  city = String(city||"").trim();
  storeId = String(storeId||"").trim();
  if (!name) throw new Error("Store name is required.");
  if (!address) throw new Error("Store address is required.");

  const list = storeIndex();
  const key = storeKey(name,address,city);
  const existing = list.find(s => s.key === key);
  const rec = {key,name,address,city,storeId,updatedAt:new Date().toISOString()};
  if (existing) Object.assign(existing, rec);
  else list.push(rec);
  list.sort((a,b)=>a.name.localeCompare(b.name));
  saveStoreIndex(list);
  return rec;
}
function selectStore(rec) {
  currentStore = rec.name || "";
  currentStoreAddress = rec.address || "";
  currentStoreCity = rec.city || "";
  currentStoreId = rec.storeId || "";

  if ($("storeName")) $("storeName").value = currentStore;
  if ($("storeAddress")) $("storeAddress").value = currentStoreAddress;
  if ($("storeCity")) $("storeCity").value = currentStoreCity;
  if ($("storeId")) $("storeId").value = currentStoreId;

  show("upload");
}
function editStoreForm(rec) {
  $("storeMgrName").value = rec.name || "";
  $("storeMgrAddress").value = rec.address || "";
  $("storeMgrCity").value = rec.city || "";
  $("storeMgrId").value = rec.storeId || "";
  window.scrollTo({top:0,behavior:"smooth"});
}
function deleteStore(key) {
  const list = storeIndex();
  const rec = list.find(s=>s.key===key);
  if (!rec) return;
  if (!confirm('Delete store "' + rec.name + '"?')) return;
  saveStoreIndex(list.filter(s=>s.key!==key));
  renderStores();
}
function renderStores() {
  if (!$("storesList")) return;
  const q = ($("storeSearch")?.value || "").trim().toLowerCase();
  const list = storeIndex().filter(s =>
    !q ||
    s.name.toLowerCase().includes(q) ||
    (s.address||"").toLowerCase().includes(q) ||
    (s.city||"").toLowerCase().includes(q) ||
    (s.storeId||"").toLowerCase().includes(q)
  );

  if (!list.length) {
    $("storesList").innerHTML = '<div class="card">No saved stores yet.</div>';
    return;
  }

  $("storesList").innerHTML = list.map(s => `
    <div class="storecard">
      <div class="storetitle">${esc(s.name)}</div>
      <div class="storeaddress">${esc([s.address,s.city].filter(Boolean).join(", "))}</div>
      ${s.storeId ? `<div class="small">Store ID: ${esc(s.storeId)}</div>` : ""}
      <div class="storeactions">
        <button class="primary" data-select-store="${esc(s.key)}">Use Store</button>
        <button data-map-store="${esc(s.key)}">Google Maps</button>
        <button data-edit-store="${esc(s.key)}">Edit</button>
      </div>
      <button data-delete-store="${esc(s.key)}" style="width:100%;margin-top:7px">Delete Store</button>
    </div>
  `).join("");

  document.querySelectorAll("[data-select-store]").forEach(b => b.onclick = () => {
    const rec = storeIndex().find(s=>s.key===b.dataset.selectStore);
    if (rec) selectStore(rec);
  });
  document.querySelectorAll("[data-map-store]").forEach(b => b.onclick = () => {
    const rec = storeIndex().find(s=>s.key===b.dataset.mapStore);
    if (rec) window.open(mapsUrl(rec.address,rec.city), "_blank");
  });
  document.querySelectorAll("[data-edit-store]").forEach(b => b.onclick = () => {
    const rec = storeIndex().find(s=>s.key===b.dataset.editStore);
    if (rec) editStoreForm(rec);
  });
  document.querySelectorAll("[data-delete-store]").forEach(b => b.onclick = () => deleteStore(b.dataset.deleteStore));
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




function finalCounts() {
  const counts = {complete:0, missing:0, notComplete:0, total:DATA.length};
  DATA.forEach(r => {
    const s = getStatus(r);
    if (s === "Complete") counts.complete++;
    else if (s === "Missing") counts.missing++;
    else counts.notComplete++;
  });
  return counts;
}


function renderVisualPlanogram() {
  if (!$("visualPlanogram")) return;
  if (!DATA.length) {
    $("visualPlanogram").innerHTML = '<div class="small">Upload a planogram to see the visual layout.</div>';
    return;
  }

  const cols = columns();
  $("visualPlanogram").innerHTML = cols.map(c => {
    const rows = DATA.filter(r => r.column === c)
      .sort((a,b) => rowNumber(a.row)-rowNumber(b.row));

    return `<div class="visual-column">
      <div class="visual-column-title">${esc(c)}</div>
      ${rows.map(r => {
        const s = getStatus(r);
        const cls = s === "Complete" ? "complete" : s === "Missing" ? "missing" : "notcomplete";
        return `<div class="visual-cell ${cls}">
          <b>${esc(r.position || (c + "-" + r.row))}</b>
          <div>${esc(r.card)}</div>
          ${r.denomination ? `<div class="meta">${esc(r.denomination)}</div>` : ""}
          <div class="visual-status">${esc(s)}</div>
        </div>`;
      }).join("")}
    </div>`;
  }).join("");
}


let PLANO_PNG_BLOB = null;
let PLANO_PNG_NAME = "";

async function buildPlanoPNG() {
  if (!DATA.length) throw new Error("Upload or create a planogram first.");

  const store = ($("storeName")?.value || currentStore || "Store").trim();
  const address = ($("storeAddress")?.value || currentStoreAddress || "").trim();
  const city = ($("storeCity")?.value || currentStoreCity || "").trim();
  const planName = ($("planogramName")?.value || fileName || "Planogram").trim();
  const cols = columns();
  if (!cols.length) throw new Error("No columns found.");

  const grouped = {};
  cols.forEach(c => {
    grouped[c] = DATA.filter(r => r.column === c)
      .sort((a,b) => rowNumber(a.row)-rowNumber(b.row));
  });

  const colW = 300;
  const gap = 16;
  const margin = 36;
  const headerH = 140;
  const rowH = 78;
  const maxRows = Math.max(...cols.map(c => grouped[c].length), 1);

  const canvas = $("planoExportCanvas");
  const ctx = canvas.getContext("2d");

  canvas.width = margin*2 + cols.length*colW + Math.max(0, cols.length-1)*gap;
  canvas.height = headerH + 48 + maxRows*rowH + 70;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0,0,canvas.width,canvas.height);

  // Header
  ctx.fillStyle = "#111111";
  ctx.font = "bold 34px Arial";
  ctx.fillText(store, margin, 42);

  ctx.font = "bold 25px Arial";
  ctx.fillText(planName, margin, 78);

  ctx.fillStyle = "#555555";
  ctx.font = "15px Arial";
  const location = [address, city].filter(Boolean).join(", ");
  if (location) ctx.fillText(location, margin, 104);

  const counts = finalCounts();
  ctx.font = "14px Arial";
  ctx.fillText(
    `${counts.total} Total • ${counts.complete} Complete • ${counts.missing} Missing • ${counts.notComplete} Not Complete`,
    margin, 128
  );

  cols.forEach((c, ci) => {
    const x = margin + ci*(colW+gap);

    ctx.fillStyle = "#111111";
    ctx.fillRect(x, headerH, colW, 44);

    ctx.fillStyle = "#ffffff";
    ctx.font = "bold 21px Arial";
    ctx.fillText(c, x+12, headerH+29);

    grouped[c].forEach((r, ri) => {
      const y = headerH + 44 + ri*rowH;
      const status = getStatus(r);

      // Requested colors
      if (status === "Complete") ctx.fillStyle = "#e7f5e5";
      else if (status === "Missing") ctx.fillStyle = "#fff3bf";
      else ctx.fillStyle = "#fde2e2";

      ctx.fillRect(x,y,colW,rowH);
      ctx.strokeStyle = "#cfcfcf";
      ctx.strokeRect(x,y,colW,rowH);

      ctx.fillStyle = "#111111";
      ctx.font = "bold 13px Arial";
      ctx.fillText(r.position || `${c}-${r.row}`, x+9, y+18);

      ctx.font = "14px Arial";
      const label = String(r.card||"");
      const shortLabel = label.length > 34 ? label.slice(0,33)+"…" : label;
      ctx.fillText(shortLabel, x+9, y+41);

      if (status === "Complete") ctx.fillStyle = "#2f7d32";
      else if (status === "Missing") ctx.fillStyle = "#946e00";
      else ctx.fillStyle = "#a33c3c";

      ctx.font = "bold 12px Arial";
      ctx.fillText(status, x+9, y+61);

      if (r.denomination) {
        ctx.fillStyle = "#666666";
        ctx.font = "11px Arial";
        ctx.fillText(String(r.denomination), x+colW-8, y+61, 95);
      }
    });
  });

  // Legend
  const ly = canvas.height - 34;
  const legend = [
    ["#e7f5e5","Complete"],
    ["#fff3bf","Missing"],
    ["#fde2e2","Not Complete"]
  ];
  let lx = margin;
  legend.forEach(([color,label]) => {
    ctx.fillStyle = color;
    ctx.fillRect(lx,ly,22,22);
    ctx.strokeStyle = "#bbb";
    ctx.strokeRect(lx,ly,22,22);
    ctx.fillStyle = "#111";
    ctx.font = "13px Arial";
    ctx.fillText(label,lx+30,ly+16);
    lx += 150;
  });

  const blob = await new Promise((resolve,reject) => {
    canvas.toBlob(b => b ? resolve(b) : reject(new Error("PNG generation failed.")), "image/png");
  });

  const safeStore = store.replace(/[^a-z0-9_-]+/gi,"_");
  const safePlan = planName.replace(/[^a-z0-9_-]+/gi,"_");
  PLANO_PNG_NAME = `${safeStore}_${safePlan}.png`;
  PLANO_PNG_BLOB = blob;

  return {blob,name:PLANO_PNG_NAME};
}

async function generatePlanoPNG() {
  try {
    const {blob,name} = await buildPlanoPNG();
    const url = URL.createObjectURL(blob);

    // Most reliable iPhone behavior: show the generated image visibly in the page.
    $("planoPngPreview").src = url;
    $("planoPngPreview").style.display = "block";
    $("planoPngStatus").textContent =
      "✓ PNG generated below. On iPhone, press and hold the image or tap Share PNG.";

    $("planoPngPreview").scrollIntoView({behavior:"smooth",block:"start"});
  } catch (e) {
    $("planoPngStatus").textContent = "PNG failed: " + e.message;
  }
}

async function sharePlanoPNG() {
  try {
    const result = PLANO_PNG_BLOB
      ? {blob:PLANO_PNG_BLOB,name:PLANO_PNG_NAME}
      : await buildPlanoPNG();

    const file = new File([result.blob], result.name, {type:"image/png"});

    if (navigator.canShare && navigator.canShare({files:[file]})) {
      await navigator.share({title:"Plano", files:[file]});
      $("planoPngStatus").textContent = "✓ PNG shared.";
    } else {
      const url = URL.createObjectURL(result.blob);
      $("planoPngPreview").src = url;
      $("planoPngPreview").style.display = "block";
      $("planoPngStatus").textContent =
        "PNG is shown below. Press and hold it to save/share from Safari.";
    }
  } catch (e) {
    if (e.name !== "AbortError") {
      $("planoPngStatus").textContent = "Could not share PNG: " + e.message;
    }
  }
}

function renderFinalReview() {
  if (!$("finalComplete")) return;
  renderVisualPlanogram();

  const counts = finalCounts();
  if ($("progressIdentity")) {
    const store = ($("storeName")?.value || currentStore || "Store").trim();
    const planName = ($("planogramName")?.value || fileName || "Planogram").trim();
    const address = ($("storeAddress")?.value || currentStoreAddress || "").trim();
    const city = ($("storeCity")?.value || currentStoreCity || "").trim();
    $("progressIdentity").innerHTML = `<div>${esc(store)} • ${esc(planName)}</div>` +
      ((address||city) ? `<div class="small" style="margin-top:4px">${esc([address,city].filter(Boolean).join(", "))}</div>` : "");
  }
  $("finalTotal").textContent = counts.total;
  $("finalComplete").textContent = counts.complete;
  $("finalMissing").textContent = counts.missing;
  $("finalRemaining").textContent = counts.notComplete;

  const progressPct = counts.total ? Math.round((counts.complete / counts.total) * 100) : 0;
  const resolvedPct = counts.total ? Math.round(((counts.complete + counts.missing) / counts.total) * 100) : 0;

  $("finalFill").style.width = progressPct + "%";

  if (!counts.total) {
    $("finalMessage").textContent = "No planogram loaded.";
    $("columnProgressList").innerHTML = "";
    $("outstandingList").innerHTML = '<div class="small">Upload a planogram first.</div>';
    return;
  }

  $("finalMessage").textContent =
    `${progressPct}% complete • ${counts.complete} complete • ${counts.missing} missing • ${counts.notComplete} not complete • ${resolvedPct}% reviewed`;

  // Progress by column
  const cols = columns();
  $("columnProgressList").innerHTML = cols.map(c => {
    const rows = DATA.filter(r => r.column === c);
    const complete = rows.filter(r => getStatus(r) === "Complete").length;
    const missing = rows.filter(r => getStatus(r) === "Missing").length;
    const notComplete = rows.length - complete - missing;
    const pct = rows.length ? Math.round((complete / rows.length) * 100) : 0;

    return `<div class="colprogress">
      <div class="colprogress-head">
        <span>${esc(c)}</span>
        <span>${complete}/${rows.length} complete • ${missing} missing • ${notComplete} not complete</span>
      </div>
      <div class="colbar"><div class="colbarfill" style="width:${pct}%"></div></div>
    </div>`;
  }).join("");

  const outstanding = DATA
    .filter(r => getStatus(r) !== "Complete")
    .sort((a,b) =>
      columnNumber(a.column)-columnNumber(b.column) ||
      rowNumber(a.row)-rowNumber(b.row)
    );

  if (!outstanding.length) {
    $("outstandingList").innerHTML =
      '<div class="finalrow status-complete"><b>✓ All items are complete.</b></div>';
    return;
  }

  $("outstandingList").innerHTML = outstanding.map(r => {
    const s = getStatus(r);
    const cls = s === "Missing" ? "status-missing" : "status-notcomplete";
    return `<div class="finalrow ${cls}">
      <b>${esc(r.position || r.column)} — ${esc(r.card)}</b>
      <span class="small">${esc(s)}${r.denomination ? " • " + esc(r.denomination) : ""}</span>
    </div>`;
  }).join("");
}

async function buildFinalPlanogramBlob() {
  if (!DATA.length) throw new Error("Upload a planogram first.");

  const store = ($("storeName")?.value || currentStore || "Store").trim();
  const planName = ($("planogramName")?.value || "Planogram").trim();
  const cols = columns();
  if (!cols.length) throw new Error("No columns found.");

  const grouped = {};
  cols.forEach(c => grouped[c] = DATA.filter(r => r.column === c)
    .sort((a,b) => rowNumber(a.row)-rowNumber(b.row)));

  const colWidth = 320;
  const gap = 18;
  const margin = 40;
  const headerH = 155;
  const rowH = 88;
  const maxRows = Math.max(...cols.map(c => grouped[c].length), 1);

  const canvas = $("exportCanvas");
  const ctx = canvas.getContext("2d");
  canvas.width = margin*2 + cols.length*colWidth + (cols.length-1)*gap;
  canvas.height = headerH + 50 + maxRows*rowH + 110;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0,0,canvas.width,canvas.height);

  const counts = finalCounts();

  ctx.fillStyle = "#111111";
  ctx.font = "bold 34px Arial";
  ctx.fillText(store, margin, 45);
  ctx.font = "bold 26px Arial";
  ctx.fillText(planName, margin, 82);
  ctx.font = "16px Arial";
  ctx.fillStyle = "#555555";
  ctx.fillText(
    `Reset Progress • ${counts.complete}/${counts.total} Complete • ${counts.missing} Missing • ${counts.notComplete} Not Complete`,
    margin, 112
  );
  ctx.fillText(new Date().toLocaleString(), margin, 136);

  cols.forEach((c, ci) => {
    const x = margin + ci*(colWidth+gap);
    ctx.fillStyle = "#111111";
    ctx.fillRect(x, headerH, colWidth, 46);
    ctx.fillStyle = "#ffffff";
    ctx.font = "bold 22px Arial";
    ctx.fillText(c, x+14, headerH+30);

    grouped[c].forEach((r, ri) => {
      const y = headerH + 46 + ri*rowH;
      const status = getStatus(r);

      if (status === "Complete") ctx.fillStyle = "#e7f4e4";
      else if (status === "Missing") ctx.fillStyle = "#fde9e7";
      else ctx.fillStyle = "#fff6d9";

      ctx.fillRect(x, y, colWidth, rowH);
      ctx.strokeStyle = "#cccccc";
      ctx.strokeRect(x, y, colWidth, rowH);

      ctx.fillStyle = "#111111";
      ctx.font = "bold 14px Arial";
      ctx.fillText(r.position || (c + "-" + r.row), x+10, y+19);

      ctx.font = "14px Arial";
      const name = r.card || "";
      const maxChars = 36;
      const displayName = name.length > maxChars ? name.slice(0,maxChars-1)+"…" : name;
      ctx.fillText(displayName, x+10, y+42);

      ctx.font = "bold 12px Arial";
      ctx.fillStyle = status === "Missing" ? "#a94442" : status === "Complete" ? "#2d6a2d" : "#8a6d00";
      ctx.fillText(status, x+10, y+63);

      if (r.denomination) {
        ctx.font = "11px Arial";
        ctx.fillStyle = "#666666";
        ctx.fillText(r.denomination, x+10, y+79);
      }
    });
  });

  const legendY = canvas.height - 48;
  const legend = [
    ["#e7f4e4","Complete"],
    ["#fde9e7","Missing"],
    ["#fff6d9","Not Complete"]
  ];
  let lx = margin;
  legend.forEach(([color,label]) => {
    ctx.fillStyle = color;
    ctx.fillRect(lx, legendY, 24, 24);
    ctx.strokeStyle = "#bbbbbb";
    ctx.strokeRect(lx, legendY, 24, 24);
    ctx.fillStyle = "#111111";
    ctx.font = "14px Arial";
    ctx.fillText(label, lx+32, legendY+17);
    lx += 150;
  });

  const blob = await new Promise((resolve,reject) => {
    canvas.toBlob(b => b ? resolve(b) : reject(new Error("PNG generation failed.")), "image/png");
  });

  const safeStore = store.replace(/[^a-z0-9_-]+/gi,"_");
  const safePlan = planName.replace(/[^a-z0-9_-]+/gi,"_");
  const name = safeStore + "_" + safePlan + "_FINAL.png";
  return {blob,name};
}

async function generateFinalPNG() {
  try {
    const {blob,name} = await buildFinalPlanogramBlob();
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.target = "_blank";
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();

    $("finalPngStatus").textContent =
      "✓ Final PNG generated. On iPhone, use Share → Save to Files or Photos.";
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (e) {
    $("finalPngStatus").textContent = "Final PNG failed: " + e.message;
  }
}

async function shareFinalPNG() {
  try {
    const {blob,name} = await buildFinalPlanogramBlob();
    const file = new File([blob], name, {type:"image/png"});

    if (navigator.canShare && navigator.canShare({files:[file]})) {
      await navigator.share({title:"Final Planogram", files:[file]});
      $("finalPngStatus").textContent = "✓ Final PNG shared.";
    } else {
      const url = URL.createObjectURL(blob);
      window.open(url, "_blank");
      $("finalPngStatus").textContent = "Final PNG opened. Use Safari Share to save it.";
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    }
  } catch (e) {
    if (e.name !== "AbortError") {
      $("finalPngStatus").textContent = "Could not share final PNG: " + e.message;
    }
  }
}


let LAST_PDF_BLOB = null;
let LAST_PDF_NAME = "";

function pdfColorForStatus(status) {
  if (status === "Complete") return [231,245,229];
  if (status === "Missing") return [255,243,191];
  return [253,226,226];
}

async function buildFinalPlanogramPDFBlob() {
  if (!DATA.length) throw new Error("Upload a planogram first.");
  if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF engine did not load.");

  const { jsPDF } = window.jspdf;
  const store = ($("storeName")?.value || currentStore || "Store").trim();
  const planName = ($("planogramName")?.value || "Planogram").trim();
  const cols = columns();
  if (!cols.length) throw new Error("No columns found.");

  // Landscape A4. Multiple columns will flow across multiple pages if needed.
  const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4"});
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();

  const margin = 10;
  const headerH = 28;
  const colW = 63;
  const gap = 3;
  const rowH = 15;
  const colsPerPage = Math.max(1, Math.floor((pageW - margin*2 + gap) / (colW + gap)));

  const counts = finalCounts();

  function drawPageHeader(pageCols, pageNum) {
    doc.setTextColor(20,20,20);
    doc.setFont("helvetica","bold");
    doc.setFontSize(16);
    doc.text(store, margin, 10);
    doc.setFontSize(12);
    doc.text(planName, margin, 17);

    doc.setFont("helvetica","normal");
    doc.setFontSize(8.5);
    doc.setTextColor(90,90,90);
    doc.text(
      `${counts.complete} Complete | ${counts.missing} Missing | ${counts.notComplete} Not Complete | ${DATA.length} Total`,
      margin, 23
    );
    doc.text(`Page ${pageNum}`, pageW - margin - 14, 10);

    // legend
    const ly = pageH - 7;
    const legend = [
      [[231,245,229],"Complete"],
      [[255,243,191],"Missing"],
      [[253,226,226],"Not Complete"]
    ];
    let lx = margin;
    legend.forEach(([rgb,label]) => {
      doc.setFillColor(...rgb);
      doc.rect(lx, ly-3.5, 5, 4.5, "F");
      doc.setDrawColor(190,190,190);
      doc.rect(lx, ly-3.5, 5, 4.5);
      doc.setTextColor(40,40,40);
      doc.setFontSize(7.5);
      doc.text(label, lx+7, ly);
      lx += 30;
    });
  }

  let pageNum = 1;
  for (let start=0; start<cols.length; start += colsPerPage) {
    if (start > 0) {
      doc.addPage("a4","landscape");
      pageNum++;
    }

    const pageCols = cols.slice(start, start + colsPerPage);
    drawPageHeader(pageCols, pageNum);

    pageCols.forEach((c, localIndex) => {
      const x = margin + localIndex*(colW+gap);
      const y0 = headerH;

      doc.setFillColor(20,20,20);
      doc.rect(x, y0, colW, 9, "F");
      doc.setTextColor(255,255,255);
      doc.setFont("helvetica","bold");
      doc.setFontSize(11);
      doc.text(c, x+3, y0+6);

      const rows = DATA.filter(r=>r.column===c)
        .sort((a,b)=>rowNumber(a.row)-rowNumber(b.row));

      let y = y0 + 9;
      rows.forEach(r => {
        const status = getStatus(r);
        const rgb = pdfColorForStatus(status);

        // If column exceeds page, continue on next page for this column.
        if (y + rowH > pageH - 14) {
          doc.addPage("a4","landscape");
          pageNum++;
          drawPageHeader([c], pageNum);
          y = headerH;
          doc.setFillColor(20,20,20);
          doc.rect(margin, y, colW, 9, "F");
          doc.setTextColor(255,255,255);
          doc.setFont("helvetica","bold");
          doc.setFontSize(11);
          doc.text(c + " (cont.)", margin+3, y+6);
          y += 9;
        }

        const dx = (y === y0 + 9) ? x : (y < headerH + 12 ? margin : x);
        const drawX = (pageNum > 1 && start===0 && y < headerH+20) ? margin : dx;

        doc.setFillColor(...rgb);
        doc.rect(drawX, y, colW, rowH, "F");
        doc.setDrawColor(210,210,210);
        doc.rect(drawX, y, colW, rowH);

        doc.setTextColor(20,20,20);
        doc.setFont("helvetica","bold");
        doc.setFontSize(7.5);
        doc.text((r.position || `${c}-${r.row}`).slice(0,18), drawX+2, y+4);

        doc.setFont("helvetica","normal");
        doc.setFontSize(8);
        const cardLines = doc.splitTextToSize(r.card || "", colW-4);
        doc.text(cardLines.slice(0,2), drawX+2, y+8);

        doc.setFont("helvetica","bold");
        doc.setFontSize(6.5);
        if (status === "Complete") doc.setTextColor(47,125,50);
        else if (status === "Missing") doc.setTextColor(145,106,0);
        else doc.setTextColor(160,55,55);
        doc.text(status, drawX+2, y+13);

        if (r.denomination) {
          doc.setTextColor(90,90,90);
          doc.setFont("helvetica","normal");
          doc.text(String(r.denomination).slice(0,18), drawX+colW-2, y+13, {align:"right"});
        }

        y += rowH;
      });
    });
  }

  const blob = doc.output("blob");
  const safeStore = store.replace(/[^a-z0-9_-]+/gi,"_");
  const safePlan = planName.replace(/[^a-z0-9_-]+/gi,"_");
  LAST_PDF_NAME = `${safeStore}_${safePlan}_FINAL.pdf`;
  LAST_PDF_BLOB = blob;
  return {blob, name:LAST_PDF_NAME};
}

async function generateFinalPDF() {
  try {
    const {blob,name} = await buildFinalPlanogramPDFBlob();
    const url = URL.createObjectURL(blob);

    // iPhone/Safari: open first because browser PDF viewer has a reliable Share button.
    const opened = window.open(url, "_blank");
    if (!opened) {
      const a = document.createElement("a");
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }

    $("finalPdfStatus").textContent =
      "✓ PDF generated. In Safari PDF view, tap Share to Save to Files or send it.";
    setTimeout(() => URL.revokeObjectURL(url), 120000);
  } catch (e) {
    $("finalPdfStatus").textContent = "PDF generation failed: " + e.message;
  }
}

async function openFinalPDF() {
  try {
    const result = LAST_PDF_BLOB
      ? {blob:LAST_PDF_BLOB,name:LAST_PDF_NAME}
      : await buildFinalPlanogramPDFBlob();

    const url = URL.createObjectURL(result.blob);
    window.open(url, "_blank");
    $("finalPdfStatus").textContent = "PDF opened. Use Safari Share → Save to Files.";
    setTimeout(() => URL.revokeObjectURL(url), 120000);
  } catch (e) {
    $("finalPdfStatus").textContent = "Could not open PDF: " + e.message;
  }
}

async function downloadFinalPDF() {
  try {
    const result = LAST_PDF_BLOB
      ? {blob:LAST_PDF_BLOB,name:LAST_PDF_NAME}
      : await buildFinalPlanogramPDFBlob();

    const url = URL.createObjectURL(result.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = result.name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    $("finalPdfStatus").textContent = "✓ PDF download requested.";
    setTimeout(() => URL.revokeObjectURL(url), 120000);
  } catch (e) {
    $("finalPdfStatus").textContent = "Could not download PDF: " + e.message;
  }
}

async function shareFinalPDF() {
  try {
    const result = LAST_PDF_BLOB
      ? {blob:LAST_PDF_BLOB,name:LAST_PDF_NAME}
      : await buildFinalPlanogramPDFBlob();

    const file = new File([result.blob], result.name, {type:"application/pdf"});

    if (navigator.canShare && navigator.canShare({files:[file]})) {
      await navigator.share({title:"Final Planogram", files:[file]});
      $("finalPdfStatus").textContent = "✓ PDF shared.";
    } else {
      const url = URL.createObjectURL(result.blob);
      window.open(url, "_blank");
      $("finalPdfStatus").textContent =
        "PDF opened. Safari can save/share it from the PDF viewer.";
      setTimeout(() => URL.revokeObjectURL(url), 120000);
    }
  } catch (e) {
    if (e.name !== "AbortError") {
      $("finalPdfStatus").textContent = "Could not share PDF: " + e.message;
    }
  }
}

let LAST_PNG_BLOB = null;
let LAST_PNG_NAME = "";

async function buildPlanogramPNGBlob() {
  if (!DATA.length) throw new Error("Upload a planogram first.");

  const store = ($("storeName")?.value || currentStore || "Store").trim();
  const planName = ($("planogramName")?.value || "Planogram").trim();
  const cols = columns();
  if (!cols.length) throw new Error("No columns found in this planogram.");

  const grouped = {};
  cols.forEach(c => grouped[c] = DATA.filter(r => r.column === c).sort((a,b) => rowNumber(a.row)-rowNumber(b.row)));

  const colWidth = 300, gap = 18, margin = 40, headerH = 125, rowH = 76;
  const maxRows = Math.max(...cols.map(c => grouped[c].length), 1);

  const canvas = $("exportCanvas");
  const ctx = canvas.getContext("2d");
  canvas.width = margin*2 + cols.length*colWidth + (cols.length-1)*gap;
  canvas.height = headerH + 50 + maxRows*rowH + 80;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0,0,canvas.width,canvas.height);

  ctx.fillStyle = "#111111";
  ctx.font = "bold 32px Arial";
  ctx.fillText(store, margin, 44);
  ctx.font = "bold 24px Arial";
  ctx.fillText(planName, margin, 80);
  ctx.fillStyle = "#666666";
  ctx.font = "16px Arial";
  ctx.fillText("Plano Layout", margin, 106);

  cols.forEach((c, ci) => {
    const x = margin + ci*(colWidth+gap);
    ctx.fillStyle = "#111111";
    ctx.fillRect(x, headerH, colWidth, 46);
    ctx.fillStyle = "#ffffff";
    ctx.font = "bold 22px Arial";
    ctx.fillText(c, x+14, headerH+30);

    grouped[c].forEach((r, ri) => {
      const y = headerH + 46 + ri*rowH;
      ctx.fillStyle = ri%2===0 ? "#f7f7f7" : "#ffffff";
      ctx.fillRect(x, y, colWidth, rowH);
      ctx.strokeStyle = "#dddddd";
      ctx.strokeRect(x, y, colWidth, rowH);

      ctx.fillStyle = "#111111";
      ctx.font = "bold 14px Arial";
      ctx.fillText(r.position || (c + "-" + r.row), x+10, y+19);

      ctx.font = "14px Arial";
      const name = r.card || "";
      const maxChars = 33;
      const displayName = name.length > maxChars ? name.slice(0,maxChars-1)+"…" : name;
      ctx.fillText(displayName, x+10, y+41);

      ctx.fillStyle = "#666666";
      ctx.font = "12px Arial";
      const extra = [r.denomination, getStatus(r)].filter(Boolean).join(" • ");
      if (extra) ctx.fillText(extra, x+10, y+61);
    });
  });

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob(b => b ? resolve(b) : reject(new Error("PNG generation failed.")), "image/png");
  });

  const safeStore = store.replace(/[^a-z0-9_-]+/gi,"_");
  const safePlan = planName.replace(/[^a-z0-9_-]+/gi,"_");
  LAST_PNG_NAME = safeStore + "_" + safePlan + ".png";
  LAST_PNG_BLOB = blob;
  return {blob, name: LAST_PNG_NAME};
}

async function exportPlanogramPNG() {
  try {
    const {blob, name} = await buildPlanogramPNGBlob();
    const url = URL.createObjectURL(blob);

    // Safari/iPhone: opening a real blob URL is more reliable than a hidden anchor/data URL.
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.target = "_blank";
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();

    $("pngStatus").textContent = "✓ PNG generated. If Safari opened the image, use Share → Save to Files/Photos.";
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (e) {
    $("pngStatus").textContent = "PNG export failed: " + e.message;
  }
}

async function openPlanogramPNG() {
  try {
    const {blob} = LAST_PNG_BLOB ? {blob:LAST_PNG_BLOB} : await buildPlanogramPNGBlob();
    const url = URL.createObjectURL(blob);
    window.open(url, "_blank");
    $("pngStatus").textContent = "PNG opened in a new tab. Use Safari Share to save it.";
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (e) {
    $("pngStatus").textContent = "Could not open PNG: " + e.message;
  }
}

async function sharePlanogramPNG() {
  try {
    const result = LAST_PNG_BLOB ? {blob:LAST_PNG_BLOB,name:LAST_PNG_NAME} : await buildPlanogramPNGBlob();
    const file = new File([result.blob], result.name, {type:"image/png"});
    if (navigator.canShare && navigator.canShare({files:[file]})) {
      await navigator.share({title:"Planogram", files:[file]});
      $("pngStatus").textContent = "✓ PNG shared.";
    } else {
      const url = URL.createObjectURL(result.blob);
      window.open(url, "_blank");
      $("pngStatus").textContent = "Share-files is not supported here. PNG opened instead.";
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    }
  } catch (e) {
    if (e.name !== "AbortError") $("pngStatus").textContent = "Could not share PNG: " + e.message;
  }
}

function saveAsNamedPlanogram() {
  if (!DATA.length) {
    alert("Upload an Excel planogram first.");
    return;
  }
  const name = $("planogramName").value.trim();
  const store = $("storeName").value.trim();
  const storeAddress = $("storeAddress").value.trim();
  const storeCity = $("storeCity").value.trim();
  const storeId = $("storeId").value.trim();
  currentStore = store;
  currentStoreAddress = storeAddress;
  currentStoreCity = storeCity;
  currentStoreId = storeId;
  if (!store) {
    $("uploadStatus").textContent = "Enter the store name first.";
    return;
  }
  if (!storeAddress) {
    $("uploadStatus").textContent = "Enter the store address first.";
    return;
  }
  if (!name) {
    $("uploadStatus").textContent = "Enter a name for the planogram first.";
    return;
  }

  try { saveStoreRecord(store, storeAddress, storeCity, storeId); } catch (_) {}

  const index = savedPlansIndex();
  const existing = index.find(p => p.name.toLowerCase() === name.toLowerCase() && (p.store||"").toLowerCase() === store.toLowerCase());
  const id = existing ? existing.id : safePlanId(store + "_" + name);

  const record = {
    id,
    name,
    store,
    storeAddress,
    storeCity,
    storeId,
    scanImage: currentScanImage || "",
    extractedRows: EXTRACTED_SCAN_ROWS || [],
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
    existing.storeAddress = storeAddress;
    existing.storeCity = storeCity;
    existing.storeId = storeId;
    existing.count = DATA.length;
    existing.fileName = fileName;
  } else {
    index.push({
      id,
      name,
      store,
      storeAddress,
      storeCity,
      storeId,
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
    currentStoreAddress = record.storeAddress || "";
    currentStoreCity = record.storeCity || "";
    currentStoreId = record.storeId || "";
    currentScanImage = record.scanImage || "";
    EXTRACTED_SCAN_ROWS = Array.isArray(record.extractedRows) ? record.extractedRows : [];
    if ($("planogramName")) $("planogramName").value = record.name || "";
    if ($("storeName")) $("storeName").value = record.store || "";
    if ($("storeAddress")) $("storeAddress").value = record.storeAddress || "";
    if ($("storeCity")) $("storeCity").value = record.storeCity || "";
    if ($("storeId")) $("storeId").value = record.storeId || "";
    renderExtractedReview();
    if ($("scanPreview")) {
      if (currentScanImage) {
        $("scanPreview").src = currentScanImage;
        $("scanPreview").style.display = "block";
      } else {
        $("scanPreview").style.display = "none";
      }
    }
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
          <div class="planmeta">${esc([p.storeAddress,p.storeCity].filter(Boolean).join(", "))}</div>
          <div class="planmeta">${p.count || 0} positions${p.fileName ? " • " + esc(p.fileName) : ""}</div>
          <div class="planmeta">Saved ${new Date(p.savedAt).toLocaleString()}</div>
        </div>
      </div>
      <div class="planbuttons">
        <button class="primary" data-open-plan="${esc(p.id)}">Open</button>
        <button data-map-plan="${esc(p.id)}">Map</button>
        <button data-delete-plan="${esc(p.id)}">Delete</button>
      </div>
    </div>`).join("");

  document.querySelectorAll("[data-open-plan]").forEach(b =>
    b.addEventListener("click", () => openSavedPlanogram(b.dataset.openPlan))
  );
  document.querySelectorAll("[data-map-plan]").forEach(b =>
    b.addEventListener("click", () => {
      const p = savedPlansIndex().find(x=>x.id===b.dataset.mapPlan);
      if (p && p.storeAddress) window.open(mapsUrl(p.storeAddress,p.storeCity), "_blank");
    })
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



$("extractScan").addEventListener("click", async () => {
  try {
    $("extractStatus").textContent="Preparing OCR…";
    await extractScanToRows();
    renderExtractedReview();
    $("extractStatus").textContent="✓ Extraction complete. Review the rows below before creating the planogram.";
    $("extractedReviewCard").scrollIntoView({behavior:"smooth",block:"start"});
  } catch (e) {
    console.error(e);
    $("extractStatus").textContent="Extraction failed: "+e.message;
  }
});

$("clearExtracted").addEventListener("click", () => {
  EXTRACTED_SCAN_ROWS=[];
  renderExtractedReview();
  $("extractStatus").textContent="Extracted scan cleared.";
});

$("createFromScan").addEventListener("click", createPlanogramFromScan);
$("exportExtractedExcel").addEventListener("click", exportExtractedExcel);

$("saveStore").addEventListener("click", () => {
  try {
    const rec = saveStoreRecord(
      $("storeMgrName").value,
      $("storeMgrAddress").value,
      $("storeMgrCity").value,
      $("storeMgrId").value
    );
    $("storeMgrName").value = rec.name;
    $("storeMgrAddress").value = rec.address;
    $("storeMgrCity").value = rec.city;
    $("storeMgrId").value = rec.storeId || "";
    renderStores();
  } catch (e) {
    alert(e.message);
  }
});
$("clearStoreForm").addEventListener("click", () => {
  $("storeMgrName").value="";
  $("storeMgrAddress").value="";
  $("storeMgrCity").value="";
  $("storeMgrId").value="";
});
$("storeSearch").addEventListener("input", renderStores);


$("findStoreGoogle").addEventListener("click", () => {
  const name = $("storeName").value.trim();
  const city = $("storeCity").value.trim();
  if (!name) { alert("Enter the store name first."); return; }
  window.open(googleStoreSearchUrl(name, city), "_blank");
});

$("findStoreMgrGoogle").addEventListener("click", () => {
  const name = $("storeMgrName").value.trim();
  const city = $("storeMgrCity").value.trim();
  if (!name) { alert("Enter the store name first."); return; }
  window.open(googleStoreSearchUrl(name, city), "_blank");
});

$("openCurrentMap").addEventListener("click", () => {
  const address = $("storeAddress").value.trim();
  const city = $("storeCity").value.trim();
  if (!address) { alert("Enter a store address first."); return; }
  window.open(mapsUrl(address,city), "_blank");
});

$("scanImage").addEventListener("change", e => {
  const f = e.target.files?.[0];
  if (!f) return;
  const r = new FileReader();
  r.onload = () => {
    currentScanImage = r.result;
    $("scanPreview").src = currentScanImage;
    $("scanPreview").style.display = "block";
    $("scanStatus").textContent = "✓ Planogram image attached. Save the planogram to retain it.";
  };
  r.readAsDataURL(f);
});

$("saveAsPlanogram").addEventListener("click", saveAsNamedPlanogram);

$("generateFinalPdf").addEventListener("click", generateFinalPDF);
$("shareFinalPdf").addEventListener("click", shareFinalPDF);
$("openFinalPdf").addEventListener("click", openFinalPDF);
$("downloadFinalPdf").addEventListener("click", downloadFinalPDF);



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
  const statusClass =
    status === "Complete" ? "card-complete" :
    status === "Missing" ? "card-missing" :
    "card-notcomplete";

  return `
  <div class="card ${statusClass}">
    <div class="line">
      <div class="pos">${esc(r.position || r.column)}</div>
      <div class="name">
        ${esc(r.card)}
        ${r.denomination ? `<div class="meta">Denomination: ${esc(r.denomination)}</div>` : ""}
        ${r.category ? `<div class="meta">Category: ${esc(r.category)}</div>` : ""}
        <div class="meta"><b>Status: ${esc(status)}</b></div>
      </div>
    </div>
    <div class="actions">
      <button class="complete ${status==="Complete"?"active":""}" data-id="${esc(r.id)}" data-status="Complete">✓ Complete</button>
      <button class="missing ${status==="Missing"?"active":""}" data-id="${esc(r.id)}" data-status="Missing">Missing</button>
      <button class="notcomplete ${status==="Not Complete"?"active":""}" data-id="${esc(r.id)}" data-status="Not Complete">Not Complete</button>
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
      renderFinalReview();
      renderDatabase();
      renderFinalReview();
      renderFinalReview();
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



function setRowsStatus(rows, status) {
  if (!DATA.length || !rows.length) return;
  rows.forEach(r => {
    localStorage.setItem(statusKey(r), status);
  });
  saveAll(false);
  renderReset();
  renderDatabase();
  renderFinalReview();
}

function currentColumnRowsForBulk() {
  const col = $("colFilter").value || currentCol;
  if (!col) return [];
  return DATA.filter(r => r.column === col);
}

function setCurrentColumnStatus(status) {
  const rows = currentColumnRowsForBulk();
  if (!rows.length) {
    alert("No current column is selected.");
    return;
  }
  setRowsStatus(rows, status);
}

function setAllPlanogramStatus(status) {
  if (!DATA.length) {
    alert("Upload a planogram first.");
    return;
  }
  setRowsStatus(DATA, status);
}

$("columnComplete").addEventListener("click", () => setCurrentColumnStatus("Complete"));
$("columnMissing").addEventListener("click", () => setCurrentColumnStatus("Missing"));
$("columnNotComplete").addEventListener("click", () => setCurrentColumnStatus("Not Complete"));

$("allComplete").addEventListener("click", () => setAllPlanogramStatus("Complete"));
$("allMissing").addEventListener("click", () => setAllPlanogramStatus("Missing"));
$("allNotComplete").addEventListener("click", () => setAllPlanogramStatus("Not Complete"));

$("generatePlanoPng").addEventListener("click", generatePlanoPNG);
$("sharePlanoPng").addEventListener("click", sharePlanoPNG);

$("search").addEventListener("input", renderReset);
$("statusFilter").addEventListener("change", renderReset);
$("colFilter").addEventListener("change", function() {
  if (this.value) currentCol = this.value;
  renderReset();
  renderTabs();
});


function fillDatabaseFilters() {
  if (!$("dbCategory")) return;
  const categories = [...new Set(DATA.map(r => r.category).filter(Boolean))].sort();
  const cols = columns();
  const oldCat = $("dbCategory").value;
  const oldCol = $("dbColumn").value;
  $("dbCategory").innerHTML = '<option value="">All categories</option>' + categories.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join("");
  $("dbColumn").innerHTML = '<option value="">All columns</option>' + cols.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join("");
  if (categories.includes(oldCat)) $("dbCategory").value = oldCat;
  if (cols.includes(oldCol)) $("dbColumn").value = oldCol;
}

function renderDatabase() {
  if (!$("dbList")) return;
  fillDatabaseFilters();

  const q = $("dbSearch").value.trim().toLowerCase();
  const category = $("dbCategory").value;
  const column = $("dbColumn").value;
  const status = $("dbStatus").value;
  const sort = $("dbSort").value;

  if (!DATA.length) {
    $("dbCount").textContent = "0 imported positions";
    $("dbList").innerHTML = '<div class="card">Upload an Excel planogram first.</div>';
    return;
  }

  const filteredRows = DATA.filter(r =>
    (!q || r.card.toLowerCase().includes(q)) &&
    (!category || r.category === category) &&
    (!column || r.column === column) &&
    (!status || getStatus(r) === status)
  );

  const products = new Map();
  filteredRows.forEach(r => {
    const key = r.card.toLowerCase();
    if (!products.has(key)) {
      products.set(key, { name:r.card, category:r.category, rows:[] });
    }
    products.get(key).rows.push(r);
  });

  let list = [...products.values()];
  if (sort === "category") {
    list.sort((a,b) => (a.category||"").localeCompare(b.category||"") || a.name.localeCompare(b.name));
  } else if (sort === "column") {
    list.sort((a,b) => {
      const ac = Math.min(...a.rows.map(r => columnNumber(r.column)));
      const bc = Math.min(...b.rows.map(r => columnNumber(r.column)));
      return ac-bc || a.name.localeCompare(b.name);
    });
  } else {
    list.sort((a,b) => a.name.localeCompare(b.name));
  }

  $("dbCount").textContent = `${list.length} products • ${filteredRows.length} matching positions`;

  $("dbList").innerHTML = list.map(p => {
    const counts = {Complete:0,"Not Complete":0,Missing:0};
    p.rows.forEach(r => counts[getStatus(r)]++);
    const positions = p.rows
      .sort((a,b)=>columnNumber(a.column)-columnNumber(b.column)||rowNumber(a.row)-rowNumber(b.row))
      .map(r => `${r.position} — ${getStatus(r)}`);

    return `
      <div class="card">
        <div class="dbname">${esc(p.name)}</div>
        ${p.category ? `<span class="pill">${esc(p.category)}</span>` : ""}
        <div>
          ${counts.Complete ? `<span class="dbstatus">✓ ${counts.Complete} Complete</span>` : ""}
          ${counts["Not Complete"] ? `<span class="dbstatus">${counts["Not Complete"]} Not Complete</span>` : ""}
          ${counts.Missing ? `<span class="dbstatus">${counts.Missing} Missing</span>` : ""}
        </div>
        <div class="dbpositions">${positions.map(esc).join("<br>")}</div>
      </div>`;
  }).join("") || '<div class="card">No database matches.</div>';
}

$("dbSearch").addEventListener("input", renderDatabase);
$("dbCategory").addEventListener("change", renderDatabase);
$("dbColumn").addEventListener("change", renderDatabase);
$("dbStatus").addEventListener("change", renderDatabase);
$("dbSort").addEventListener("change", renderDatabase);

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
