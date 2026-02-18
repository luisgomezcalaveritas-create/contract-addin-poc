/* Contract Drafting PoC - GitHub Pages hosted task pane */

const CLAUSE_INDEX_URL = "./clauses.json"; // relative to taskpane.html on GitHub Pages

// ---------- UI helpers ----------
function setStatus(msg, cls) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = cls || "small";
}
function setStatus2(msg, cls) {
  const el = document.getElementById("status2");
  el.textContent = msg || "";
  el.className = cls || "small";
}

function normalizeText(s) {
  return (s || "").replace(/\s+/g, " ").trim();
}

async function sha256Hex(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}

function parseBaselineHashFromTag(tag) {
  // expected last segment: h<hash>
  const parts = (tag || "").split("|");
  const last = parts[parts.length - 1] || "";
  return last.startsWith("h") ? last.slice(1) : "";
}

// ---------- Fetch helpers ----------
async function fetchJson(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return res.json();
}

async function fetchAsBase64(url) {
  // Fetch DOCX as arrayBuffer -> base64
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  const buffer = await res.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  // Convert to base64 safely (no Node)
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

// ---------- Clause library ----------
let clauseIndex = [];

async function loadClauses() {
  setStatus("Loading clauses…", "small");
  clauseIndex = await fetchJson(CLAUSE_INDEX_URL);

  const sel = document.getElementById("clauseSelect");
  sel.innerHTML = "";
  sel.appendChild(new Option("(Select a clause)", ""));

  for (const c of clauseIndex) {
    const label = `${c.clauseId} v${c.version} — ${c.title}`;
    sel.appendChild(new Option(label, `${c.clauseId}|${c.version}`));
  }

  setStatus("Clauses loaded ✅", "ok");
  setStatus2(`Found ${clauseIndex.length} clause(s) in clauses.json`, "small");
}

function getSelectedClause() {
  const sel = document.getElementById("clauseSelect");
  const val = sel.value;
  if (!val) return null;

  const [clauseId, version] = val.split("|");
  return clauseIndex.find(c => c.clauseId === clauseId && String(c.version) === String(version)) || null;
}

// ---------- Word actions ----------
async function insertSelectedClause() {
  const chosen = getSelectedClause();
  if (!chosen) {
    setStatus("Select a clause first.", "warn");
    return;
  }

  setStatus(`Fetching ${chosen.docxFile}…`, "small");

  // Files are relative to docs/ root
  const meta = await fetchJson(`./${chosen.metaFile}`);
  const base64Docx = await fetchAsBase64(`./${chosen.docxFile}`);

  setStatus("Inserting into Word…", "small");

  await Word.run(async (context) => {
    const sel = context.document.getSelection();

    // Create placeholder and wrap in content control
    const p = sel.insertParagraph("[[CLAUSE]]", Word.InsertLocation.replace);
    const cc = p.getRange().insertContentControl();

    cc.title = meta.title || chosen.title;
    cc.tag = `APPROVED|${meta.clauseId}|v${meta.version}|h${meta.baselineHash}`;
    cc.appearance = "BoundingBox";

    // Insert DOCX snippet into CC
    cc.getRange().insertFileFromBase64(base64Docx, Word.InsertLocation.replace);

    // mark green initially
    cc.getRange().font.highlightColor = "Green";

    await context.sync();
  });

  setStatus(`Inserted ${chosen.clauseId} v${chosen.version} ✅`, "ok");
}

// ---------- Validation (Red / Yellow / Green) ----------
async function validateDocument() {
  setStatus("Validating…", "small");

  await Word.run(async (context) => {
    const doc = context.document;

    // 1) Default everything to Red (custom)
    doc.body.font.highlightColor = "Red";

    // 2) Load content controls
    const ccs = doc.contentControls;
    ccs.load("items/tag,text");
    await context.sync();

    for (const cc of ccs.items) {
      const tag = cc.tag || "";
      const isStandard = tag.startsWith("TEMPLATE|") || tag.startsWith("APPROVED|");
      if (!isStandard) continue;

      const baselineHash = parseBaselineHashFromTag(tag);
      const currentText = normalizeText(cc.text || "");
      const currentHash = await sha256Hex(currentText);

      cc.getRange().font.highlightColor = (currentHash === baselineHash) ? "Green" : "Yellow";
    }

    await context.sync();
  });

  setStatus("Validation complete ✅", "ok");
  setStatus2("Red=custom, Green=unchanged standard, Yellow=edited standard", "small");
}

// ---------- Boot ----------
function wireUi() {
  document.getElementById("btnLoad").addEventListener("click", () => {
    loadClauses().catch(err => setStatus(`Load error: ${err.message}`, "err"));
  });

  document.getElementById("btnInsert").addEventListener("click", () => {
    insertSelectedClause().catch(err => setStatus(`Insert error: ${err.message}`, "err"));
  });

  document.getElementById("btnValidate").addEventListener("click", () => {
    validateDocument().catch(err => setStatus(`Validate error: ${err.message}`, "err"));
  });
}

(function init() {
  // Always show something (prevents "blank page" confusion)
  setStatus("JS loaded ✅ (waiting for Office host)", "small");

  wireUi();

  if (typeof Office === "undefined") {
    setStatus("This page is meant to run inside Word (Office host not detected).", "warn");
    setStatus2("Tip: Sideload the manifest in Word for the web to test.", "small");
    return;
  }

  Office.onReady(() => {
    setStatus("Office.onReady ✅ (running inside Word)", "ok");
    setStatus2("Click 'Load clauses' to populate the dropdown.", "small");
  });
})();
