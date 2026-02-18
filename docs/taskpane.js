(function () {
  // HARD-CODED clause index URL (as requested)
  const CLAUSE_INDEX_URL = "https://luisgomezcalaveritas-create.github.io/contract-addin-poc/clauses.json";

  // Highlight colors (Word supports a fixed palette; these are commonly available)
  const HIGHLIGHT = {
    RED: "Red",
    YELLOW: "Yellow",
    GREEN: "BrightGreen"
  };

  const elStatus = document.getElementById("status");
  const elStatus2 = document.getElementById("status2");
  const elSearch = document.getElementById("search");
  const elResults = document.getElementById("results");
  const btnValidate = document.getElementById("btnValidate");
  const btnReload = document.getElementById("btnReload");

  let clauses = [];        // normalized list
  let lastFilter = "";

  function set1(msg, cls) {
    if (elStatus) {
      elStatus.textContent = msg;
      elStatus.className = cls || "small";
    }
    console.log("[status]", msg);
  }

  function set2(msg, cls) {
    if (elStatus2) {
      elStatus2.textContent = msg || "";
      elStatus2.className = cls || "small";
    }
    console.log("[detail]", msg);
  }

  function escapeHtml(str) {
    return (str || "").replace(/[&<>"']/g, (m) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[m]));
  }

  function normalizeClauseRecord(r) {
    return {
      clauseId: r.clauseId || r.id || "",
      title: r.title || r.name || r.clauseId || "Untitled clause",
      version: r.version || "v1",
      approved: (r.approved !== undefined) ? !!r.approved : true,
      clauseJsonUrl: r.clauseJsonUrl || r.metaUrl || r.jsonUrl || "",
      clauseDocxUrl: r.clauseDocxUrl || r.docxUrl || ""
    };
  }

  async function fetchJson(url) {
    // cache-bust to avoid GitHub Pages caching surprises during iteration
    const bust = (url.includes("?") ? "&" : "?") + "cb=" + Date.now();
    const res = await fetch(url + bust, { method: "GET" });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
    return res.json();
  }

  async function fetchBase64(url) {
    const bust = (url.includes("?") ? "&" : "?") + "cb=" + Date.now();
    const res = await fetch(url + bust, { method: "GET" });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);

    const blob = await res.blob();
    const arrayBuffer = await blob.arrayBuffer();
    return arrayBufferToBase64(arrayBuffer);
  }

  function arrayBufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
    }
    return btoa(binary);
  }

  function normalizeText(s) {
    return (s || "")
      .replace(/\r\n/g, "\n")
      .replace(/\u00A0/g, " ")  // NBSP → space
      .replace(/[ \t]+/g, " ")
      .replace(/\n[ \t]+/g, "\n")
      .trim();
  }

  async function sha256Hex(text) {
    const enc = new TextEncoder();
    const data = enc.encode(text);
    const digest = await crypto.subtle.digest("SHA-256", data);
    return [...new Uint8Array(digest)]
      .map(b => b.toString(16).padStart(2, "0"))
      .join("");
  }

  function renderList(filterText) {
    lastFilter = (filterText || "").trim().toLowerCase();
    const q = lastFilter;

    const filtered = clauses.filter(c => {
      const hay = `${c.clauseId} ${c.title} ${c.version}`.toLowerCase();
      return hay.includes(q);
    });

    elResults.innerHTML = "";

    if (filtered.length === 0) {
      const li = document.createElement("li");
      li.innerHTML = `<div><b>No results</b></div><div class="meta">Try a different search term.</div>`;
      li.style.cursor = "default";
      elResults.appendChild(li);
      return;
    }

    filtered.forEach(c => {
      const li = document.createElement("li");
      const pill = c.approved
        ? `<span class="pill ok">approved</span>`
        : `<span class="pill warn">not approved</span>`;

      li.innerHTML = `
        <div><b>${escapeHtml(c.title)}</b>${pill}</div>
        <div class="meta">${escapeHtml(c.clauseId)} • ${escapeHtml(c.version)}</div>
      `;

      li.onclick = () => insertClause(c);
      elResults.appendChild(li);
    });
  }

  async function loadClauses() {
    set1("Loading clause index…", "ok");
    set2(CLAUSE_INDEX_URL, "small");

    const index = await fetchJson(CLAUSE_INDEX_URL);
    const list = Array.isArray(index) ? index : (index.clauses || []);

    clauses = list.map(normalizeClauseRecord);

    // Basic validation of data
    const missing = clauses.filter(c => !c.clauseJsonUrl || !c.clauseDocxUrl);
    if (missing.length) {
      set1("Loaded, but some clauses are missing URLs", "warn");
      set2(`Missing URLs for: ${missing.map(m => m.clauseId || m.title).join(", ")}`, "warn");
    } else {
      set1(`Loaded ${clauses.length} clauses ✅`, "ok");
      set2("Search and click a clause to insert.", "small");
    }

    // Enable UI
    elSearch.disabled = false;
    btnValidate.disabled = false;
    btnReload.disabled = false;

    renderList(elSearch.value || "");
  }

  async function insertClause(c) {
    if (!c.approved) {
      set1("Insertion blocked", "warn");
      set2("This clause is not approved.", "warn");
      return;
    }
    if (!c.clauseJsonUrl || !c.clauseDocxUrl) {
      set1("Cannot insert", "err");
      set2("Clause record missing clauseJsonUrl or clauseDocxUrl.", "err");
      return;
    }

    try {
      set1(`Downloading metadata: ${c.clauseId}…`, "ok");
      const meta = await fetchJson(c.clauseJsonUrl);

      const baselineHash = meta.baselineHash || meta.hash || "";
      if (!baselineHash) {
        set1("Cannot insert", "err");
        set2("Clause metadata missing baselineHash.", "err");
        return;
      }

      set1(`Downloading DOCX: ${c.clauseId}…`, "ok");
      const base64Docx = await fetchBase64(c.clauseDocxUrl);

      set1(`Inserting: ${c.clauseId}…`, "ok");
      await Word.run(async (context) => { // standard Word add-in pattern [2](https://learn.microsoft.com/en-us/javascript/api/manifest/appdomain?view=word-js-preview)
        const selection = context.document.getSelection();

        const insertedRange = selection.insertFileFromBase64(
          base64Docx,
          Word.InsertLocation.replace
        );

        const cc = insertedRange.insertContentControl();
        cc.title = `${c.title} (${c.clauseId} ${c.version})`;
        cc.tag = `APPROVED|${c.clauseId}|${c.version}|h${baselineHash}`;
        cc.appearance = "BoundingBox";

        // Optional: initial highlight as Green
        cc.getRange().font.highlightColor = HIGHLIGHT.GREEN;

        await context.sync();
      });

      set1(`Inserted ${c.clauseId} ✅`, "ok");
      set2("Edit inside the clause then click Validate to see Yellow.", "small");
    } catch (e) {
      const msg = e?.debugInfo?.message || e?.message || String(e);
      set1("Insert failed ❌", "err");
      set2(msg, "err");
      console.error(e);
    }
  }

  async function validateDocument() {
    try {
      set1("Validating…", "ok");
      set2("Painting document Red, then checking content controls.", "small");

      await Word.run(async (context) => { // standard Word add-in pattern [2](https://learn.microsoft.com/en-us/javascript/api/manifest/appdomain?view=word-js-preview)
        const bodyRange = context.document.body.getRange();
        bodyRange.font.highlightColor = HIGHLIGHT.RED;

        const controls = context.document.contentControls;
        controls.load("items/tag,title");
        await context.sync();

        for (const cc of controls.items) {
          const tag = cc.tag || "";
          if (!tag.startsWith("TEMPLATE|") && !tag.startsWith("APPROVED|")) continue;

          const parts = tag.split("|");
          const last = parts[parts.length - 1] || "";
          const expected = last.startsWith("h") ? last.slice(1) : "";

          const range = cc.getRange();
          range.load("text");
          await context.sync();

          const currentHash = await sha256Hex(normalizeText(range.text));

          range.font.highlightColor = (expected && currentHash === expected)
            ? HIGHLIGHT.GREEN
            : HIGHLIGHT.YELLOW;
        }

        await context.sync();
      });

      set1("Validation complete ✅", "ok");
      set2("Green=match, Yellow=changed, Red=outside standard blocks.", "small");
    } catch (e) {
      const msg = e?.debugInfo?.message || e?.message || String(e);
      set1("Validate failed ❌", "err");
      set2(msg, "err");
      console.error(e);
    }
  }

  // Boot
  set1("taskpane.js loaded ✅", "ok");

  if (typeof Office === "undefined") {
    set1("Office is undefined ❌", "err");
    set2("Office.js did not load. Check Network for office.js.", "small");
    return;
  }

  // Office.onReady ensures we only touch Word APIs when the host is ready. [2](https://learn.microsoft.com/en-us/javascript/api/manifest/appdomain?view=word-js-preview)
  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Word) {
      set1("Loaded, but not running in Word", "warn");
      set2(`Detected host: ${info.host}`, "warn");
      return;
    }

    set1("Running inside Word ✅", "ok");
    set2("Loading approved clauses…", "small");

    // Wire UI
    elSearch.addEventListener("input", () => renderList(elSearch.value));
    btnValidate.onclick = validateDocument;
    btnReload.onclick = async () => {
      btnReload.disabled = true;
      try { await loadClauses(); }
      finally { btnReload.disabled = false; }
    };

    // Load index immediately
    try {
      await loadClauses();
    } catch (e) {
      const msg = e?.message || String(e);
      set1("Failed to load clauses ❌", "err");
      set2(msg, "err");
      console.error(e);
    }
  });
})();
