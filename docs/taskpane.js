(function () {
  /**
   * Hardcoded clause index URL (as requested)
   */
  const CLAUSE_INDEX_URL =
    "https://luisgomezcalaveritas-create.github.io/contract-addin-poc/clauses.json";

  /**
   * IMPORTANT: Use OpenXML highlight values (lowercase).
   * These are widely supported: red, yellow, green, etc. [1](https://thelinuxcode.com/host-a-website-on-github-for-free-a-practical-modern-guide-2026/)
   * Avoid nonstandard values like "BrightGreen" which can cause invalid-argument errors. [2](https://www.hostragons.com/en/blog/free-static-website-hosting-with-github-pages/)
   */
  const HIGHLIGHT = {
    RED: "red",
    YELLOW: "yellow",
    GREEN: "green"
  };

  // --- UI elements ---
  const elStatus = document.getElementById("status");
  const elStatus2 = document.getElementById("status2");
  const elSearch = document.getElementById("search");
  const elResults = document.getElementById("results");
  const btnValidate = document.getElementById("btnValidate");
  const btnReload = document.getElementById("btnReload");

  // --- State ---
  let clauses = [];
  let indexBaseUrl = CLAUSE_INDEX_URL;

  // ---------------------------
  // Status helpers
  // ---------------------------
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

  function logOfficeError(prefix, e) {
    const msg = e?.debugInfo?.message || e?.message || String(e);
    console.error(prefix, msg);
    console.error("Full error:", e);
    console.error("Office debugInfo:", e?.debugInfo); // Office tells you to inspect debugInfo for invalid args [2](https://www.hostragons.com/en/blog/free-static-website-hosting-with-github-pages/)
    return msg;
  }

  function escapeHtml(str) {
    return (str || "").replace(/[&<>"']/g, (m) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[m]));
  }

  // ---------------------------
  // URL + fetch helpers
  // ---------------------------
  function resolveUrl(baseUrl, maybeRelativeUrl) {
    try {
      return new URL(maybeRelativeUrl, baseUrl).toString();
    } catch {
      return maybeRelativeUrl;
    }
  }

  async function fetchJson(url) {
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
    const buf = await blob.arrayBuffer();
    return arrayBufferToBase64(buf);
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

  // ---------------------------
  // Hash helpers
  // ---------------------------
  function normalizeText(s) {
    return (s || "")
      .replace(/\r\n/g, "\n")
      .replace(/\u00A0/g, " ")
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

  // ---------------------------
  // Data normalization
  // ---------------------------
  function normalizeClauseRecord(r) {
    const clauseId = r.clauseId || r.id || "";
    return {
      clauseId,
      title: r.title || r.name || clauseId || "Untitled clause",
      version: r.version || "v1",
      approved: (r.approved !== undefined) ? !!r.approved : true,
      category: r.category || "",
      tags: Array.isArray(r.tags) ? r.tags : [],
      clauseJsonUrl: r.clauseJsonUrl || r.metaUrl || r.jsonUrl || "",
      clauseDocxUrl: r.clauseDocxUrl || r.docxUrl || ""
    };
  }

  // ---------------------------
  // UI rendering
  // ---------------------------
  function renderList(filterText) {
    const q = (filterText || "").trim().toLowerCase();
    const filtered = clauses.filter(c => {
      const hay = `${c.clauseId} ${c.title} ${c.version} ${c.category} ${(c.tags || []).join(" ")}`
        .toLowerCase();
      return hay.includes(q);
    });

    elResults.innerHTML = "";

    if (!filtered.length) {
      const li = document.createElement("li");
      li.style.cursor = "default";
      li.innerHTML = `<div><b>No results</b></div><div class="meta">Try searching: nda, risk, term, payment…</div>`;
      elResults.appendChild(li);
      return;
    }

    filtered.forEach(c => {
      const li = document.createElement("li");
      const pill = c.approved
        ? `<span class="pill ok">approved</span>`
        : `<span class="pill warn">not approved</span>`;

      const catLine = c.category
        ? `<div class="meta">category: ${escapeHtml(c.category)}</div>`
        : "";

      const tagLine = (c.tags && c.tags.length)
        ? `<div class="meta">tags: ${escapeHtml(c.tags.slice(0, 6).join(", "))}${c.tags.length > 6 ? "…" : ""}</div>`
        : "";

      li.innerHTML = `
        <div><b>${escapeHtml(c.title)}</b>${pill}</div>
        <div class="meta">${escapeHtml(c.clauseId)} • ${escapeHtml(c.version)}</div>
        ${catLine}
        ${tagLine}
      `;

      li.onclick = () => insertClause(c);
      elResults.appendChild(li);
    });
  }

  // ---------------------------
  // Load clauses
  // ---------------------------
  async function loadClauses() {
    set1("Loading clause index…", "ok");
    set2(CLAUSE_INDEX_URL, "small");

    const index = await fetchJson(CLAUSE_INDEX_URL);
    const list = Array.isArray(index) ? index : (index.clauses || []);
    indexBaseUrl = CLAUSE_INDEX_URL;

    clauses = list.map(r => {
      const c = normalizeClauseRecord(r);
      c.clauseJsonUrl = resolveUrl(indexBaseUrl, c.clauseJsonUrl);
      c.clauseDocxUrl = resolveUrl(indexBaseUrl, c.clauseDocxUrl);
      return c;
    });

    const missing = clauses.filter(c => !c.clauseJsonUrl || !c.clauseDocxUrl);
    if (missing.length) {
      set1("Loaded, but some clauses are missing URLs", "warn");
      set2("Missing URLs for: " + missing.map(m => m.clauseId || m.title).join(", "), "warn");
    } else {
      set1(`Loaded ${clauses.length} clauses ✅`, "ok");
      set2("Search and click a clause to insert.", "small");
    }

    elSearch.disabled = false;
    btnValidate.disabled = false;
    btnReload.disabled = false;

    renderList(elSearch.value || "");
  }

  // ---------------------------
  // Insert clause (DOCX + content control tagging)
  // ---------------------------
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

      const baselineHash = (meta.baselineHash || "").trim();
      if (!baselineHash) {
        set1("Cannot insert", "err");
        set2("Clause metadata missing baselineHash.", "err");
        return;
      }

      set1(`Downloading DOCX: ${c.clauseId}…`, "ok");
      const base64Docx = await fetchBase64(c.clauseDocxUrl);

      set1(`Inserting: ${c.clauseId}…`, "ok");

      // Word.run is the standard batching pattern for Word add-ins [6](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest)[5](https://codesandbox.io/examples/package/office-addin-taskpane-js)
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        const insertedRange = selection.insertFileFromBase64(
          base64Docx,
          Word.InsertLocation.replace
        );

        const cc = insertedRange.insertContentControl();
        cc.title = `${c.title} (${c.clauseId} ${c.version})`;
        cc.tag = `APPROVED|${c.clauseId}|${c.version}|h${baselineHash}`;
        cc.appearance = "BoundingBox";

        // Apply initial highlight using OpenXML highlight name (lowercase). [1](https://thelinuxcode.com/host-a-website-on-github-for-free-a-practical-modern-guide-2026/)
        cc.getRange().font.highlightColor = HIGHLIGHT.GREEN;

        await context.sync();
      });

      set1(`Inserted ${c.clauseId} ✅`, "ok");
      set2("Edit inside the clause then click Validate to see Yellow.", "small");
    } catch (e) {
      const msg = logOfficeError("Insert failed", e);
      set1("Insert failed ❌", "err");
      set2(msg, "err");
    }
  }

  // ---------------------------
  // Validate traffic lights (two-pass: read -> hash -> write)
  // This avoids async crypto work inside the Word batch and aligns with batching guidance. [3](https://stackoverflow.com/questions/40639456/office-addin-manifest)[4](https://learn.microsoft.com/en-us/javascript/api/manifest/appdomain?view=word-js-preview)
  // ---------------------------
  async function validateDocument() {
    try {
      set1("Validating…", "ok");
      set2("Painting document Red, then checking content controls.", "small");

      // PASS 1: Paint body red + read tags + read text from relevant content controls
      const snapshot = await Word.run(async (context) => {
        const bodyRange = context.document.body.getRange();
        bodyRange.font.highlightColor = HIGHLIGHT.RED;

        const controls = context.document.contentControls;
        controls.load("items/tag");
        await context.sync();

        const temp = [];
        for (const cc of controls.items) {
          const tag = cc.tag || "";
          if (!tag.startsWith("TEMPLATE|") && !tag.startsWith("APPROVED|")) continue;
          const range = cc.getRange();
          range.load("text");
          temp.push({ tag, range });
        }

        await context.sync();

        // Return plain JS (no Word objects)
        return temp.map(t => ({ tag: t.tag, text: t.range.text }));
      });

      // PASS 1.5: Hash outside Word.run
      const decisions = [];
      for (const item of snapshot) {
        const parts = (item.tag || "").split("|");
        const last = parts[parts.length - 1] || "";
        const expected = last.startsWith("h") ? last.slice(1) : "";

        const currentHash = await sha256Hex(normalizeText(item.text));
        const ok = expected && currentHash === expected;

        decisions.push({
          tag: item.tag,
          highlight: ok ? HIGHLIGHT.GREEN : HIGHLIGHT.YELLOW
        });
      }

      // PASS 2: Apply green/yellow highlights back in Word
      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load("items/tag");
        await context.sync();

        for (const cc of controls.items) {
          const tag = cc.tag || "";
          if (!tag.startsWith("TEMPLATE|") && !tag.startsWith("APPROVED|")) continue;

          const match = decisions.find(d => d.tag === tag);
          if (!match) continue;

          cc.getRange().font.highlightColor = match.highlight;
        }

        await context.sync();
      });

      set1("Validation complete ✅", "ok");
      set2("Green=match, Yellow=changed, Red=outside standard blocks.", "small");
    } catch (e) {
      const msg = logOfficeError("Validate failed", e);
      set1("Validate failed ❌", "err");
      set2(msg, "err");
    }
  }

  // ---------------------------
  // Boot
  // ---------------------------
  set1("taskpane.js loaded ✅", "ok");

  // Office.js should be referenced from the Microsoft CDN for add-ins [7](https://code.visualstudio.com/docs/other/office)
  if (typeof Office === "undefined") {
    set1("Office is undefined ❌", "err");
    set2("Office.js did not load. Check Network for office.js.", "small");
    return;
  }

  // Office.onReady ensures the host is initialized before calling Word APIs [5](https://codesandbox.io/examples/package/office-addin-taskpane-js)[6](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest)
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

    // Load clauses
    btnReload.disabled = true;
    try {
      await loadClauses();
    } catch (e) {
      const msg = e?.message || String(e);
      set1("Failed to load clauses ❌", "err");
      set2(msg, "err");
      console.error(e);
      btnReload.disabled = false;
    }
  });
})();
