(function () {
  const CLAUSE_INDEX_URL =
    "https://luisgomezcalaveritas-create.github.io/contract-addin-poc/clauses.json";

  const BUILD_VERSION = (window.BUILD_VERSION || "dev");
  console.log("Contract PoC build:", BUILD_VERSION);

  // OpenXML highlight values (lowercase)
  const HIGHLIGHT = {
    RED: "red",
    YELLOW: "yellow",
    GREEN: "green"
  };

  // UI
  const elStatus = document.getElementById("status");
  const elStatus2 = document.getElementById("status2");
  const elSearch = document.getElementById("search");
  const elResults = document.getElementById("results");
  const btnValidate = document.getElementById("btnValidate");
  const btnReset = document.getElementById("btnReset");

  const elBadge = document.getElementById("statusBadge");
  const elStatusText = document.getElementById("statusText");

  let clauses = [];
  let indexBaseUrl = CLAUSE_INDEX_URL;

  function set1(msg, cls) {
    if (elStatus) {
      elStatus.textContent = msg;
      elStatus.className = cls || "small";
    }
    console.log("[status]", msg);
  }

  function set2(msg, cls) {
    if (elStatus2) elStatus2.className = cls || "small status2";
    if (elStatusText) elStatusText.textContent = msg || "";
    console.log("[detail]", msg);
  }

  function setBadge(label, variant, text) {
    if (elBadge) {
      elBadge.textContent = label || "";
      elBadge.className = `badge badge--${variant || "neutral"}`;
    }
    const buildSuffix = `Build ${BUILD_VERSION}`;
    const combined = text ? `${text} • ${buildSuffix}` : buildSuffix;
    set2(combined, "small status2");
  }

  function logOfficeError(prefix, e) {
    const msg = e?.debugInfo?.message || e?.message || String(e);
    console.error(prefix, msg);
    console.error("Full error:", e);
    console.error("Office debugInfo:", e?.debugInfo);
    return msg;
  }

  function escapeHtml(str) {
    return (str || "").replace(/[&<>"']/g, (m) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;"
    }[m]));
  }

  function resolveUrl(baseUrl, maybeRelativeUrl) {
    try { return new URL(maybeRelativeUrl, baseUrl).toString(); }
    catch { return maybeRelativeUrl; }
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

  function parseExpectedHashFromTag(tag) {
    const parts = (tag || "").split("|");
    const last = parts[parts.length - 1] || "";
    return last.startsWith("h") ? last.slice(1) : "";
  }

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

  async function ensureTrackAllEnabled() {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
    });
  }

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
      li.innerHTML =
        `<div><b>No results</b></div>` +
        `<div class="meta">Try searching: nda, risk, term, payment…</div>`;
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

  async function loadClauses() {
    set1("Loading clause index…", "ok");
    setBadge("Initializing", "neutral", "Loading approved clauses…");

    const index = await fetchJson(CLAUSE_INDEX_URL);
    const list = Array.isArray(index) ? index : (index.clauses || []);
    indexBaseUrl = CLAUSE_INDEX_URL;

    clauses = list.map(r => {
      const c = normalizeClauseRecord(r);
      c.clauseJsonUrl = resolveUrl(indexBaseUrl, c.clauseJsonUrl);
      c.clauseDocxUrl = resolveUrl(indexBaseUrl, c.clauseDocxUrl);
      return c;
    });

    set1(`Loaded ${clauses.length} clauses ✅`, "ok");
    setBadge("Track Changes ON", "positive", "Search and click a clause to insert.");

    elSearch.disabled = false;
    btnValidate.disabled = false;
    btnReset.disabled = false;

    renderList(elSearch.value || "");
  }

  async function insertClause(c) {
    if (!c.approved) {
      set1("Insertion blocked", "warn");
      setBadge("Review Needed", "notice", "This clause is not approved.");
      return;
    }

    try {
      await ensureTrackAllEnabled();
      setBadge("Inserting", "info", `Inserting ${c.clauseId}…`);

      set1(`Downloading metadata: ${c.clauseId}…`, "ok");
      const meta = await fetchJson(c.clauseJsonUrl);
      const baselineHash = (meta.baselineHash || "").trim();
      if (!baselineHash) throw new Error("Clause metadata missing baselineHash.");

      set1(`Downloading DOCX: ${c.clauseId}…`, "ok");
      const base64Docx = await fetchBase64(c.clauseDocxUrl);

      set1(`Inserting: ${c.clauseId}…`, "ok");

      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const insertedRange = selection.insertFileFromBase64(base64Docx, Word.InsertLocation.replace);

        const cc = insertedRange.insertContentControl();
        cc.title = `${c.title} (${c.clauseId} ${c.version})`;
        cc.tag = `APPROVED|${c.clauseId}|${c.version}|h${baselineHash}`;
        cc.appearance = "BoundingBox";

        await context.sync();
      });

      set1(`Inserted ${c.clauseId} ✅`, "ok");
      setBadge("Track Changes ON", "positive", "Inserted. Edit text (tracked), then Validate.");
    } catch (e) {
      const msg = logOfficeError("Insert failed", e);
      set1("Insert failed ❌", "err");
      setBadge("Error", "negative", msg);
    }
  }

  // ✅ RESET: remove the traffic-light highlight colors applied during validation.
  // Office.js supports removing highlight by setting highlightColor = null. [1](https://github.com/OfficeDev/office-js/issues/4638)
  async function resetHighlights() {
    try {
      set1("Resetting highlights…", "ok");
      setBadge("Resetting", "info", "Clearing highlight colors…");

      await Word.run(async (context) => {
        const bodyRange = context.document.body.getRange();
        bodyRange.font.highlightColor = null; // remove highlight [1](https://github.com/OfficeDev/office-js/issues/4638)
        await context.sync();
      });

      set1("Reset complete ✅", "ok");
      setBadge("Track Changes ON", "positive", "Highlights cleared. You can Validate again anytime.");
    } catch (e) {
      const msg = logOfficeError("Reset failed", e);
      set1("Reset failed ❌", "err");
      setBadge("Error", "negative", msg);
    }
  }

  async function validateDocument() {
    try {
      set1("Validating…", "ok");
      setBadge("Validating", "info", "Green baseline + Yellow tracked insertions…");

      await ensureTrackAllEnabled();

      const trackedApiSupported = Office.context.requirements.isSetSupported("WordApi", "1.6");
      if (!trackedApiSupported) {
        await validateDocumentHashOnlyFallback();
        return;
      }

      const snapshot = await Word.run(async (context) => {
        context.document.body.getRange().font.highlightColor = HIGHLIGHT.RED;

        const controls = context.document.contentControls;
        controls.load("items/tag");
        await context.sync();

        const rows = [];
        for (const cc of controls.items) {
          const tag = cc.tag || "";
          if (!tag.startsWith("TEMPLATE|") && !tag.startsWith("APPROVED|")) continue;

          const range = cc.getRange();
          range.load("text");

          const tcs = range.getTrackedChanges();
          tcs.load("items/type");

          rows.push({ tag, range, tcs });
        }

        await context.sync();

        return rows.map(r => {
          const insertionCount = (r.tcs.items || [])
            .filter(tc => String(tc.type || "").toLowerCase().includes("insertion"))
            .length;
          return { tag: r.tag, text: r.range.text || "", insertionCount };
        });
      });

      const decisionsByTag = new Map();
      let okCount = 0;
      let changedCount = 0;
      let fallbackCount = 0;

      for (const item of snapshot) {
        const expected = parseExpectedHashFromTag(item.tag);
        const currentHash = await sha256Hex(normalizeText(item.text));
        const isMatch = expected && currentHash === expected;

        if (isMatch) okCount++;
        else changedCount++;

        const needsFallbackYellow = (!isMatch && item.insertionCount === 0);
        if (needsFallbackYellow) fallbackCount++;

        decisionsByTag.set(item.tag, {
          insertionCount: item.insertionCount,
          needsFallbackYellow
        });
      }

      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load("items/tag");
        await context.sync();

        const overlayBuckets = [];
        for (const cc of controls.items) {
          const tag = cc.tag || "";
          if (!tag.startsWith("TEMPLATE|") && !tag.startsWith("APPROVED|")) continue;

          const decision = decisionsByTag.get(tag) || { insertionCount: 0, needsFallbackYellow: false };
          const range = cc.getRange();

          range.font.highlightColor = decision.needsFallbackYellow ? HIGHLIGHT.YELLOW : HIGHLIGHT.GREEN;

          if (decision.insertionCount > 0) {
            const tcs = range.getTrackedChanges();
            tcs.load("items/type");
            overlayBuckets.push(tcs);
          }
        }

        await context.sync();

        for (const tcs of overlayBuckets) {
          for (const tc of tcs.items) {
            const typeStr = String(tc.type || "").toLowerCase();
            if (typeStr.includes("insertion")) {
              tc.getRange().font.highlightColor = HIGHLIGHT.YELLOW;
            }
          }
        }

        await context.sync();
      });

      if (changedCount === 0) {
        set1("Validation complete ✅", "ok");
        setBadge("Validated", "positive", "All standard blocks match baseline.");
      } else if (fallbackCount > 0) {
        set1("Validation complete ✅", "warn");
        setBadge("Review Needed", "notice",
          `Some blocks differ but have no tracked insertions (possibly accepted). Green=${okCount}, Changed=${changedCount}, Fallback=${fallbackCount}.`);
      } else {
        set1("Validation complete ✅", "ok");
        setBadge("Validated", "positive",
          `Yellow marks only inserted/changed visible text. Green=${okCount}, Changed=${changedCount}.`);
      }
    } catch (e) {
      const msg = logOfficeError("Validate failed", e);
      set1("Validate failed ❌", "err");
      setBadge("Error", "negative", msg);
    }
  }

  async function validateDocumentHashOnlyFallback() {
    try {
      set1("Validating (fallback)…", "warn");
      setBadge("Review Needed", "notice", "Tracked changes API not available. Using block-level Green/Yellow.");

      const snapshot = await Word.run(async (context) => {
        context.document.body.getRange().font.highlightColor = HIGHLIGHT.RED;

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
        return temp.map(t => ({ tag: t.tag, text: t.range.text }));
      });

      const decisions = [];
      for (const item of snapshot) {
        const expected = parseExpectedHashFromTag(item.tag);
        const currentHash = await sha256Hex(normalizeText(item.text));
        decisions.push({
          tag: item.tag,
          highlight: (expected && currentHash === expected) ? HIGHLIGHT.GREEN : HIGHLIGHT.YELLOW
        });
      }

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

      set1("Validation complete ✅ (fallback)", "warn");
      setBadge("Review Needed", "notice", "Fallback applied: Green=match, Yellow=changed, Red=outside blocks.");
    } catch (e) {
      const msg = logOfficeError("Validate fallback failed", e);
      set1("Validate failed ❌", "err");
      setBadge("Error", "negative", msg);
    }
  }

  // Boot
  set1("taskpane.js loaded ✅", "ok");
  setBadge("Initializing", "neutral", "Waiting for Word…");

  if (typeof Office === "undefined") {
    set1("Office is undefined ❌", "err");
    setBadge("Error", "negative", "Office.js did not load. Check Network for office.js.");
    return;
  }

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Word) {
      set1("Loaded, but not running in Word", "warn");
      setBadge("Review Needed", "notice", `Detected host: ${info.host}. This add-in expects Word.`);
      return;
    }

    set1("Running inside Word ✅", "ok");
    setBadge("Initializing", "neutral", "Enabling Track Changes (track everyone)…");

    try {
      await ensureTrackAllEnabled();
      setBadge("Track Changes ON", "positive", "Tracking everyone’s changes.");
    } catch (e) {
      console.error("Failed to enable Track Changes:", e, e?.debugInfo);
      setBadge("Track Changes OFF", "negative", "Could not enable Track Changes.");
    }

    elSearch.addEventListener("input", () => renderList(elSearch.value));
    btnValidate.onclick = validateDocument;
    btnReset.onclick = resetHighlights;

    try {
      await loadClauses();
    } catch (e) {
      set1("Failed to load clauses ❌", "err");
      setBadge("Error", "negative", e?.message || String(e));
      console.error(e);
    }
  });
})();
