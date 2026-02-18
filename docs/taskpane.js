(function () {
  const CLAUSE_INDEX_URL =
    "https://luisgomezcalaveritas-create.github.io/contract-addin-poc/clauses.json";

  // OpenXML highlight values (lowercase) - safest across Word on the web.
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
  const btnReload = document.getElementById("btnReload");

  let clauses = [];
  let indexBaseUrl = CLAUSE_INDEX_URL;

  // Track Changes UX (persistent prefix on status2)
  let trackPrefix = "Track Changes: (checking…)";
  let detailSuffix = "";

  function set1(msg, cls) {
    if (elStatus) {
      elStatus.textContent = msg;
      elStatus.className = cls || "small";
    }
    console.log("[status]", msg);
  }

  function set2(msg, cls) {
    detailSuffix = msg || "";
    renderStatus2(cls || "small");
    console.log("[detail]", msg);
  }

  function setTrackPrefix(msg, cls) {
    trackPrefix = msg || "Track Changes: (unknown)";
    renderStatus2(cls || "small");
  }

  function renderStatus2(cls) {
    if (!elStatus2) return;
    // Keep Track Changes indicator always visible.
    // Append short secondary detail if present.
    const combined = detailSuffix
      ? `${trackPrefix} • ${detailSuffix}`
      : `${trackPrefix}`;
    elStatus2.textContent = combined;
    elStatus2.className = cls || "small";
  }

  function logOfficeError(prefix, e) {
    const msg = e?.debugInfo?.message || e?.message || String(e);
    console.error(prefix, msg);
    console.error("Full error:", e);
    console.error("Office debugInfo:", e?.debugInfo);
    return msg;
  }

  function escapeHtml(str) {
    // Correct escaping for dynamic HTML strings
    return (str || "").replace(/[&<>"']/g, (m) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "\"": "&quot;",
      "'": "&#39;"
    }[m]));
  }

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
    set2("Loading approved clauses…", "small");

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
    set2("Search and click a clause to insert.", "small");

    elSearch.disabled = false;
    btnValidate.disabled = false;
    btnReload.disabled = false;

    renderList(elSearch.value || "");
  }

  /**
   * Turn on Track Changes for everyone and leave it ON.
   * Uses Word.ChangeTrackingMode.trackAll. [1](https://stackoverflow.com/questions/79562806/is-it-possible-to-sideload-office-web-extension-manifest-xml-in-production-mode)[2](https://www.udacity.com/blog/2025/08/how-to-host-your-website-for-free-using-github-pages-a-step-by-step-guide.html)
   */
  async function ensureTrackAllEnabled() {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
    });
  }

  async function insertClause(c) {
    if (!c.approved) {
      set1("Insertion blocked", "warn");
      set2("This clause is not approved.", "warn");
      return;
    }

    try {
      // Ensure Track Changes stays ON even if user toggled it off.
      await ensureTrackAllEnabled();
      setTrackPrefix("Track Changes: ON (Track everyone)", "ok");

      set1(`Downloading metadata: ${c.clauseId}…`, "ok");
      set2("Fetching clause metadata…", "small");

      const meta = await fetchJson(c.clauseJsonUrl);
      const baselineHash = (meta.baselineHash || "").trim();
      if (!baselineHash) throw new Error("Clause metadata missing baselineHash.");

      set1(`Downloading DOCX: ${c.clauseId}…`, "ok");
      set2("Fetching DOCX snippet…", "small");

      const base64Docx = await fetchBase64(c.clauseDocxUrl);

      set1(`Inserting: ${c.clauseId}…`, "ok");
      set2("Inserting clause into document…", "small");

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
      set2("Edit text (tracked). Click Validate to mark only inserted/changed text Yellow.", "small");
    } catch (e) {
      const msg = logOfficeError("Insert failed", e);
      set1("Insert failed ❌", "err");
      set2(msg, "err");
    }
  }

  /**
   * Validate with partial-yellow behavior:
   * - Paint whole document RED (custom by default)
   * - Standard blocks (TEMPLATE| / APPROVED|) set to GREEN
   * - Only tracked INSERTIONS inside those blocks set to YELLOW
   * - Deletions remain visible as Track Changes deletions (no highlight)
   * - Fallback: if hash != baseline but no tracked insertions exist, set whole block YELLOW.
   *
   * Tracked changes APIs are WordApi 1.6+. [3](https://www.hostragons.com/en/blog/free-static-website-hosting-with-github-pages/)[1](https://stackoverflow.com/questions/79562806/is-it-possible-to-sideload-office-web-extension-manifest-xml-in-production-mode)
   */
  async function validateDocument() {
    try {
      set1("Validating…", "ok");
      set2("Green standard blocks; Yellow only inserted/changed text; Red elsewhere.", "small");

      // Ensure Track Changes is ON.
      await ensureTrackAllEnabled();
      setTrackPrefix("Track Changes: ON (Track everyone)", "ok");

      const trackedApiSupported = Office.context.requirements.isSetSupported("WordApi", "1.6");
      if (!trackedApiSupported) {
        // If tracked changes APIs aren't available, fall back to old behavior.
        await validateDocumentHashOnlyFallback();
        return;
      }

      // PASS 1: paint body RED and snapshot tags + texts + presence of insertion tracked changes
      const snapshot = await Word.run(async (context) => {
        // 1) Whole document red
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

          // Track changes in this range
          // Range.getTrackedChanges is used widely; we load type of each item. [3](https://www.hostragons.com/en/blog/free-static-website-hosting-with-github-pages/)[4](https://github.com/OfficeDev/office-js/issues/6514)
          const tcs = range.getTrackedChanges();
          tcs.load("items/type");

          rows.push({ tag, range, tcs });
        }

        await context.sync();

        // Return plain data
        return rows.map(r => {
          const insertCount = (r.tcs.items || [])
            .filter(tc => String(tc.type || "").toLowerCase().includes("insertion"))
            .length;

          return {
            tag: r.tag,
            text: r.range.text || "",
            insertionCount: insertCount
          };
        });
      });

      // PASS 1.5: compute hash decisions outside Word.run
      const decisionsByTag = new Map();
      let fallbackCount = 0;
      let changedCount = 0;
      let okCount = 0;

      for (const item of snapshot) {
        const expected = parseExpectedHashFromTag(item.tag);
        const currentHash = await sha256Hex(normalizeText(item.text));
        const isMatch = expected && currentHash === expected;

        if (isMatch) okCount++;
        else changedCount++;

        // Fallback rule: changed hash AND no insertion revisions found → mark whole block Yellow.
        const needsFallbackYellow = (!isMatch && item.insertionCount === 0);

        if (needsFallbackYellow) fallbackCount++;

        decisionsByTag.set(item.tag, {
          isMatch,
          insertionCount: item.insertionCount,
          needsFallbackYellow
        });
      }

      // PASS 2: apply GREEN for blocks, then overlay YELLOW insertions (and fallback YELLOW block if needed)
      await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load("items/tag");
        await context.sync();

        // Paint each standard block green or yellow (fallback), then overlay insertion ranges yellow.
        for (const cc of controls.items) {
          const tag = cc.tag || "";
          if (!tag.startsWith("TEMPLATE|") && !tag.startsWith("APPROVED|")) continue;

          const decision = decisionsByTag.get(tag);
          const range = cc.getRange();

          // Base color:
          // - If hash matches: Green
          // - If hash mismatch but no insertion tracked changes: Yellow fallback
          // - Else: Green base, and insertions will be painted yellow below
          if (!decision) {
            range.font.highlightColor = HIGHLIGHT.GREEN;
            continue;
          }

          if (decision.needsFallbackYellow) {
            range.font.highlightColor = HIGHLIGHT.YELLOW;
          } else {
            range.font.highlightColor = HIGHLIGHT.GREEN;
          }

          // Overlay insertions: only when there are insertion revisions.
          if (decision.insertionCount > 0) {
            const tcs = range.getTrackedChanges();
            tcs.load("items/type");
            await context.sync();

            for (const tc of tcs.items) {
              const typeStr = String(tc.type || "").toLowerCase();
              // Only highlight visible inserted/changed text.
              if (typeStr.includes("insertion")) {
                tc.getRange().font.highlightColor = HIGHLIGHT.YELLOW;
              }
              // Deletions are shown as tracked deletions; no highlight.
            }
          }
        }

        await context.sync();
      });

      // UX summary
      if (changedCount === 0) {
        set1("Validation complete ✅", "ok");
        set2("All standard blocks match baseline (Green).", "ok");
      } else if (fallbackCount > 0) {
        set1("Validation complete ✅ (with fallback)", "warn");
        set2(
          `Green blocks: ${okCount}. Changed blocks: ${changedCount}. ` +
          `Fallback Yellow blocks (no insertion revisions found): ${fallbackCount} (changes may have been accepted).`,
          "warn"
        );
      } else {
        set1("Validation complete ✅", "ok");
        set2(
          `Green blocks: ${okCount}. Changed blocks: ${changedCount}. ` +
          `Yellow marks only inserted/changed text (tracked insertions).`,
          "small"
        );
      }
    } catch (e) {
      const msg = logOfficeError("Validate failed", e);
      set1("Validate failed ❌", "err");
      set2(msg, "err");
    }
  }

  /**
   * Fallback validator (no tracked-changes API): keep original hash-only block behavior.
   */
  async function validateDocumentHashOnlyFallback() {
    try {
      set1("Validating (hash-only fallback)…", "warn");
      set2("Tracked changes API not available; using block-level Green/Yellow.", "warn");

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
      set2("Green=match, Yellow=changed, Red=outside standard blocks.", "warn");
    } catch (e) {
      const msg = logOfficeError("Validate fallback failed", e);
      set1("Validate failed ❌", "err");
      set2(msg, "err");
    }
  }

  // Boot
  set1("taskpane.js loaded ✅", "ok");

  if (typeof Office === "undefined") {
    set1("Office is undefined ❌", "err");
    set2("Office.js did not load. Check Network for office.js.", "err");
    return;
  }

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Word) {
      set1("Loaded, but not running in Word", "warn");
      set2(`Detected host: ${info.host}`, "warn");
      return;
    }

    set1("Running inside Word ✅", "ok");
    setTrackPrefix("Track Changes: (enabling…)", "small");
    set2("Initializing…", "small");

    // Turn on Track Changes for everyone and keep it on.
    try {
      await ensureTrackAllEnabled();
      setTrackPrefix("Track Changes: ON (Track everyone)", "ok");
    } catch (e) {
      console.error("Failed to enable Track Changes:", e, e?.debugInfo);
      setTrackPrefix("Track Changes: (could not enable)", "warn");
      set2("Partial-yellow requires Track Changes. You may still validate with fallback.", "warn");
    }

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
      btnReload.disabled = false;
    } catch (e) {
      set1("Failed to load clauses ❌", "err");
      set2(e?.message || String(e), "err");
      console.error(e);
      btnReload.disabled = false;
    }
  });
})();
