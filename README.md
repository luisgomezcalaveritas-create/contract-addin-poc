# Contract Drafting PoC (Word on the web + GitHub Pages)

This repository contains a Proof of Concept (PoC) **Word (Office.js) task pane add-in** for contract drafting:

- Loads an approved clause index from **GitHub Pages** (`clauses.json`).
- Provides a searchable clause picker and inserts selected clauses as **DOCX snippets**.
- Wraps inserted clauses in **Content Controls** tagged with baseline SHA-256 hashes.
- Uses **Track Changes (trackAll)** and a **traffic-light overlay**:
  - **Red** = outside standard blocks (custom by default)
  - **Green** = standard baseline text inside `TEMPLATE|` / `APPROVED|` blocks
  - **Yellow** = *only* visible inserted/changed text (tracked insertions) inside standard blocks
- Includes **Reset** to clear highlight coloring applied during Validate.

> Office.js is referenced from Microsoft's CDN in the `<head>` of `taskpane.html` (required for reliable initialization).

---

## 1) Live URLs (GitHub Pages)

- Base: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/`
- Task pane: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/taskpane.html`
- Script: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/taskpane.js`
- Clause index: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/clauses.json`

---

## 2) Repository structure (recommended)

```
contract-addin-poc/
  manifest.xml              # sideload this (local upload in Word on the web)
  README.md
  AUTHORING.md
  CONTRIBUTING.md
  TROUBLESHOOTING.md
  docs/                     # GitHub Pages root
    taskpane.html
    taskpane.js
    clauses.json
    *_v1.json
    *_v1.docx
    icon-v104.png
    SalesContractTemplate_PoC.docx
```

**Why manifest at repo root?** The manifest is uploaded locally for sideloading; GitHub Pages hosts only the taskpane assets.

---

## 3) Sideload in Word on the web

1. Open **Word on the web** and open/create a document.
2. **Insert → Add-ins → Upload My Add-in**.
3. Select `manifest.xml` from your local machine.

> Note: Word on the web stores sideloaded add-ins in browser storage. If you clear site data or switch browsers, you may need to sideload again.

---

## 4) Manifest configuration essentials

- `SourceLocation` must be **HTTPS** and should include a cache-busting version query.
  - Example: `taskpane.html?v=1.0.0.4`
- Add the GitHub Pages host domain under `<AppDomains>`.
- Icons should use stable file names (or versioned file names like `icon-v104.png`) to avoid UI caching issues.

---

## 5) Build version + cache busting (important)

Word on the web caches aggressively.

### Single build version
We standardize on **one build version** (e.g., `1.0.0.4`) used in:
- `manifest.xml` `<Version>` and `taskpane.html?v=...`
- `docs/taskpane.html` `window.BUILD_VERSION = "..."`
- `docs/taskpane.js` logs `Contract PoC build: ...` and shows `Build ...` in the task pane status.

### How to bump
For each deploy:
1. Update `window.BUILD_VERSION` in `docs/taskpane.html`.
2. Update `manifest.xml` `<Version>` and `taskpane.html?v=...` to match.
3. Commit + push.
4. Reopen Word tab and re-run add-in (and re-sideload manifest if needed).

---

## 6) Clause library format

### `clauses.json`
Portable format with relative URLs recommended:

```json
{
  "schemaVersion": "1.0",
  "clauses": [
    {
      "clauseId": "LOL",
      "title": "Limitation of Liability",
      "version": "v1",
      "approved": true,
      "category": "Risk Allocation",
      "tags": ["liability", "risk", "damages"],
      "clauseJsonUrl": "./LOL_v1.json",
      "clauseDocxUrl": "./LOL_v1.docx"
    }
  ]
}
```

### Clause metadata `*_v1.json`
Each clause metadata JSON must contain a baseline hash:

```json
{
  "clauseId": "LOL",
  "version": "v1",
  "title": "Limitation of Liability",
  "approved": true,
  "baselineHash": "<64-hex-sha256>"
}
```

---

## 7) Tagging scheme

### Template blocks
Template content controls are tagged:

- `TEMPLATE|HD|<Name>|h<hash>` (headings)
- `TEMPLATE|TB|<Name>|h<hash>` (tables)
- `TEMPLATE|BP|<Name>|h<hash>` (boilerplate blocks)

### Inserted approved clauses
Inserted clauses are wrapped in a content control tagged:

- `APPROVED|<ClauseId>|<Version>|h<hash>`

---

## 8) Validate + Reset behavior

### Validate
- Sets entire document highlight to **Red**.
- For each `TEMPLATE|` and `APPROVED|` content control:
  - base highlight **Green** (or **Yellow** fallback)
  - overlays **Yellow** only on **tracked insertions** (visible changed/added text)
- Deletions remain visible as tracked deletions (no highlight needed).

### Reset
- Clears highlight coloring by setting `font.highlightColor = null` on the body range.

---

## 9) Quick demo script (5 minutes)

1. Open template doc.
2. Insert a clause from the list.
3. Edit inserted clause (Track Changes is ON).
4. Click **Validate**.
   - standard text is Green
   - only inserted/changed visible text becomes Yellow
   - anything outside standard blocks remains Red
5. Click **Reset** to clear coloring.

---

## 10) Further docs

- **AUTHORING.md** — template content controls, hashes, and re-baselining
- **CONTRIBUTING.md** — branching, PR checklist, coding standards
- **TROUBLESHOOTING.md** — common issues (Office.js load, caching, tracked-changes quirks)
