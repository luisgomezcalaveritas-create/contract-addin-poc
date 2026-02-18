# Contract Drafting PoC (Word on the web + GitHub Pages)

This repository contains a Proof of Concept (PoC) **Word (Office.js) task pane add-in** for contract drafting:

- Loads an approved clause index from **GitHub Pages** (`clauses.json`).
- Provides a **searchable clause picker** and inserts selected clauses as **DOCX snippets**.
- Wraps inserted content in **Word Content Controls** and tags them with a baseline SHA‑256 hash.
- Runs **Validate** to apply a **traffic‑light** model:
  - **Red**: custom text (outside standard blocks)
  - **Green**: standard block unchanged vs baseline
  - **Yellow**: standard block edited vs baseline

> **Why Content Controls?** Content controls are bounded containers designed for structured documents and templates (including restricting edits/deletions and adding semantic identifiers). citeturn26search84turn26search79

---

## 1. Live URLs (GitHub Pages)

- **Base URL**: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/` citeturn3search7
- **Task pane**: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/taskpane.html`
- **Script**: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/taskpane.js`
- **Clause index**: `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/clauses.json`

> GitHub Pages is HTTPS by default; Office add-ins require HTTPS for add-in pages (SourceLocation). citeturn3search1turn3search7

---

## 2. Repository structure

```
contract-addin-poc/
  manifest.xml
  README.md
  AUTHORING.md
  CONTRIBUTING.md
  TROUBLESHOOTING.md
  docs/
    taskpane.html
    taskpane.js
    clauses.json
    CONF_v1.json
    CONF_v1.docx
    LOL_v1.json
    LOL_v1.docx
    PAY_v1.json
    PAY_v1.docx
    TERM_v1.json
    TERM_v1.docx
    TERMNT_v1.json
    TERMNT_v1.docx
    SalesContractTemplate_PoC.docx
```

---

## 3. Sideload the add-in in **Word on the web**

### 3.1 Prerequisites
- A Microsoft 365 account with access to **Word on the web**.
- `manifest.xml` downloaded locally.

### 3.2 Steps (manual sideload)
1. Open **Word on the web** and open/create a document.
2. Go to **Insert → Add-ins**.
3. Choose **Upload My Add-in**.
4. Select `manifest.xml`.

> Sideloading on Office on the web stores the manifest in the browser’s local storage; switching browsers or clearing storage requires sideloading again. citeturn3search13

---

## 4. Manifest configuration (critical)

### 4.1 SourceLocation must be HTTPS
In `manifest.xml`, ensure:

```xml
<SourceLocation DefaultValue="https://luisgomezcalaveritas-create.github.io/contract-addin-poc/taskpane.html"/>
```

Office requires **HTTPS** for SourceLocation. citeturn3search1

### 4.2 Trust domains (AppDomains)
If you use additional domains, list them in `<AppDomains>`. For GitHub Pages:

```xml
<AppDomains>
  <AppDomain>https://luisgomezcalaveritas-create.github.io</AppDomain>
</AppDomains>
```

The `AppDomain` element declares additional trusted domains used by the add-in. citeturn3search5

### 4.3 Icons
Use HTTPS URLs for `IconUrl` and `HighResolutionIconUrl` (or remove icons for PoC).

---

## 5. Clause library format

### 5.1 `clauses.json`
Recommended `clauses.json` supports **portable relative URLs** and **tags** (improves search):

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

### 5.2 Clause metadata (`*_v1.json`)
Each clause metadata JSON must include a **baselineHash**:

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

## 6. Tagging scheme (Template + Inserted Clauses)

### 6.1 Template blocks
Template content controls are tagged:

- `TEMPLATE|HD|<Name>|h<hash>` (headings)
- `TEMPLATE|TB|<Name>|h<hash>` (tables)
- `TEMPLATE|BP|<Name>|h<hash>` (boilerplate blocks)

### 6.2 Inserted clauses
Inserted clauses are wrapped in a content control tagged:

- `APPROVED|<ClauseId>|<Version>|h<hash>`

The add-in inserts a DOCX snippet and then wraps it in a content control for validation.

---

## 7. Validation (Traffic Light)

### 7.1 Rules
Validate does:

1. Paint the **whole document** as **Red**.
2. For each content control with tag starting with `TEMPLATE|` or `APPROVED|`:
   - compute `sha256(normalize(text))`
   - compare to `h<hash>` in the tag
   - match → **Green**
   - mismatch → **Yellow**
3. Anything outside standard content controls remains **Red**.

### 7.2 Highlight color compatibility (Word on the web)
Word highlight colors are constrained; use standard **OpenXML highlight names** in lowercase (e.g., `red`, `yellow`, `green`). citeturn20search69turn20search70

---

## 8. Template authoring workflow (iteration)

When you edit text **inside a TEMPLATE content control**, you must:

1) Edit content
2) Recompute baseline hash
3) Update the content control **Tag**
4) Save the template

See **AUTHORING.md** for detailed steps, including Developer → Properties (Title/Tag) and a recommended in‑add‑in hash helper.

---

## 9. Development notes

### 9.1 Office.js loading
The task pane loads Office.js from Microsoft’s CDN. This is the recommended distribution mechanism. citeturn10search49

### 9.2 Word API batching
Use `Word.run` batching correctly and avoid excessive `context.sync()` in loops; prefer split‑loop / correlated objects patterns. citeturn20search61turn20search63

---

## 10. Quick demo script (5 minutes)

1. Open `SalesContractTemplate_PoC.docx`
2. Open the add-in pane → click **Validate**
   - template blocks should become **Green**
3. Search and click a clause (e.g., `LOL`) to insert
4. Edit a sentence inside the inserted clause
5. Click **Validate** again
   - edited clause becomes **Yellow**
6. Type a paragraph outside any content control
   - stays **Red**

---

## 11. Further docs

- **AUTHORING.md** — how to edit content controls and recompute baseline hashes
- **CONTRIBUTING.md** — branching, PR checklist, coding standards
- **TROUBLESHOOTING.md** — common issues and fixes (invalid argument, caching, etc.)
