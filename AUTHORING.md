# AUTHORING (Template Content Controls + Baseline Hashes)

This guide is for template authors who maintain **SalesContractTemplate_PoC.docx** or any future template.

Content controls are a strong fit for templates because they can be labeled (Title/Tag), bounded, and optionally restricted from deletion/editing. citeturn26search84turn26search79

---

## 1) Enable Developer tab (Word Desktop)

1. Word **File → Options**
2. **Customize Ribbon**
3. Check **Developer**
4. OK citeturn26search78turn26search81

> Word on the web may not expose all content control property editing. Use Word Desktop for authoring. citeturn26search78turn26search81

---

## 2) Select the correct content control

**Tip:** Turn on **Developer → Design Mode** to make selection easier. citeturn26search81turn26search84

1. Click inside the target block (heading/paragraph/table).
2. If needed, enable **Design Mode**.
3. Click the control boundary so the entire control is selected.

---

## 3) Open Properties and set Title/Tag

1. With the control selected: **Developer → Properties** citeturn26search78turn26search81
2. Update:

### Title (human-friendly)
Examples:
- `GoverningLaw`
- `Notices`
- `Pricing`

### Tag (machine-readable)
Use the schema:

- `TEMPLATE|HD|<Name>|h<64-hex>`
- `TEMPLATE|TB|<Name>|h<64-hex>`
- `TEMPLATE|BP|<Name>|h<64-hex>`

Example:

```
TEMPLATE|BP|GoverningLaw|h0123...abcd   (64 hex total)
```

### Optional settings in Properties
- **Show as**: choose **Bounding box** for easy PoC visualization. citeturn26search84
- **Locking**:
  - ✅ Content control cannot be deleted
  - (optional) ✅ Contents cannot be edited (strict governance) citeturn26search79turn26search86

Click **OK**, then **Save**.

---

## 4) Recompute baseline hash after editing template text

Whenever you change text inside a TEMPLATE block, update the hash in the tag.

### Recommended PoC approach: compute hash with the add-in
Add an **Authoring** button in the task pane:
- **Compute hash for selected Content Control**

Workflow:
1. Edit the text in the content control.
2. Click inside that same content control.
3. Click **Compute hash**.
4. Copy the new hash (64 hex).
5. Open **Developer → Properties** and replace the trailing `h<oldhash>` with `h<newhash>`.
6. Save the template.

---

## 5) Updating the “NEGOTIATED CLAUSES” section

Two patterns:

### Option A (recommended for demos)
Wrap the instruction paragraph in a TEMPLATE boilerplate control so it validates **Green** when unchanged:

- Title: `NegotiatedClausesInstruction`
- Tag: `TEMPLATE|BP|NegotiatedClausesInstruction|h<hash>`

### Option B (strict)
Leave the instruction paragraph outside any content control so it remains **Red**.

---

## 6) Quick verification

1. Open template in Word.
2. Run **Validate**.
3. Unchanged TEMPLATE blocks should be **Green**.
4. Edited TEMPLATE blocks should be **Yellow** until you update their hash.
