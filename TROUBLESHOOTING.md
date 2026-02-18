# TROUBLESHOOTING

This file covers common PoC issues in **Word on the web**.

---

## 1) RichApi.Error: “The argument you provided is not valid”

Most often caused by invalid highlight colors (or other invalid property values). Word highlight colors are constrained; use standard **OpenXML highlight names** (lowercase) such as `red`, `yellow`, `green`. citeturn20search69turn20search70

**Fix:**
- Ensure your code uses: `red`, `yellow`, `green`
- Log debug info: `console.error(e, e.debugInfo)` citeturn20search70

---

## 2) Template blocks stay Yellow after editing

If you edit text inside a TEMPLATE content control, the baseline hash in the tag must be recomputed and updated.

See **AUTHORING.md**.

---

## 3) Word on the web caching (changes not showing)

If you change `taskpane.js` but Word still runs an older version:

```html
<script src="./taskpane.js?v=YYYYMMDD-N" defer></script>
```

Then reload the Word tab or close/open the document.

---

## 4) “SourceLocation must be HTTPS” / add-in won’t load

Office add-in SourceLocation must be HTTPS. citeturn3search1

GitHub Pages is HTTPS; confirm your manifest points to:

- `https://luisgomezcalaveritas-create.github.io/contract-addin-poc/taskpane.html` citeturn3search7turn3search1

---

## 5) Content control properties not editable

Word on the web may not provide the full content control Properties UI; author templates in Word Desktop. citeturn26search78turn26search81
