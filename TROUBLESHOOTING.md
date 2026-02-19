# TROUBLESHOOTING

## 1) Office.js did not load (Office is undefined)

Symptoms:
- Task pane shows: “Office.js did not load. Check Network for office.js.”

Checks:
- Confirm `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` is reachable.
- Verify Office.js is referenced in `<head>`.
- In Word web DevTools, inspect the **task pane iframe** (not the top frame) and check Network.

Fixes:
- Try InPrivate window (extensions off).
- Allowlist Microsoft Office.js CDN in corporate networks if blocked.

## 2) Word on the web caching (changes not showing)

Fix:
- Bump `window.BUILD_VERSION` in `docs/taskpane.html`.
- Match `taskpane.html?v=...` in `manifest.xml`.
- Close Word tab fully and reopen.

## 3) Validate error: “Cannot read properties of null (reading 'clone')”

Cause:
- Word on the web can throw internal errors when calling `range.getTrackedChanges()` across multiple ranges.

Fix:
- Use `contentControl.getTrackedChanges()` for tracked changes retrieval.

## 4) Insert failed: ooxmlIsMalformated

Cause:
- Base64 doc is not a real DOCX or contains unsupported content.

Fix:
- Confirm the DOCX URL serves a real DOCX (not a GitHub blob page or LFS pointer).
- Re-export the snippet as a clean DOCX if needed.

## 5) Reset clears user highlights too

Current PoC behavior:
- Reset clears all highlights in the document.

If you need to preserve user highlights:
- Implement an “add-in-only highlight tagging” strategy (future enhancement).
