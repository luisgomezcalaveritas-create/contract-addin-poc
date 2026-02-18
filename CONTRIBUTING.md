# CONTRIBUTING

Thanks for contributing to the Contract Drafting PoC.

---

## 1) Branching model

- `main` — stable demo-ready branch
- feature branches — `feature/<short-name>`
  - examples: `feature/hash-button`, `feature/template-authoring-ui`

---

## 2) Pull request checklist

Before you open a PR:

- [ ] `taskpane.html` script query string bumped (e.g., `taskpane.js?v=YYYYMMDD-N`) to avoid Word web caching issues.
- [ ] `clauses.json` validated (JSON syntax + URLs resolve correctly).
- [ ] Insert flow tested: at least one clause inserts successfully.
- [ ] Validate flow tested: TEMPLATE + APPROVED controls colorize correctly.
- [ ] Console clean of **RichApi.Error** (warnings from Word web telemetry can be ignored).

---

## 3) Coding standards

- Use `Office.onReady` before any Word APIs.
- Use `Word.run` batching for document changes. citeturn10search47turn20search74
- Avoid `context.sync()` inside loops; use split-loop/correlated patterns for performance. citeturn20search61turn20search63
- Log Office errors with `e.debugInfo` when available; it often identifies invalid arguments. citeturn20search70

---

## 4) Release / Demo tagging

Use Git tags for demo milestones:
- `demo-YYYY-MM-DD`

Optionally keep a `CHANGELOG.md` once iteration speeds up.
