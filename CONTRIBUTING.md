# CONTRIBUTING

Thanks for contributing to the Contract Drafting PoC.

## 1) Branching model

- `main` — demo-ready
- `feature/<name>` — feature work

## 2) PR checklist

- [ ] `manifest.xml` version bumped (a.b.c.d) when behavior changes.
- [ ] `docs/taskpane.html` `window.BUILD_VERSION` bumped.
- [ ] `manifest.xml` `taskpane.html?v=...` query matches build version.
- [ ] Insert tested (at least one clause inserts correctly).
- [ ] Validate tested (Red/Green + Yellow insertions overlay).
- [ ] Reset tested (clears highlights).
- [ ] No blocking console errors.

## 3) Coding standards (Office.js)

- Use `Office.onReady` before Word APIs.
- Use `Word.run` batching and avoid extra `context.sync()` in tight loops.
- Log `error.debugInfo` on failures; it usually identifies invalid statements.

## 4) Release tags

Use Git tags for demo milestones:
- `demo-YYYY-MM-DD`
