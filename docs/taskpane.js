(function () {
  const status = document.getElementById("status");
  const status2 = document.getElementById("status2");

  const set1 = (m, cls) => { status.textContent = m; status.className = cls || "small"; console.log("[status]", m); };
  const set2 = (m, cls) => { status2.textContent = m || ""; status2.className = cls || "small"; console.log("[detail]", m); };

  // Prove JS ran
  set1("taskpane.js loaded ✅", "ok");

  // Prove Office.js is available
  if (typeof Office === "undefined") {
    set1("Office is undefined ❌", "err");
    set2("Office.js did not load. Check taskpane.html script tag and Network tab.", "small");
    return;
  }

  set1("Office detected ✅ (waiting for Office.onReady)", "ok");

  const timeout = setTimeout(() => {
    set1("Office.onReady did NOT fire (timeout) ❌", "err");
    set2("If Office is detected but onReady doesn't fire, check console errors.", "small");
  }, 8000);

  Office.onReady().then(() => {
    clearTimeout(timeout);
    set1("Office.onReady ✅ (running inside Word)", "ok");
    set2("Click Validate to insert a test paragraph.", "small");

    document.getElementById("btnValidate").addEventListener("click", async () => {
      set2("Validate clicked…", "small");
      try {
        await Word.run(async (context) => {
          context.document.body.insertParagraph("Validate clicked (debug).", Word.InsertLocation.end);
          await context.sync();
        });
        set2("Inserted paragraph ✅", "ok");
      } catch (e) {
        set2("Word.run failed ❌ " + (e?.message || e), "err");
        console.error(e);
      }
    });
  }).catch(err => {
    clearTimeout(timeout);
    set1("Office.onReady error ❌", "err");
    set2(err?.message || String(err), "err");
    console.error(err);
  });
})();
