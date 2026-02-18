(function () {
  const status = document.getElementById("status");
  const status2 = document.getElementById("status2");

  const set1 = (m) => { status.textContent = m; console.log("[status]", m); };
  const set2 = (m) => { status2.textContent = m || ""; console.log("[detail]", m); };

  set1("taskpane.js loaded ✅");

  if (typeof Office === "undefined") {
    set1("Office is undefined ❌");
    set2("Office.js did not load. Check Network for office.js.");
    return;
  }

  set1("Office detected ✅ (waiting for Office.onReady)");

  Office.onReady().then(() => {
    set1("Office.onReady ✅ (running inside Word)");
    set2("Click Validate to insert a test paragraph.");

    document.getElementById("btnValidate").addEventListener("click", async () => {
      set2("Validate clicked…");
      await Word.run(async (context) => {
        context.document.body.insertParagraph("Validate clicked (debug).", Word.InsertLocation.end);
        await context.sync();
      });
      set2("Inserted paragraph ✅");
    });
  }).catch((err) => {
    set1("Office.onReady error ❌");
    set2(err?.message || String(err));
    console.error(err);
  });
})();
