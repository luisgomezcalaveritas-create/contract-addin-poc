(function () {
  const status = document.getElementById("status");
  const status2 = document.getElementById("status2");

  function set1(msg, cls) {
    if (status) {
      status.textContent = msg;
      status.className = cls || "small";
    }
    console.log("[status]", msg);
  }

  function set2(msg, cls) {
    if (status2) {
      status2.textContent = msg || "";
      status2.className = cls || "small";
    }
    console.log("[detail]", msg);
  }

  // Always show something so the page isn't "blank" in a normal browser tab.
  set1("taskpane.js loaded ✅", "ok");

  // If Office.js didn't load, stop here with a clear message.
  if (typeof Office === "undefined") {
    set1("Office is undefined ❌", "err");
    set2("Office.js did not load. Check Network for office.js.", "small");
    return;
  }

  set1("Office detected ✅ (waiting for Office.onReady)", "ok");

  // If Office.onReady never fires, show a timeout message.
  const timeout = setTimeout(() => {
    set1("Office.onReady did NOT fire (timeout) ❌", "err");
    set2("If this is inside Word, check console errors or tenant restrictions.", "small");
  }, 8000);

  Office.onReady()
    .then(() => {
      clearTimeout(timeout);
      set1("Office.onReady ✅ (running inside Word)", "ok");
      set2("Click Validate to insert a test paragraph.", "small");

      const btn = document.getElementById("btnValidate");
      if (!btn) {
        set2("Button not found. Check that taskpane.html has id='btnValidate'.", "warn");
        return;
      }

      btn.addEventListener("click", async () => {
        set2("Validate clicked…", "small");
        try {
          await Word.run(async (context) => {
            context.document.body.insertParagraph(
              "Validate clicked (debug).",
              Word.InsertLocation.end
            );
            await context.sync();
          });
          set2("Inserted paragraph ✅", "ok");
        } catch (e) {
          set2("Word.run failed ❌ " + (e?.message || String(e)), "err");
          console.error(e);
        }
      });
    })
    .catch((err) => {
      clearTimeout(timeout);
      set1("Office.onReady error ❌", "err");
      set2(err?.message || String(err), "err");
      console.error(err);
    });
})();
