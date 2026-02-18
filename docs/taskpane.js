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

  // Use the info argument so we can verify the host.
  Office.onReady((info) => {
    clearTimeout(timeout);

    if (info.host !== Office.HostType.Word) {
      set1("Office.onReady ✅ but host is not Word", "warn");
      set2(`Detected host: ${info.host}. This add-in expects Word.`, "warn");
      return;
    }

    set1("Office.onReady ✅ (running inside Word)", "ok");
    set2("Click Validate to insert a test paragraph.", "small");

    const btn = document.getElementById("btnValidate");
    if (!btn) {
      set2("Button not found. Check that taskpane.html has id='btnValidate'.", "warn");
      return;
    }

    // IMPORTANT: enable the button (you set it disabled in HTML).
    btn.disabled = false;

    // Idempotent wiring: assign onclick rather than stacking listeners.
    btn.onclick = async () => {
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
        const msg =
          e?.debugInfo?.message ||
          e?.message ||
          String(e);
        set2("Word.run failed ❌ " + msg, "err");
        console.error(e);
      }
    };
  });
})();
