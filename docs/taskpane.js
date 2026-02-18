(function () {
  const status = document.getElementById("status");

  // Prove JS is running (even in a normal browser)
  if (status) status.textContent = "JS loaded ✅";

  // If Office isn't available (normal browser tab), stop here.
  // This prevents a blank page or errors.
  if (typeof Office === "undefined") {
    if (status) status.textContent += " — Office host not detected (expected in browser).";
    return;
  }

  // Only runs inside Word (Office host)
  Office.onReady(() => {
    if (status) status.textContent = "Office.onReady ✅ (running inside Word)";

    const btn = document.getElementById("validate");
    if (btn) {
      btn.addEventListener("click", async () => {
        await Word.run(async (context) => {
          context.document.body.insertParagraph(
            "Validate clicked (PoC).",
            Word.InsertLocation.end
          );
          await context.sync();
        });
      });
    }
  });
})();
