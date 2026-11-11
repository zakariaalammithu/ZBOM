// web-online.js — ONLINE-ONLY patch for the web copy (web/renderer.js).
// desktop-client/renderer.js is untouched; this file does not exist there.
//
// Why it exists: renderer.js searches Fabric Material Code client-side
// over state.snapshot.materials (renderMaterials -> filterRows). On the
// web the full 103,710-row master must never be sent to the browser, so
// this script intercepts the #materialSearch input, asks
// /api/materials?q=... (server-side search over the whole master), puts
// the returned page into state.snapshot.materials, and re-renders.
// `state` and `renderMaterials` are top-level bindings of renderer.js, so
// a classic script loaded after it can use them directly.
//
// Behaviour:
//   - Typing 2+ characters searches the whole master (debounced 350 ms);
//     the renderer's own 180 ms client-side filter runs first on the old
//     window, so there can be a brief "0 matching records" flash before
//     the server results arrive — the totals below are then corrected.
//   - Clearing the box reloads the first 50 rows (the default window).
//   - The true match total from the API is shown in #materialCount and
//     #materialMatches, because renderer.js only knows the loaded window.
//   - A failed request leaves the previously loaded rows in place.
"use strict";

(function () {
  if (typeof state === "undefined" || typeof renderMaterials !== "function") return;
  var input = document.querySelector("#materialSearch");
  if (!input) return;

  var seq = 0;
  var timer = null;

  function showTotals(total) {
    var count = document.querySelector("#materialCount");
    if (count) count.textContent = Number(total || 0).toLocaleString();
    var matches = document.querySelector("#materialMatches");
    if (matches) matches.textContent = Number(total || 0).toLocaleString() + " matching records";
  }

  async function runSearch() {
    var query = input.value.trim();
    var mine = ++seq;
    var url = "/api/materials?q=" + encodeURIComponent(query) + "&limit=50";
    try {
      var res = await fetch(url, { headers: { Accept: "application/json" } });
      if (!res.ok) return;
      var data = await res.json();
      if (mine !== seq) return; // a newer keystroke already fired
      state.snapshot.materials = Array.isArray(data.rows) ? data.rows : [];
      renderMaterials();
      showTotals(data.total);
    } catch (error) {
      // Keep the previously loaded rows; the renderer already shows them.
    }
  }

  input.addEventListener("input", function () {
    window.clearTimeout(timer);
    timer = window.setTimeout(runSearch, 350);
  });
})();
