/*
 * Nexl Check Bridge — content script. Runs only inside the Nexl Check add-in's task pane frame
 * and relays its requests to the background worker (which holds the Nexl permission).
 */
(function () {
  const VERSION = chrome.runtime.getManifest().version;
  const reply = (data) => window.postMessage(Object.assign({ source: "nexl-check-bridge" }, data), window.location.origin);

  window.addEventListener("message", (e) => {
    if (e.source !== window || e.origin !== window.location.origin) return;
    const d = e.data;
    if (!d || d.source !== "nexl-check-addin" || !d.id) return;
    try {
      chrome.runtime.sendMessage({ type: d.type, path: d.path }, (resp) => {
        const err = chrome.runtime.lastError;
        reply(Object.assign({ id: d.id }, err ? { ok: false, error: "BRIDGE_ERROR", detail: err.message } : resp));
      });
    } catch (err) {
      // Happens after the extension is reloaded/updated: the page needs a refresh.
      reply({ id: d.id, ok: false, error: "BRIDGE_RELOADED", detail: "The Nexl Check Bridge was updated. Reload Excel." });
    }
  });

  reply({ type: "hello", version: VERSION });
})();
