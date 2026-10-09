/*
 * Nexl Check — talks to the "Nexl Check Bridge" browser extension, which performs the
 * read-only Nexl requests with the user's Nexl login. Also orchestrates one full Nexl pull.
 */
(function (root) {
  "use strict";
  const P = root.NexlParsers;
  let seq = 0;
  const pending = new Map();
  let bridgeVersion = null;

  window.addEventListener("message", (e) => {
    if (e.source !== window) return;
    const d = e.data;
    if (!d || d.source !== "nexl-check-bridge") return;
    if (d.type === "hello") { bridgeVersion = d.version; return; }
    const p = pending.get(d.id);
    if (!p) return;
    pending.delete(d.id);
    clearTimeout(p.timer);
    p.resolve(d);
  });

  function send(type, path, timeoutMs = 45000, extra = {}) {
    return new Promise((resolve) => {
      const id = "r" + ++seq + "_" + Date.now();
      const timer = setTimeout(() => {
        pending.delete(id);
        resolve({ ok: false, error: type === "ping" ? "NOT_INSTALLED" : "TIMEOUT" });
      }, timeoutMs);
      pending.set(id, { resolve, timer });
      window.postMessage(Object.assign({ source: "nexl-check-addin", id, type, path }, extra), window.location.origin);
    });
  }

  /** Focus Nexl and run its search. Needs bridge 1.1+; older bridges answer UNKNOWN_REQUEST. */
  /** Bridge 1.3+: with an instruction number, opens it from Active/Completed Instructions like a controller would. */
  function openInNexl(filter, word, instruction, screen) {
    return send("open", "", 60000, { filter, word, instruction: instruction || null, screen: screen || null });
  }

  class NexlError extends Error {
    constructor(code, detail) { super(detail || code); this.code = code; }
  }

  async function get(path) {
    const r = await send("fetch", path);
    if (!r.ok) throw new NexlError(r.error || "FAILED", r.detail);
    return r.text;
  }

  async function ping() {
    const r = await send("ping", "", 2500);
    if (r.ok) bridgeVersion = r.version;
    return r;
  }

  /** Run tasks with a small concurrency limit so we don't hammer Nexl. */
  async function pool(items, limit, fn) {
    const out = new Array(items.length);
    let i = 0;
    const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (i < items.length) { const k = i++; out[k] = await fn(items[k], k); }
    });
    await Promise.all(workers);
    return out;
  }

  const statusPath = (type, region) =>
    `/php/ajax/statusscreen/get.status.screen.template.php?status_type=${type}&p_region=${region}&p_tab_selected=4`;

  /** Active + recently completed instructions for a region. */
  async function getInstructions(region) {
    const [a, c] = await Promise.all([get(statusPath("active", region)), get(statusPath("completed", region))]);
    const act = P.parseStatusScreen(a), com = P.parseStatusScreen(c);
    if (!act || !com) throw new NexlError("LAYOUT_CHANGED", "Nexl's instruction list looks different than expected.");
    const seen = new Set();
    return [...act.map((x) => ({ ...x, state: "Active" })), ...com.map((x) => ({ ...x, state: "Completed" }))]
      .filter((x) => !seen.has(x.id) && seen.add(x.id));
  }

  async function getContainers(ids, onProgress) {
    const out = {};
    let done = 0;
    await pool(ids, 4, async (id) => {
      const list = P.parseContainers(await get(`/php/ajax/instruction_containers/get.containers.fresh.php?p_id=${encodeURIComponent(id)}`), id);
      out[id] = list || [];
      onProgress && onProgress(++done, ids.length);
    });
    return out;
  }

  async function getTracking() {
    try { return P.parseDriverTracking(await get("/php/ajax/drivertracking/get.driver.table.php")) || []; }
    catch (e) { if (e.code === "NOT_LOGGED_IN") throw e; return []; } // tracking is optional
  }

  async function searchContainers(list) {
    const out = {};
    await pool(list, 4, async (c) => {
      const word = String(c).toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (word.length < 4) return;
      try { out[c] = P.parseSearch(await get(`/php/ajax/search/get.search.php?p_filter=container&p_word=${word}`)) || []; }
      catch (e) { if (e.code === "NOT_LOGGED_IN") throw e; }
    });
    return out;
  }

  /** Upload Viewer HTML for one container row (lists the driver's photos). */
  function getUploads(rowId) {
    if (!/^\d+$/.test(String(rowId))) return Promise.reject(new NexlError("NO_ROW", "No Nexl row id for this container"));
    return get(`/php/ajax/document_viewer/get.viewer.php?p_prog=${rowId}`);
  }
  /** One uploaded photo as a (resized) data URL. Needs bridge 1.2+. */
  async function getImage(path) {
    const r = await send("image", String(path).replace(/^https?:\/\/[^/]+/, ""), 45000);
    if (!r.ok) throw new NexlError(r.error || "FAILED", r.error === "UNKNOWN_REQUEST" ? "Update the Nexl Check Bridge extension to 1.2 to check photos" : r.detail);
    return r.dataUrl;
  }

  root.NexlClient = { ping, getUploads, getImage, openInNexl, getInstructions, getContainers, getTracking, searchContainers, NexlError, get bridgeVersion() { return bridgeVersion; } };
})(window);
