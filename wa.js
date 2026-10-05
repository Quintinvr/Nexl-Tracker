/*
 * Nexl Check — WhatsApp client updates (copy only).
 * Source: every load on the PE CITRUS tab (loads move to LOADS DONE when finished).
 * Key: LOAD REF, then status. Status = live Nexl progress while a truck is on it,
 * otherwise the sheet's COMMENT (e.g. "STACKED"), then NAVIS CHECK, then "Awaiting update".
 */
(function (root) {
  "use strict";
  const clean = (s) => String(s == null ? "" : s).replace(/\s+/g, " ").trim();
  const compact = (s) => clean(s).toUpperCase().replace(/[^A-Z0-9]/g, "");
  const parseRoute = (r) => root.NexlMatcher.parseRoute(r);

  /** Live status from Nexl, in client-friendly words. Returns {text, icon} or null. */
  function liveStatus(d) {
    if (!d || !d.leg) return null;
    const leg = d.leg, stops = leg.stops || [];
    const pod = (parseRoute((d.instr || {}).route).find((s) => s.role === "POD") || {}).name;
    if (leg.stage === "allocated") return { icon: "🕒", text: `Truck allocated${stops[0] ? ", collecting at " + stops[0].name : ""}` };
    if (leg.stage === "moving") {
      const i = stops.map((s) => s.done).lastIndexOf(true), at = stops[i] || {}, nx = stops[i + 1] || {};
      return { icon: "🚚", text: `In transit to ${nx.name || "next stop"}${at.time ? ` (left ${at.name} ${at.time})` : ""}` };
    }
    if (leg.stage === "delivered") { const s = stops[stops.length - 1] || {}; return { icon: "✅", text: `Delivered at ${s.name || pod || "destination"}${s.time ? " " + s.time : ""}` }; }
    if (leg.stage === "completed") return { icon: "✅", text: `Delivered${pod ? " at " + pod : ""}`, completed: true };
    return null;
  }

  /**
   * @param grid   PE CITRUS values (row 1 = headers), as read from Excel
   * @param cols   column map from NexlMatcher.detectColumns
   * @param rows   result.rows from the matcher (keyed "PE CITRUS|<row>") for live status
   */
  function buildLoads(grid, cols, rows, tabName) {
    const get = (r, f) => (cols[f] === undefined ? "" : clean(r[cols[f]]));
    const loads = [];
    (grid || []).forEach((r, i) => {
      if (i === 0 || !r) return; // header
      const client = get(r, "customer");
      const loadRef = get(r, "loadRef");
      const container = get(r, "container");
      if (!client || (!loadRef && !container)) return; // date banners / empty rows
      const d = rows ? rows[`${tabName}|${i + 1}`] : null;
      const live = liveStatus(d);
      const sheet = get(r, "comment") || get(r, "navis");
      // Live progress wins while a truck is on it; once complete, the sheet's own wording (e.g. STACKED) wins.
      let status, icon;
      if (live && !live.completed) { status = live.text; icon = live.icon; }
      else if (sheet) { status = sheet; icon = live ? live.icon : "🔹"; }
      else if (live) { status = live.text; icon = live.icon; }
      else { status = "Awaiting update"; icon = "⏳"; }
      loads.push({
        id: `${tabName}|${i + 1}`, row: i + 1, client, clientKey: compact(client),
        vessel: get(r, "vessel"), loadRef: loadRef || container, container, seal: get(r, "seal"),
        tare: get(r, "tare"), booking: get(r, "booking"), status, icon, live: !!(live && !live.completed),
      });
    });
    return loads;
  }

  function byClient(loads) {
    const m = new Map();
    for (const l of loads) {
      if (!m.has(l.clientKey)) m.set(l.clientKey, { client: l.client, key: l.clientKey, loads: [] });
      m.get(l.clientKey).loads.push(l);
    }
    return [...m.values()].sort((a, b) => a.client.localeCompare(b.client));
  }

  /**
   * opts: { emoji: bool, fields: { container, seal, tare, booking } }
   * One message, split per vessel; each line = LOAD REF then status (+ optional details on the next line).
   */
  function formatMessage(loads, opts = {}) {
    const e = opts.emoji !== false, f = opts.fields || {};
    const sep = e ? " | " : " — ";
    const out = [e ? "📢 Good day," : "Good day,", e ? "🔷 UPDATE" : "UPDATE"];
    const vessels = new Map();
    for (const l of loads) { const v = l.vessel || "VESSEL TBC"; if (!vessels.has(v)) vessels.set(v, []); vessels.get(v).push(l); }
    for (const [vessel, list] of vessels) {
      out.push("", e ? `🚢 ${vessel}` : vessel);
      for (const l of list) {
        out.push(`${e ? "🔹 " : ""}${l.loadRef}${sep}${l.status}`);
        const extra = [
          f.container && l.container && `Container: ${l.container}`,
          f.seal && l.seal && `Seal: ${l.seal}`,
          f.tare && l.tare && `Tare: ${l.tare}`,
          f.booking && l.booking && `Booking: ${l.booking}`,
        ].filter(Boolean);
        if (extra.length) out.push(`${e ? "      " : "   "}${extra.join(e ? " · " : ", ")}`);
      }
    }
    out.push("Kind regards.");
    return out.join("\n");
  }

  root.NexlWhatsApp = { buildLoads, byClient, formatMessage, liveStatus };
})(typeof window !== "undefined" ? window : globalThis);
