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
  const MONTHS = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, SEPT: 8, OCT: 9, NOV: 10, DEC: 11 };
  /** Load date cell -> "YYYY-MM-DD" (Excel serial, 05/10/2026, 2026-10-05, 05-Oct-2026, 5 Oct). "" if none. */
  function dateKey(v, year = new Date().getFullYear()) {
    const pad = (n) => String(n).padStart(2, "0");
    const k = (y, m, d) => (y > 1999 && m >= 0 && m < 12 && d >= 1 && d <= 31 ? `${y}-${pad(m + 1)}-${pad(d)}` : "");
    if (typeof v === "number" && v > 30000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 86400000)); return k(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
    const s = String(v == null ? "" : v).trim().toUpperCase();
    let m;
    if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/.exec(s))) return k(+m[1], +m[2] - 1, +m[3]);
    if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(s))) return k(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]); // SA: day/month/year
    if ((m = /^(\d{1,2})[-\s]?([A-Z]{3,4})[A-Z]*[-\s,]*(\d{2,4})?$/.exec(s)) && MONTHS[m[2]] !== undefined) return k(m[3] ? (m[3].length === 2 ? 2000 + +m[3] : +m[3]) : year, MONTHS[m[2]], +m[1]);
    return "";
  }
  /** "2026-10-05" -> "Mon 05 Oct" */
  function dateLabel(key) {
    if (!key) return "No date";
    const [y, m, d] = key.split("-").map(Number);
    const dt = new Date(y, m - 1, d);
    return `${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][dt.getDay()]} ${String(d).padStart(2, "0")} ${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][m - 1]}`;
  }

  // Comments that are just a status word (already shown as the status) are not repeated as a comment line.
  const STATUS_WORD = /^(STACKED|STACK(ED)? IN|DEPARTED|LOADED|YARD|IN YARD|COLLECTED|DELIVERED|GATED ?(IN|OUT)|IN TRANSIT|ON ROUTE|EN ROUTE|PLUGGED( IN)?|GROUNDED|PACKED|PACKING|EMPTY|DONE|COMPLETED?|ARRIVED|OFFLOADED|AT PORT|SHIPPED|SAILED|ON BOARD|ONBOARD|BOOKED|ALLOCATED|DISPATCHED|RELEASED|CANCELL?ED|ROLLED|TBC|OK|N\/?A)\.?$/i;
  const isStatusWord = (t) => STATUS_WORD.test(String(t || "").trim());

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
      const comment = get(r, "comment");
      // The client update uses ONLY what the controller typed in the COMMENT column.
      // Nexl's live status is not put in the message (it is shown in the panel as a hint only: nexlHint).
      const status = comment || "Awaiting update", icon = comment ? "🔹" : "⏳", remark = "";
      const nexlHint = live ? live.text : "";
      loads.push({
        id: `${tabName}|${i + 1}`, row: i + 1, client, clientKey: compact(client),
        vessel: get(r, "vessel"), loadRef: loadRef || container, container, seal: get(r, "seal"),
        tare: get(r, "tare"), booking: get(r, "booking"), status, icon, live: !!(live && !live.completed), remark, nexlHint, hasComment: !!comment,
        date: cols.loadDate === undefined ? "" : dateKey(r[cols.loadDate]),
        instruction: get(r, "instruction"), driver: get(r, "driver"), transporter: get(r, "transporter") || (d && d.nexl && d.nexl.owner) || "",
        nexlDriver: (d && d.nexl && d.nexl.driver) || "", portBooking: portBooking(r),
        // Route for the driver message: Nexl's instruction route first, the sheet's depot/packstore/port as fallback.
        route: routeFor(d, { collect: get(r, "emptyDepot"), packing: get(r, "packstore"), dropoff: get(r, "port") }),
      });
    });
    return loads;
  }

  /**
   * Port booking (terminal slot) as controllers type it, e.g. "29734729 - 16:00". It sits in different columns
   * per sheet (EXPORTS: SEAL, IMPORTS: NAVIS CHECK, sometimes COMMENT), so look for the pattern on the whole row.
   */
  const PORT_BOOKING = /\b(\d{6,10})\s*[-–]\s*(\d{1,2})[:h.](\d{2})\b/;
  function portBooking(values) {
    for (const v of values || []) {
      const m = PORT_BOOKING.exec(String(v == null ? "" : v));
      if (m) return `${m[1]} - ${m[2].padStart(2, "0")}:${m[3]}`;
    }
    return "";
  }

  function routeFor(d, sheet) {
    const stops = parseRoute(((d && d.instr) || {}).route || "");
    if (stops.length >= 2) {
      return { source: "nexl", collect: stops[0].name, packing: stops.slice(1, -1).map((s) => s.name), dropoff: stops[stops.length - 1].name };
    }
    return { source: "sheet", collect: sheet.collect, packing: sheet.packing ? [sheet.packing] : [], dropoff: sheet.dropoff };
  }

  /**
   * Basic job message for the driver:
   *   Job on app
   *
   *   Collect - MSC COEGA
   *   Packing - KHOLD COEGA COLDSTORE
   *   Drop off - COEGA CT
   */
  function driverMessage(l, opts = {}) {
    const e = !!opts.emoji, r = l.route || {}, f = opts.fields || {};
    const line = (icon, label, v) => `${e ? icon + " " : ""}${label} - ${v || ""}`;
    const out = [e ? "📱 Job on app" : "Job on app", ""];
    // Optional on the Collect line: container and booking (imports / full exports).
    const extra = [f.container && l.container && `Container: ${l.container}`, f.booking && l.portBooking && `Port booking: ${l.portBooking}`].filter(Boolean);
    out.push(line("📦", "Collect", r.collect) + (extra.length ? ` (${extra.join(" · ")})` : ""));
    const packs = (r.packing || []).filter(Boolean); // no packing stop (e.g. imports: port -> depot) = no Packing line
    packs.forEach((p, i) => out.push(line("🏭", packs.length > 1 ? `Packing ${i + 1}` : "Packing", p)));
    out.push(line("⚓", "Drop off", r.dropoff));
    return out.join("\n");
  }

  /** Driver jobs from any plan tab's rows ({row, values}) as read by the sync (current Nexl instructions only). */
  function driverJobs(tab, sheetRows, cols, rowsMap, fallback) {
    const get = (r, f) => (!f || cols[f] === undefined ? "" : clean(r[cols[f]]));
    const out = [];
    for (const { row, values } of sheetRows || []) {
      const container = get(values, "container"), instruction = get(values, "instruction");
      if (!instruction || (!container && !get(values, "loadRef"))) continue;
      const d = rowsMap ? rowsMap[`${tab}|${row}`] : null;
      out.push({
        id: `${tab}|${row}`, row, tab, client: get(values, "customer"), loadRef: get(values, "loadRef") || container || instruction, container,
        booking: get(values, "booking"), portBooking: portBooking(values), instruction, driver: get(values, "driver"),
        transporter: get(values, "transporter") || (d && d.nexl && d.nexl.owner) || "", nexlDriver: (d && d.nexl && d.nexl.driver) || "", date: cols.loadDate === undefined ? "" : dateKey(values[cols.loadDate]),
        route: routeFor(d, { collect: get(values, fallback.collect), packing: get(values, fallback.packing), dropoff: get(values, fallback.dropoff) }),
      });
    }
    return out;
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
  /**
   * WhatsApp message. *text* shows as bold in WhatsApp.
   *   Good day,
   *
   *   UPDATE
   *
   *   *CMA CGM KRIBI*
   *
   *   *CR154444* — In transit to COEGA CT (left KHOLD COEGA COLDSTORE 12:48)
   *      Container: MNBU0333896
   *      Seal: ML-ZA6734477
   *
   *   Kind regards.
   * With emojis on, the same layout gets 📢 / 🔷 / 🚢 / 🔹 in front.
   */
  function formatMessage(loads, opts = {}) {
    const e = opts.emoji !== false, f = opts.fields || {};
    const sep = " — ";
    const bold = (t) => `*${String(t).replace(/\*/g, "")}*`;
    const out = [e ? "📢 Good day," : "Good day,", "", e ? "🔷 UPDATE" : "UPDATE"];
    const vessels = new Map();
    for (const l of loads) { const v = l.vessel || "VESSEL TBC"; if (!vessels.has(v)) vessels.set(v, []); vessels.get(v).push(l); }
    for (const [vessel, list] of vessels) {
      out.push("", `${e ? "🚢 " : ""}${bold(vessel)}`);
      for (const l of list) {
        out.push("", `${e ? "🔹 " : ""}${bold(l.loadRef)}${sep}${l.status}`);
        const extra = [
          l.remark && `Comment: ${l.remark}`, // controller's comment: always included
          f.container && l.container && `Container: ${l.container}`,
          f.seal && l.seal && `Seal: ${l.seal}`,
          f.tare && l.tare && `Tare: ${l.tare}`,
          f.booking && l.booking && `Booking: ${l.booking}`,
        ].filter(Boolean);
        for (const x of extra) out.push(`   ${x}`);
      }
    }
    out.push("", "Kind regards.");
    return out.join("\n");
  }

  root.NexlWhatsApp = { portBooking, driverMessage, driverJobs, routeFor, isStatusWord, dateKey, dateLabel, buildLoads, byClient, formatMessage, liveStatus };
})(typeof window !== "undefined" ? window : globalThis);
