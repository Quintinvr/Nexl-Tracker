/*
 * Nexl Check — parsers for the HTML fragments returned by Nexl Controller.
 * Nexl renders its tables server-side, so we parse the <table> markup by header name
 * (not by position) to survive small layout changes.
 */
(function (root) {
  "use strict";

  function getParser() {
    if (root.__NEXL_PARSE_HTML) return root.__NEXL_PARSE_HTML; // test hook
    const dp = new DOMParser();
    return (html) => dp.parseFromString(html, "text/html");
  }

  const clean = (s) => String(s == null ? "" : s).replace(/\s+/g, " ").trim();

  function normInstr(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") v = v.toFixed(2);
    let s = clean(v).replace(/,/g, "");
    if (!/^\d+(\.\d+)?$/.test(s)) return "";
    let [a, b = ""] = s.split(".");
    b = (b + "00").slice(0, 2);
    return b === "00" ? a : a + "." + b;
  }
  const baseInstr = (id) => String(id).split(".")[0];

  /** Read a table into objects keyed by normalised header text. */
  function readTable(table) {
    const heads = [...table.querySelectorAll("thead th")].map((h) => clean(h.textContent).toUpperCase());
    const out = [];
    for (const tr of table.querySelectorAll("tbody tr")) {
      const cells = [...tr.children].filter((c) => c.tagName === "TD" || c.tagName === "TH");
      if (!cells.length) continue;
      const rec = { __tr: tr };
      heads.forEach((h, i) => {
        if (!h || !cells[i]) return;
        if (!(h in rec)) rec[h] = clean(cells[i].textContent);
      });
      out.push(rec);
    }
    return { heads, rows: out };
  }

  function firstTable(html, mustHave) {
    const doc = getParser()(html);
    for (const t of doc.querySelectorAll("table")) {
      const { heads, rows } = readTable(t);
      if (mustHave.every((m) => heads.includes(m))) return { heads, rows };
    }
    return null;
  }

  /** Active / Completed Instructions status screen. */
  function parseStatusScreen(html) {
    const t = firstTable(html, ["INSTRUCTION", "CUSTOMER"]);
    if (!t) return null;
    return t.rows
      .map((r) => {
        const id = normInstr(r["INSTRUCTION"]);
        if (!id) return null; // skips the red TOTAL row
        return {
          id,
          base: baseInstr(id),
          startTime: r["START TIME"] || "",
          customer: r["CUSTOMER"] || "",
          booking: r["BOOKING REFERENCE"] || "",
          vessel: r["VESSEL"] || "",
          type: r["TYPE"] || "",
          remaining: r["REMAINING"] || "",
          route: r["ROUTE"] || "",
          completion: r["COMPLETION"] || "",
          location: r["LOCATION"] || "",
        };
      })
      .filter(Boolean);
  }

  /** Planning board: planned instructions with their LOAD REF (row tooltip) and the refs in "Booking / Ref". */
  function parsePlanning(html) {
    const t = firstTable(html, ["INSTRUCTION", "CUSTOMER"]);
    if (!t) return null;
    return t.rows.map((r) => {
      const id = normInstr(r["INSTRUCTION"]);
      if (!id) return null;
      const title = (r.__tr && (r.__tr.getAttribute ? r.__tr.getAttribute("title") : r.__tr.title)) || "";
      const refs = [];
      for (const m of String(title).matchAll(/LOAD REF:\s*([^\n\r]+)/gi)) {
        const v = m[1].trim();
        refs.push(v);
        // "GE3832/33/34/35" -> GE3832, GE3833, GE3834, GE3835
        const parts = v.replace(/\s+/g, "").split("/");
        const n = parts.length > 1 ? parts[1].length : 0, tail = (x) => +x.slice(-n);
        if (n && /^[A-Za-z]*\d{3,}$/.test(parts[0]) && parts.slice(1).every((x, i) => /^\d{1,3}$/.test(x) && x.length === n && +x > tail(i ? parts[i] : parts[0]))) {
          for (const x of parts.slice(1)) refs.push(parts[0].slice(0, parts[0].length - x.length) + x);
          refs.push(parts[0]);
        }
      }
      const book = String(r["BOOKING REFERENCE"] || "").split("/").map((x) => x.trim()).filter(Boolean);
      refs.push(...book.slice(1));
      return { id, base: baseInstr(id), customer: r["CUSTOMER"] || "", booking: book[0] || "", vessel: r["VESSEL"] || "", startTime: r["START TIME"] || "",
        loadRefs: [...new Set(refs.filter((x) => x.replace(/[^A-Za-z0-9]/g, "").length >= 4 && !/^TBA$/i.test(x)))] };
    }).filter(Boolean);
  }

  /** Containers on one instruction (the container table in an instruction's detail view). */
  function parseContainers(html, instrId) {
    const t = firstTable(html, ["CONTAINER", "SEAL"]);
    if (!t) return null;
    return t.rows
      .map((r) => {
        const container = r["CONTAINER"] || "";
        if (!container) return null;
        const rid = /container_table_row_(\d+)/.exec((r.__tr && r.__tr.id) || "");
        return {
          instruction: normInstr(instrId),
          rowId: rid ? rid[1] : "",
          container,
          seal: r["SEAL"] || "",
          ref2: r["REFERENCE 2"] || "",
          ref3: r["REFERENCE 3"] || "",
          owner: r["OWNER"] || "",
          driver: r["DRIVER"] || "",
          start: r["START"] || "",
          end: r["END"] || "",
          podStatus: r["POD STATUS"] || "",
          moveStatus: r["MOVE STATUS"] || "",
          adminStatus: r["ADMIN STATUS"] || "",
          invoiceSent: r["INVOICE SENT"] || "",
        };
      })
      .filter(Boolean);
  }

  /** Driver Tracking screen: live position of each driver's current job. */
  function parseDriverTracking(html) {
    const t = firstTable(html, ["DRIVER", "CONTAINER", "LAST PING"]);
    if (!t) return null;
    const na = (v) => (v && v.toUpperCase() !== "N/A" ? v : "");
    return t.rows.map((r) => ({
      driver: r["DRIVER"] || "",
      owner: r["OWNER"] || "",
      customer: r["CUSTOMER"] || "",
      instruction: normInstr(r["INSTRUCTION NUMBER"]),
      reference: r["REFERENCE"] || "",
      container: r["CONTAINER"] || "",
      route: r["ROUTE"] || "",
      pickup: na(r["PICK-UP ENTRY"]),
      via: na(r["VIA ENTRY"]),
      via2: na(r["VIA2 ENTRY"]),
      dropoff: na(r["DROP-OFF ENTRY"]),
      jobDuration: na(r["JOB DURATION"]),
      lastPing: na(r["LAST PING"]),
    }));
  }

  /** Global container search (Nexl's search box, filter = container). */
  function parseSearch(html) {
    const t = firstTable(html, ["INSTRUCTION NUMBER", "CONTAINER"]);
    if (!t) return null;
    return t.rows
      .map((r) => ({
        instruction: normInstr(r["INSTRUCTION NUMBER"]),
        container: r["CONTAINER"] || "",
        vessel: r["VESSEL"] || "",
        customer: r["CUSTOMER"] || "",
        booking: r["BOOKING REFERENCE"] || "",
        status: r["STATUS"] || "",
        fileStatus: r["FILE STATUS"] || "",
        transporter: r["TRANSPORTER"] || "",
      }))
      .filter((r) => r.instruction && r.container);
  }

  root.NexlParsers = { parseStatusScreen, parsePlanning, parseContainers, parseDriverTracking, parseSearch, normInstr, baseInstr, clean };
})(typeof window !== "undefined" ? window : globalThis);
