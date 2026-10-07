/*
 * Nexl Check — compares transport-plan rows with Nexl data.
 * Pure functions: no Excel or network calls, so it can be unit-tested.
 */
(function (root) {
  "use strict";
  const { normInstr, baseInstr, clean } = root.NexlParsers;

  const STOP = new Set(["PTY", "LTD", "THE", "AND", "TRANSPORT", "LOGISTICS", "SOLUTIONS", "CARRIERS", "SERVICES",
    "GROUP", "HOLDINGS", "AFRICA", "SOUTH", "TRADING", "FREIGHT"]);

  const compact = (s) => clean(s).toUpperCase().replace(/[^A-Z0-9]/g, "");
  const tokens = (s, min = 3) =>
    clean(s).toUpperCase().split(/[^A-Z0-9]+/).filter((t) => t.length >= min && !STOP.has(t));

  function editDistance(a, b) {
    if (Math.abs(a.length - b.length) > 2) return 99;
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++)
      for (let j = 1; j <= b.length; j++)
        d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
  }

  function colLetter(i) {
    let s = "";
    i += 1;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  /** Lenient equality for names (customer / transporter). Empty on either side = nothing to compare. */
  // Known short names used on the sheet (sheet value -> name in Nexl). Extend in config.js > nameAliases.
  const ALIASES = { FL4U: ["FREIGHT LOGISTICS 4U", "FREIGHT LOGISTICS FOR YOU"] };
  const aliasMap = () => {
    const extra = (root.NEXL_CONFIG && root.NEXL_CONFIG.nameAliases) || {};
    const m = {};
    for (const [k, v] of Object.entries({ ...ALIASES, ...extra })) m[compact(k)] = [].concat(v).map(compact);
    return m;
  };
  // "FOR" -> 4, "YOU" -> U, "TO" -> 2 ... so "Freight Logistics For You" and "Freight Logistics 4 U" both give FL4U.
  const NUMWORD = { FOR: "4", FOUR: "4", TO: "2", TOO: "2", TWO: "2", YOU: "U", ONE: "1", AND: "N" };
  function initialsVariants(s) {
    const words = clean(s).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
    const out = new Set();
    out.add(words.map((w) => w[0]).join(""));                                  // IOEC
    out.add(words.map((w) => (/\d/.test(w) ? w : w[0])).join(""));            // FREIGHT LOGISTICS 4U -> FL4U
    out.add(words.map((w) => (NUMWORD[w] || (/\d/.test(w) ? w : w[0]))).join("")); // FREIGHT LOGISTICS FOR YOU -> FL4U
    return [...out];
  }
  function nameMatch(a, b) {
    const ca = compact(a), cb = compact(b);
    if (!ca || !cb) return true;
    if (ca === cb || ca.includes(cb) || cb.includes(ca)) return true;
    const al = aliasMap();
    if ((al[ca] || []).some((x) => cb.startsWith(x) || x.startsWith(cb)) || (al[cb] || []).some((x) => ca.startsWith(x) || x.startsWith(ca))) return true;
    // Abbreviations: "IOEC" = "Indian Ocean Export Company", "FL4U" = "Freight Logistics 4 U"
    if (ca.length >= 2 && initialsVariants(b).some((i) => i.startsWith(ca))) return true;
    if (cb.length >= 2 && initialsVariants(a).some((i) => i.startsWith(cb))) return true;
    const tb = new Set(tokens(b));
    return tokens(a).some((t) => tb.has(t) || [...tb].some((x) => x.startsWith(t) || t.startsWith(x)));
  }

  /** Strict-ish equality for vessels: one must contain the other once spaces are removed. */
  function vesselMatch(a, candidates) {
    const ca = compact(a);
    if (!ca) return true;
    const cs = candidates.map(compact).filter(Boolean);
    if (!cs.length) return true;
    return cs.some((c) => c.includes(ca) || ca.includes(c));
  }

  /** References: any part of the sheet value (split on / , space) found inside any Nexl reference. */
  function refMatch(a, candidates) {
    const ca = compact(a);
    if (!ca) return true;
    const cs = candidates.map(compact).filter(Boolean);
    if (!cs.length) return true;
    if (cs.some((c) => c.includes(ca) || ca.includes(c))) return true;
    const parts = clean(a).toUpperCase().split(/[\/,;\s]+/).map(compact).filter((p) => p.length >= 4);
    return parts.some((p) => cs.some((c) => c.includes(p)));
  }

  /**
   * @param sheetTabs [{name, compare:[field], cols:{field: index}, rows:[{row, values:[]}]}]
   * @param nexl {instructions:[instr], containers:{id:[c]}, tracking:[t]}
   */
  function compare(sheetTabs, nexl, opts = {}) {
    const now = opts.now ? new Date(opts.now) : new Date();
    const notStartedMin = opts.notStartedMinutes ?? 30;
    const stuckMin = opts.stuckMinutes ?? 120;
    const rowRecs = []; // every plan row in scope (incl. open slots) -> gets a NEXL STEP / NEXL ALERT
    const search = {}; // compact container -> search results (only for containers we looked up)
    for (const [k, v] of Object.entries(nexl.search || {})) search[compact(k)] = v;
    const instrById = new Map();
    for (const i of nexl.instructions || []) if (!instrById.has(i.id)) instrById.set(i.id, i);
    const bases = new Set([...instrById.values()].map((i) => i.base));
    for (const id of Object.keys(nexl.containers || {})) bases.add(baseInstr(id));

    // Nexl containers grouped by base, plus a global index by container number.
    const nexlByBase = new Map();
    const nexlByContainer = new Map();
    for (const [id, list] of Object.entries(nexl.containers || {})) {
      for (const c of list || []) {
        const rec = { ...c, instruction: c.instruction || id, key: compact(c.container), used: false };
        const b = baseInstr(rec.instruction);
        if (!nexlByBase.has(b)) nexlByBase.set(b, []);
        nexlByBase.get(b).push(rec);
        if (!nexlByContainer.has(rec.key)) nexlByContainer.set(rec.key, []);
        nexlByContainer.get(rec.key).push(rec);
      }
    }
    const trackByContainer = new Map();
    for (const t of nexl.tracking || []) {
      const k = compact(t.container);
      if (k && !/^PENDING/.test(k)) trackByContainer.set(k, t);
    }

    const issues = [];
    const groups = new Map(); // base -> group
    const seen = new Set();
    const stats = { rowsChecked: 0, matched: 0, mismatched: 0, notInNexl: 0, missingFromSheet: 0, openSlots: 0 };

    const groupFor = (b) => {
      if (!groups.has(b)) {
        const ids = [...instrById.values()].filter((i) => i.base === b);
        groups.set(b, { base: b, nexl: ids, tabs: new Set(), firstRef: null, containers: [], openSlots: 0 });
      }
      return groups.get(b);
    };

    for (const tab of sheetTabs) {
      const col = tab.cols;
      const get = (vals, f) => (col[f] === undefined ? "" : clean(vals[col[f]]));
      for (const r of tab.rows) {
        const id = normInstr(r.values[col.instruction]);
        if (!id) continue;
        const b = baseInstr(id);
        if (!bases.has(b)) continue;
        const g = groupFor(b);
        g.tabs.add(tab.name);
        const cellRef = (f) => ({ tab: tab.name, row: r.row, col: col[f] === undefined ? null : colLetter(col[f]) });
        if (!g.firstRef) g.firstRef = cellRef("instruction");

        const cont = get(r.values, "container");
        const key = compact(cont);
        const rec = { tab: tab.name, row: r.row, id, base: b, key, container: cont, tokens: new Set(r.values.flatMap((v) => tokens(v, 3))),
          ref: cellRef, entry: null, tracking: key ? trackByContainer.get(key) || null : null, group: g, vals: r.values, cols: col, compare: tab.compare, cutoffKind: tab.cutoff || null, slipField: tab.slip || null,
          // Planned collection: container pre-filled, no driver on the sheet yet -> no alerts until a driver is added.
          planned: !!(tab.planned && key && col.driver !== undefined && clean(r.values[col.driver]) === "") };
        rowRecs.push(rec);
        if (!key) { g.openSlots++; stats.openSlots++; continue; }
        stats.rowsChecked++;

        const entry = { container: cont, id, sheet: cellRef("container"), nexl: null, tracking: rec.tracking, issues: [], rec };
        rec.entry = entry;
        g.containers.push(entry);
        const add = (severity, field, sheetVal, nexlVal, message) => {
          const iss = { severity, field, sheet: sheetVal, nexl: nexlVal, message, instruction: id, container: cont, ref: cellRef(field === "container" || field === "instruction" ? field : field) };
          if (!iss.ref.col) iss.ref = cellRef("container");
          entry.issues.push(iss);
          issues.push(iss);
        };

        const dupKey = tab.name + "|" + id + "|" + key;
        if (seen.has(dupKey)) add("warn", "container", cont, "", "Container appears more than once for this instruction");
        seen.add(dupKey);

        const pool = nexlByBase.get(b) || [];
        const hits = pool.filter((n) => n.key === key);
        const hit = hits.find((n) => n.instruction === id && !n.used) || hits.find((n) => !n.used) || hits[0];
        if (!hit) {
          const elsewhere = (nexlByContainer.get(key) || [])[0];
          const searched = search[key];
          if (elsewhere) {
            add("error", "instruction", id, elsewhere.instruction, `Nexl has this container on instruction ${elsewhere.instruction}`);
          } else if (searched && searched.length) {
            const latest = [...searched].sort((x, y) => parseFloat(y.instruction) - parseFloat(x.instruction))[0];
            // Nexl's search "Status" says "Invoiced" once a file reaches finance; that isn't reliable (the Invoice Sent
            // Yes/No column is), so it's left out of the description.
            const st = /invoic/i.test(latest.status || "") ? "" : latest.status;
            const info = [latest.customer, st].filter(Boolean).join(", ");
            add("error", "container", cont, "", `Not on this instruction in Nexl. Latest Nexl record: ${latest.instruction}${info ? ` (${info})` : ""}`);
          } else if (searched) {
            add("error", "container", cont, "", "Container not found anywhere in Nexl");
          } else {
            add("error", "container", cont, "", "Container not found on this instruction in Nexl");
          }
          entry.notInNexl = true;
          stats.notInNexl++;
          continue;
        }
        hit.used = true;
        entry.nexl = hit;
        const instr = instrById.get(hit.instruction) || instrById.get(id) || instrById.get(b) || null;
        entry.instr = instr;
        const before = entry.issues.length;
        const rowText = rec.tokens;

        for (const f of tab.compare) {
          const sv = get(r.values, f);
          if (f === "seal") {
            const nv = /^\d{1,2}$/.test(hit.seal) ? "" : hit.seal; // Nexl uses "1" as a placeholder
            const a = compact(sv), b2 = compact(nv);
            const sameSeal = a === b2 || (Math.min(a.length, b2.length) >= 6 && (a.endsWith(b2) || b2.endsWith(a))); // "ML-ZA6724086" = "ZA6724086"
            if (sv && nv && !sameSeal) add("error", "seal", sv, nv, "Seal number differs");
            else if (!sv && nv) add("info", "seal", "", nv, "Seal not on sheet yet");
            else if (sv && !nv) add("info", "seal", sv, "", "Seal not captured in Nexl yet");
          } else if (f === "booking") {
            const cands = [instr && instr.booking, hit.ref2, hit.ref3].filter(Boolean);
            if (!refMatch(sv, cands)) add("warn", "booking", sv, cands.join(" | "), "Booking reference differs");
          } else if (f === "loadRef") {
            const cands = [hit.ref2, hit.ref3, instr && instr.booking].filter(Boolean);
            if (!refMatch(sv, cands)) add("warn", "loadRef", sv, cands.join(" | "), "Load reference differs");
          } else if (f === "vessel") {
            const cands = [instr && instr.vessel, hit.ref3].filter(Boolean);
            if (!vesselMatch(sv, cands)) add("warn", "vessel", sv, cands.join(" | "), "Vessel differs");
          } else if (f === "customer") {
            const nv = instr ? instr.customer : "";
            if (!nameMatch(sv, nv)) add("warn", "customer", sv, nv, "Customer differs");
          } else if (f === "transporter") {
            if (!nameMatch(sv, hit.owner)) add("warn", "transporter", sv, hit.owner, "Transporter differs");
          } else if (f === "driver") {
            const nt = tokens(hit.driver, 3);
            if (!nt.length) continue;
            const st = tokens(sv, 3);
            const inRow = nt.slice(0, 2).some((t) => rowText.has(t));
            if (st.length) {
              const ok = st.some((t) => nt.includes(t) || nt.some((n) => n.startsWith(t) || t.startsWith(n))) || inRow;
              if (!ok) add("warn", "driver", sv, hit.driver, "Driver differs");
            } else if (!inRow) {
              add("info", "driver", "", hit.driver, "Driver not on sheet");
            }
          }
        }
        const real = entry.issues.slice(before).some((i) => i.severity !== "info");
        if (real) stats.mismatched++; else stats.matched++;
      }
    }

    // Containers Nexl has that no checked tab lists.
    for (const [b, list] of nexlByBase) {
      if (!groups.has(b)) continue; // instruction not on the checked tabs at all
      const g = groups.get(b);
      for (const n of list) {
        if (n.used) continue;
        if (/^PENDING/i.test(n.container)) continue;
        // Likely a typo? Pair it with a sheet container on the same instruction that Nexl doesn't know.
        const twin = g.containers.find((e) => e.notInNexl && !e.twin && editDistance(compact(e.container), n.key) <= 2);
        if (twin) {
          twin.twin = n;
          const iss = twin.issues.find((i) => i.field === "container" || i.field === "instruction");
          if (iss) { iss.nexl = n.container; iss.message = `Possible typo: Nexl has ${n.container} on ${n.instruction}`; }
          n.used = true;
          continue;
        }
        stats.missingFromSheet++;
        const iss = { severity: "error", field: "container", sheet: "", nexl: n.container, message: "In Nexl but not on the sheet",
          instruction: n.instruction, container: n.container, ref: g.firstRef };
        issues.push(iss);
        g.containers.push({ container: n.container, id: n.instruction, sheet: null, nexl: n, instr: instrById.get(n.instruction) || null,
          tracking: trackByContainer.get(n.key) || null, issues: [iss] });
      }
    }

    // ---------- Live progress: link Driver Tracking jobs to plan rows ----------
    const usedTrack = new Set(rowRecs.filter((r) => r.tracking).map((r) => r.tracking));
    for (const t of nexl.tracking || []) {
      if (usedTrack.has(t)) continue;
      const b = baseInstr(t.instruction || "");
      if (!groups.has(b)) continue;
      const tk = compact(t.container);
      if (tk && !/^PENDING/.test(tk) && rowRecs.some((r) => r.key === tk)) continue;
      const dt = tokens(t.driver, 3).slice(0, 2);
      if (!dt.length) continue;
      const cands = rowRecs.filter((r) => r.base === b && !r.tracking && dt.some((d) => r.tokens.has(d)));
      // Prefer planned slots without a container, then rows whose Nexl container has no driver match yet.
      const pick = cands.find((r) => !r.key) || cands[0];
      if (pick) { pick.tracking = t; usedTrack.add(t); if (pick.entry) pick.entry.tracking = t; }
      else {
        const g = groups.get(b);
        g.containers.push({ container: "Allocated: " + t.driver.split(" ").slice(0, 2).join(" "), id: t.instruction, sheet: null, nexl: null,
          tracking: t, issues: [], pending: true, leg: legStatus(t, null, now, notStartedMin, stuckMin), step: legStatus(t, null, now, notStartedMin, stuckMin).step });
      }
    }

    const acks = opts.acks || {};
    const nowMs = now.getTime();
    const rowStatus = [], fills = [], rows = {}, slipCandidates = [], muted = new Set();
    const slips = opts.slips || {};
    const sheetKeys = new Set(rowRecs.map((r) => r.key).filter(Boolean));
    const X = extrasContext(rowRecs, nexl, instrById, nexlByContainer, now, opts);

    // Container number the driver entered in the app, for a planned slot that has no container yet.
    // Each Nexl container is offered to one slot only; every offer must be photo-checked before Apply.
    const claimed = new Set();
    const realNexl = (n) => n && n.key && !/^PENDING/.test(n.key) && !sheetKeys.has(n.key) && !claimed.has(n);
    const slotsByBase = new Map();
    for (const r of rowRecs) if (!r.key) slotsByBase.set(r.base, (slotsByBase.get(r.base) || 0) + 1);
    function slotContainer(r) {
      const pool = (nexlByBase.get(r.base) || []).filter((n) => !n.used && realNexl(n));
      if (r.tracking) {
        const tc = compact(r.tracking.container);
        if (tc && !/^PENDING/.test(tc) && !sheetKeys.has(tc)) {
          const n = pool.find((x) => x.key === tc) || (nexlByContainer.get(tc) || []).find(realNexl) || null;
          if (!claimed.has(n || tc)) return { value: n ? n.container : r.tracking.container, n, key: tc, source: "tracking" };
        }
      }
      const byDriver = pool.filter((n) => tokens(n.driver, 3).slice(0, 2).some((t) => r.tokens.has(t)));
      if (byDriver.length === 1) return { value: byDriver[0].container, n: byDriver[0], source: "driver" };
      if (pool.length === 1 && slotsByBase.get(r.base) === 1 && !tokens(pool[0].driver, 3).length) return { value: pool[0].container, n: pool[0], source: "only" };
      return null;
    }
    for (const r of rowRecs) {
      const n = r.entry ? r.entry.nexl : null;
      const ls = legStatus(r.tracking, n, now, notStartedMin, stuckMin);
      ls.alerts = ls.alerts.concat(extraAlerts(r, n, ls, X));
      const anchor = r.ref("container").col ? r.ref("container") : r.ref("instruction");
      const progressIssues = [];
      for (const a of ls.alerts) {
        const iss = { severity: a.severity || "warn", field: "progress", code: a.code, short: a.short, sheet: "", nexl: ls.step, message: a.text, instruction: r.id,
          container: r.container || (r.tracking ? "(" + r.tracking.driver.split(" ")[0] + ")" : ""), ref: anchor };
        issues.push(iss);
        progressIssues.push(iss);
        if (r.entry) r.entry.issues.push(iss);
      }
      if (r.entry) { r.entry.step = ls.step; r.entry.leg = ls; }
      else if (r.tracking) {
        r.group.containers.push({ container: "Slot: " + r.tracking.driver.split(" ").slice(0, 2).join(" "), id: r.id, sheet: anchor, nexl: null,
          tracking: r.tracking, issues: progressIssues, pending: true, step: ls.step, leg: ls });
      }

      // Blank cells the add-in could fill from Nexl (never overwrites a value).
      const rowFills = [];
      const blank = (f) => r.cols[f] !== undefined && clean(r.vals[r.cols[f]]) === "";
      const addFill = (field, value) => {
        if (!value || !blank(field)) return;
        const f = { tab: r.tab, row: r.row, col: colLetter(r.cols[field]), field, value, instruction: r.id, container: r.container };
        rowFills.push(f); fills.push(f);
      };
      if (n) {
        if (r.compare.includes("seal") && !/^\d{1,2}$/.test(n.seal) && !/^(NO ?SEAL|0+)$/i.test(n.seal)) {
          const b4 = rowFills.length;
          addFill("seal", n.seal);
          if (rowFills.length > b4) Object.assign(rowFills[rowFills.length - 1], { needsPhoto: true, photoKind: "seal", nexlRowId: n.rowId || "" });
        }
        if (r.compare.includes("driver")) addFill("driver", (n.driver || "").split(" ")[0]);
        if (r.compare.includes("transporter")) addFill("transporter", (n.owner || "").split(" ")[0]);
      } else if (!r.key && r.tracking) {
        addFill("driver", r.tracking.driver.split(" ")[0]);
      }
      // Port slip uploaded in Nexl -> container is at / out of the port.
      const rk = r.tab + "|" + r.row;
      if (r.slipField && r.cols[r.slipField] !== undefined && n && n.rowId) {
        const isImport = /IMPORT/i.test(r.tab) || /import/i.test(((n && instrById.get(n.instruction)) || {}).type || "");
        const slip = slips[rk];
        if (slip && blank(r.slipField)) {
          const b4 = rowFills.length;
          addFill(r.slipField, isImport ? "COLLECTED" : "STACKED");
          if (rowFills.length > b4) Object.assign(rowFills[rowFills.length - 1], { slipPath: slip.path, slipKind: isImport ? "collected" : "stacked", source: "portslip" });
        } else if (!slip && blank(r.slipField) && /moving|delivered|completed|other/.test(ls.stage || "")) {
          slipCandidates.push({ key: rk, rowId: n.rowId });
        }
      }
      if (!r.key && blank("container")) {
        const c = slotContainer(r);
        if (c) {
          claimed.add(c.n || c.key);
          const before = rowFills.length;
          addFill("container", c.value);
          if (rowFills.length > before) Object.assign(rowFills[rowFills.length - 1], { nexlRowId: c.n ? c.n.rowId || "" : "", needsPhoto: true, source: c.source });
        }
      }

      if (r.planned) {
        // Not flagged: everything about this row stays quiet until the controller puts a driver's name on it.
        for (const i of (r.entry ? r.entry.issues : []).concat(progressIssues)) muted.add(i);
        if (r.entry) r.entry.issues = [];
        const live = ls.stage === "moving" || ls.stage === "allocated";
        rowStatus.push({ tab: r.tab, row: r.row, step: live ? ls.short : "📋 Planned · no driver yet", alert: "", level: "" });
        rows[rk] = { tab: r.tab, row: r.row, id: r.id, container: r.container, instr: (n && instrById.get(n.instruction)) || instrById.get(r.id) || (r.group.nexl[0] || null),
          nexl: n, tracking: r.tracking, leg: ls, issues: [], fills: rowFills, ref: anchor, slip: slips[rk] || null, cutoff: r.cutoff || null, planned: true };
        stats.planned = (stats.planned || 0) + 1;
        continue;
      }
      const rowIssues = (r.entry ? r.entry.issues : []).filter((i) => i.severity !== "info" && i.field !== "progress").concat(progressIssues);
      let open = 0;
      const parts = rowIssues.map((i) => {
        i.key = issueKey(i);
        const ak = ackFor(i, acks, nowMs);
        if (ak) {
          i.ack = ak; i.bypass = ak.kind === "bypass";
          return i.bypass ? `✔ ${ak.by} bypassed: ${i.short || shortIssue(i)}` : `👀 ${ak.by}: ${i.short || shortIssue(i)}`;
        }
        open++;
        return i.short || shortIssue(i);
      });
      let level = "";
      if (rowIssues.some((i) => i.severity === "error" && !i.ack)) level = "error";
      else if (open) level = "warn";
      else if (rowIssues.some((i) => i.ack && !i.bypass)) level = "ack";
      else if (parts.length) level = "ok"; // everything on the row was bypassed
      else if (n || r.tracking) level = "ok";
      rowStatus.push({ tab: r.tab, row: r.row, step: ls.short, alert: parts.length ? parts.join(" | ") : level === "ok" ? "✓" : "", level });
      rows[r.tab + "|" + r.row] = { tab: r.tab, row: r.row, id: r.id, container: r.container, instr: (n && instrById.get(n.instruction)) || instrById.get(r.id) || (r.group.nexl[0] || null),
        nexl: n, tracking: r.tracking, leg: ls, issues: (r.entry ? r.entry.issues : progressIssues), fills: rowFills, ref: anchor,
        slip: slips[rk] || null, cutoff: r.cutoff || null };
    }
    if (muted.size) {
      const keep = issues.filter((i) => !muted.has(i));
      issues.length = 0; issues.push(...keep);
      for (const g of groups.values()) for (const c of g.containers) if (c.issues) c.issues = c.issues.filter((i) => !muted.has(i));
    }
    // Snooze state for issues that aren't tied to a row status (e.g. "In Nexl but not on the sheet").
    for (const i of issues) if (!i.key) { i.key = issueKey(i); const ak = ackFor(i, acks, nowMs); if (ak) { i.ack = ak; i.bypass = ak.kind === "bypass"; } }

    const notOnPlan = [...instrById.values()].filter((i) => !groups.has(i.base));
    const order = { error: 0, warn: 1, info: 2 };
    issues.sort((a, b) => order[a.severity] - order[b.severity] || String(a.instruction).localeCompare(String(b.instruction)));
    const groupList = [...groups.values()].map((g) => ({ ...g, tabs: [...g.tabs] })).sort((a, b) => b.base.localeCompare(a.base));
    return { issues, groups: groupList, notOnPlan, stats, rowStatus, fills, rows, cutoffs: cutoffSummary(rowRecs, X, rows), slipCandidates, legMin: X.legMin, legSamples: legSamples(nexl.tracking || []) };
  }

  const FIELD_SHORT = { seal: "Seal", booking: "Booking ref", loadRef: "Load ref", vessel: "Vessel", customer: "Customer", transporter: "Transporter", driver: "Driver" };
  function shortIssue(i) {
    const icon = i.severity === "error" ? "⛔" : "⚠";
    if (FIELD_SHORT[i.field]) return `${icon} ${FIELD_SHORT[i.field]} ≠ ${String(i.nexl).split(" | ")[0]}`;
    const m = i.message || "";
    if (/^Possible typo/.test(m)) return `⛔ Typo? Nexl: ${i.nexl}`;
    if (/anywhere/.test(m)) return "⛔ Not in Nexl";
    let x;
    if ((x = /Latest Nexl record: ([\d.]+)/.exec(m))) return `⛔ Not on this instr (Nexl: ${x[1]})`;
    if ((x = /on instruction ([\d.]+)/.exec(m))) return `⛔ Nexl has it on ${x[1]}`;
    if (/not found/.test(m)) return "⛔ Not on this instr in Nexl";
    if (/more than once/.test(m)) return "⚠ Duplicate row";
    if (/not on the sheet/.test(m)) return "⛔ Missing from sheet";
    return `${icon} ${m}`;
  }
  /** Stable id for an issue, used to snooze/acknowledge it across syncs. */
  /** What the issue "looks like" right now. A bypass only holds while this stays the same. */
  function issueFp(i) {
    if (i.field === "progress") return "progress>" + (i.code || "");
    return compact(i.sheet) + ">" + compact(i.nexl);
  }
  /** The ack (snooze or bypass) that applies to an issue, if any. */
  /** Bypass for one kind of alert on a whole instruction (e.g. "vessel differs" on every row of 86888). */
  function instrKey(i) {
    return ["INSTR", baseInstr(i.instruction || ""), i.field, i.code || ""].join("|");
  }
  /** An instruction-wide bypass holds while the SHEET value stays the same (it says "the sheet is right"). */
  function instrFp(i) {
    return i.field === "progress" ? "progress>" + (i.code || "") : compact(i.sheet) + ">*";
  }
  function ackFor(i, acks, nowMs) {
    const ak = acks[i.key];
    if (ak && ak.until > nowMs && (ak.kind !== "bypass" || ak.fp === issueFp(i))) return ak; // values changed -> flag again
    const ai = i.instruction ? acks[instrKey(i)] : null;
    if (ai && ai.until > nowMs && ai.kind === "bypass" && ai.fp === instrFp(i)) return ai;
    return null;
  }
  function issueKey(i) {
    return [i.ref ? i.ref.tab : "", i.ref ? i.ref.row : "", i.field, i.code || "", compact(i.container)].join("|");
  }

  /** "202:20:44" -> minutes */
  function durMin(s) {
    const m = /^(\d+):(\d{2})(?::(\d{2}))?$/.exec(String(s || "").trim());
    return m ? +m[1] * 60 + +m[2] : null;
  }
  function fmtMin(min) {
    if (min == null) return "";
    min = Math.round(min);
    if (min < 60) return min + "m";
    const h = Math.floor(min / 60), m = min % 60;
    return h >= 48 ? `${Math.floor(h / 24)}d ${h % 24}h` : `${h}h${String(m).padStart(2, "0")}`;
  }
  /** "2026-10-05 07:52" (local time) -> Date */
  function parseEntry(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/.exec(String(s || ""));
    return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) : null;
  }
  const hhmm = (d) => `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;

  /** Route "A(POL) >> B(VIA) >> C(POD)" -> [{name, role}] */
  function parseRoute(route) {
    return String(route || "").split(/\s*>>\s*/).map((p) => {
      const m = /^(.*?)\s*\((POL|VIA|POD)\)\s*$/i.exec(p.trim());
      return m ? { name: m[1].trim(), role: m[2].toUpperCase() } : { name: p.trim(), role: "" };
    }).filter((s) => s.name);
  }

  /**
   * Where is the truck on its instruction?
   * Uses Driver Tracking entry times (pick-up / via / via2 / drop-off) against the route stops.
   */
  const shortName = (x) => { const v = String(x || "?"); return v.length > 20 ? v.slice(0, 19) + "…" : v; };
  const dayMon = (d) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(d || ""); return m ? `${m[3]} ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][+m[2] - 1]}` : ""; };
  const podMark = (p) => (!p ? "" : /verified/i.test(p) ? " · POD ✓" : " · POD ⏳");

  function legStatus(t, n, now, notStartedMin = 30, stuckMin = 120) {
    const alerts = [];
    if (t) {
      const stops = parseRoute(t.route);
      const times = [];
      let viaN = 0;
      for (const s of stops) {
        if (s.role === "POL") times.push(parseEntry(t.pickup));
        else if (s.role === "POD") times.push(parseEntry(t.dropoff));
        else { times.push(parseEntry(viaN === 0 ? t.via : viaN === 1 ? t.via2 : "")); viaN++; }
      }
      let last = -1;
      times.forEach((d, i) => { if (d) last = i; });
      const stopList = stops.map((s, i) => ({ ...s, time: times[i] ? hhmm(times[i]) : null, done: i <= last && !!times[i] }));
      const legs = Math.max(stops.length - 1, 1);
      const who = t.driver.split(" ").slice(0, 2).join(" ");
      const first = t.driver.split(" ")[0];
      // No stop times from the app, but the driver HAS captured the container (Driver Tracking shows the real
      // number instead of PENDING1) or Nexl already has his seal -> he has loaded: not "not started".
      const capturedBox = /^[A-Z]{4}\d{7}$/.test(compact(t.container));
      const capturedSeal = !!(n && n.seal && !/^\d{1,2}$/.test(n.seal) && !/^(NO ?SEAL|0+)$/i.test(n.seal));
      if (last < 0 && (capturedBox || capturedSeal) && stops.length) {
        stopList[0] = { ...stopList[0], done: true, time: null };
        const next = stops[1];
        const what = [capturedBox && "container", capturedSeal && "seal"].filter(Boolean).join(" & ");
        return { step: `Loaded at ${stops[0].name} (${what} captured in the app, no stop times sent) → ${next ? next.name : "?"} · ${who}`,
          short: `🚚 Loaded → ${shortName(next ? next.name : "?")} · ${first}`, alerts, stage: "moving", stops: stopList, who,
          remaining: Math.max(stops.length - 1, 1), lastAt: null, nextName: next ? next.name : "", noTimes: true };
      }
      if (last < 0) {
        const jm = durMin(t.jobDuration);
        const step = `Allocated to ${who}${jm != null ? " " + fmtMin(jm) + " ago" : ""} · not started · first stop: ${stops[0] ? stops[0].name : "?"}`;
        if (jm != null && jm >= notStartedMin) alerts.push({ code: "notstarted", text: `Not started: allocated to ${who} ${fmtMin(jm)} ago, no pick-up yet`, short: `⏰ Not started ${fmtMin(jm)}` });
        return { step, short: `⏳ Allocated · ${first}${jm != null ? " · " + fmtMin(jm) : ""}`, alerts, stage: "allocated", stops: stopList, who, remaining: Math.max(stops.length - 1, 1), lastAt: null };
      }
      const at = stops[last], when = times[last];
      if (at.role === "POD" || last === stops.length - 1) {
        const pod = n && n.podStatus && !/verified/i.test(n.podStatus) ? " · POD " + n.podStatus : "";
        return { step: `Delivered at ${at.name} ${hhmm(when)}${pod}`, short: `📍 Delivered ${shortName(at.name)} ${hhmm(when)}${n ? podMark(n.podStatus) : ""}`,
          alerts, stage: "delivered", stops: stopList, who };
      }
      const next = stops[last + 1];
      const ageMin = (now - when) / 60000;
      const step = `Leg ${last + 1}/${legs}: ${at.name} (${hhmm(when)}) → ${next ? next.name : "?"} · ${who}`;
      if (ageMin >= stuckMin) alerts.push({ code: "stuck", text: `Stuck: at ${at.name} since ${hhmm(when)} (${fmtMin(ageMin)}), not at ${next ? next.name : "next stop"} yet`,
        short: `🛑 Stuck at ${shortName(at.name)} ${fmtMin(ageMin)}` });
      return { step, short: `🚚 ${last + 1}/${legs} → ${shortName(next ? next.name : "?")} · ${first}`, alerts, stage: "moving", stops: stopList, who,
        remaining: stops.length - 1 - last, lastAt: when, nextName: next ? next.name : "" };
    }
    if (n) {
      if (/complete/i.test(n.moveStatus)) {
        return { step: `✓ Completed${n.end ? " " + n.end : ""}${n.podStatus && !/verified/i.test(n.podStatus) ? " · POD " + n.podStatus : " · POD verified"}`,
          short: `✓ Done${n.end ? " " + dayMon(n.end) : ""}${podMark(n.podStatus || "Verified")}`, alerts, stage: "completed", stops: [] };
      }
      return { step: `${n.moveStatus || "In Nexl"} (not on driver tracking)`, short: `• ${n.moveStatus || "In Nexl"}`, alerts, stage: "other", stops: [] };
    }
    return { step: "", short: "", alerts, stage: "", stops: [] };
  }

  // =====================================================================
  // Safety checks: vessel cutoffs, genset, double allocation
  // =====================================================================
  const MON = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, SEPT: 8, OCT: 9, NOV: 10, DEC: 11 };
  /** Excel serial (46302.29) or "26-OCT-03 1800" -> local Date */
  function cellDate(v) {
    if (typeof v === "number" && v > 30000) {
      const ms = Math.round((v - 25569) * 86400000), d = new Date(ms);
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes());
    }
    const m = /^(\d{2})-([A-Z]{3,4})-(\d{2})\s+(\d{2}):?(\d{2})$/i.exec(String(v || "").trim());
    if (m && MON[m[2].toUpperCase()] !== undefined) return new Date(2000 + +m[1], MON[m[2].toUpperCase()], +m[3], +m[4], +m[5]);
    return null;
  }
  /** STACK DATES tab (any number of terminal blocks, each with its own header row). */
  function parseStackDates(grid) {
    const out = [];
    let H = null, terminal = "";
    for (const row of grid || []) {
      const cells = (row || []).map((c) => clean(c).toUpperCase());
      const vi = cells.indexOf("VESSEL NAME");
      if (vi >= 0) {
        H = { vessel: vi, dry: cells.findIndex((c) => /^(DRY CUTOFF|STACKS CLOSED)$/.test(c)), reefer: cells.findIndex((c) => /^(REEFER CUTOFF|REEFERS CLOSED)$/.test(c)), visit: cells.indexOf("VISIT") };
        continue;
      }
      if (!H) continue;
      const vessel = clean(row[H.vessel]);
      if (!vessel || /STACK DATES/i.test(vessel)) continue;
      const visit = H.visit >= 0 ? clean(row[H.visit]) : "";
      terminal = /^PLZ/i.test(visit) ? "PECT" : /^NCT/i.test(visit) ? "NCT" : terminal;
      out.push({ vessel, key: compact(vessel), terminal, dry: H.dry >= 0 ? cellDate(row[H.dry]) : null, reefer: H.reefer >= 0 ? cellDate(row[H.reefer]) : null });
    }
    return out;
  }
  /** DATA - TRANSPORTER tab -> { KEY: {name, genset:"YES"|"NO"|"", drivers:Set} } */
  function parseTransporters(grid) {
    const out = {};
    if (!grid || !grid.length) return out;
    const head = grid[0].map((c) => clean(c).toUpperCase());
    const gi = head.indexOf("GENSET");
    for (const row of grid.slice(1)) {
      const name = clean(row && row[0]);
      if (!name) continue;
      const k = compact(name);
      const drivers = new Set(row.slice(1, gi >= 0 ? gi : row.length).flatMap((d) => tokens(d, 3)));
      const prev = out[k];
      out[k] = { name, genset: gi >= 0 ? clean(row[gi]).toUpperCase() : (prev ? prev.genset : ""), drivers: prev ? new Set([...prev.drivers, ...drivers]) : drivers };
    }
    return out;
  }
  function findTransporter(map, ...names) {
    const keys = Object.keys(map);
    for (const nm of names) {
      const k = compact(nm);
      if (!k) continue;
      const hit = keys.find((x) => x === k) || keys.find((x) => x.length >= 3 && (k.startsWith(x) || x.startsWith(k)));
      if (hit) return map[hit];
      const tk = tokens(nm, 3)[0];
      if (tk) { const h2 = keys.find((x) => x.startsWith(compact(tk))); if (h2) return map[h2]; }
    }
    return null;
  }
  function vesselStack(stack, vessel, now) {
    const k = compact(vessel);
    if (!k || k.length < 4) return null;
    const cands = stack.filter((s) => s.key === k || s.key.startsWith(k) || k.startsWith(s.key));
    const ref = now.getTime() - 24 * 3600e3;
    return cands.filter((s) => (s.reefer || s.dry) && (s.reefer || s.dry).getTime() > ref)
      .sort((a, b) => (a.reefer || a.dry) - (b.reefer || b.dry))[0] || null;
  }
  const dayHm = (d) => `${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][d.getDay()]} ${hhmm(d)}`;

  /** Typical minutes from one stop to the next, learned from today's driver tracking (median). */
  /** Stop-to-stop times seen in today's tracking, each with an id so the panel can remember them across syncs. */
  function legSamples(tracking) {
    const out = [];
    for (const t of tracking) {
      const stops = parseRoute(t.route);
      let viaN = 0;
      const times = stops.map((st) => st.role === "POL" ? parseEntry(t.pickup) : st.role === "POD" ? parseEntry(t.dropoff) : parseEntry(viaN++ === 0 ? t.via : viaN === 2 ? t.via2 : ""));
      for (let i = 1; i < times.length; i++) {
        if (!times[i] || !times[i - 1]) continue;
        const m = Math.round((times[i] - times[i - 1]) / 60000);
        if (m >= 10 && m <= 600) out.push({ id: compact(t.driver) + "|" + +times[i - 1], m, from: stops[i - 1].name, to: stops[i].name, at: +times[i] });
      }
    }
    return out;
  }
  function typicalLegMinutes(tracking, fallback, remembered) {
    const d = (remembered || []).map((x) => x.m);
    const seen = new Set((remembered || []).map((x) => x.id));
    for (const x of legSamples(tracking)) if (!seen.has(x.id)) d.push(x.m);
    if (d.length < 3) return fallback;
    d.sort((a, b) => a - b);
    return Math.round(d[Math.floor(d.length / 2)]);
  }
  function extrasContext(rowRecs, nexl, instrById, nexlByContainer, now, opts) {
    const x = opts.extras || {};
    const ctx = {
      now, stack: x.stack ? parseStackDates(x.stack) : [], transporters: x.transporters ? parseTransporters(x.transporters) : {},
      cutoffWarnH: opts.cutoffWarnHours ?? 12, cutoffRedH: 3, silentMin: opts.pingSilentMinutes ?? 30,
      activeByDriver: new Map(), planByContainer: new Map(), nexlByContainer, instrById,
    };
    for (const t of nexl.tracking || []) {
      const st = legStatus(t, null, now, 1e9, 1e9).stage;
      if (st !== "allocated" && st !== "moving") continue;
      const k = compact(t.driver);
      if (!k) continue;
      if (!ctx.activeByDriver.has(k)) ctx.activeByDriver.set(k, []);
      ctx.activeByDriver.get(k).push(t);
    }
    ctx.legMin = typicalLegMinutes(nexl.tracking || [], opts.legMinutes || 90, opts.legHistory);
    for (const r of rowRecs) if (r.key) { if (!ctx.planByContainer.has(r.key)) ctx.planByContainer.set(r.key, []); ctx.planByContainer.get(r.key).push(r); }
    return ctx;
  }
  const rowVal = (r, f) => (r.cols[f] === undefined ? "" : clean(r.vals[r.cols[f]]));
  const isReeferRow = (r) => r.cutoffKind === "reefer" || (r.cutoffKind === "auto" && /R[HFE]?\b|REEF/i.test(rowVal(r, "equipment")));

  function extraAlerts(r, n, ls, X) {
    const out = [];
    const finished = ls.stage === "delivered" || ls.stage === "completed";
    // 1. Vessel cutoff countdown (exports only)
    if (r.cutoffKind && X.stack.length) {
      const v = vesselStack(X.stack, rowVal(r, "vessel") || (r.group.nexl[0] || {}).vessel, X.now);
      const reefer = isReeferRow(r);
      const cut = v && (reefer ? v.reefer || v.dry : v.dry || v.reefer);
      if (cut) {
        const h = (cut - X.now) / 3600e3, kind = reefer ? "reefer" : "dry";
        r.cutoff = { vessel: v.vessel, terminal: v.terminal, kind, at: cut, hoursLeft: h };
        const eta = finished ? null : predictEta(r, ls, X);
        r.cutoff.eta = eta;
        r.cutoff.late = !!(eta && eta > cut);
        const etaTxt = eta ? `ETA ~${dayHm(eta)}` : "";
        if (finished) { /* already delivered: counted in the summary, no alert */ }
        else if (h >= 0 && h <= 24 && eta && eta > cut) {
          const lateBy = fmtMin((eta - cut) / 60000);
          out.push({ code: "late", severity: "error",
            text: `Likely to MISS ${v.vessel} ${kind} cutoff (${v.terminal} ${dayHm(cut)}): ${ls.stage === "moving" ? "truck " + ls.remaining + " stop(s) from port" : ls.stage === "allocated" ? "truck not started yet" : "no truck on it yet"}, ${etaTxt} (${lateBy} late)`,
            short: `🚨 Will miss cutoff: ${etaTxt} vs ${hhmm(cut)}` });
        }
        else if (h < 0 && h > -24) out.push({ code: "cutoff", severity: "error", text: `Missed ${kind} cutoff for ${v.vessel} (${v.terminal} ${dayHm(cut)}) and not delivered`, short: `⚓ Cutoff missed ${dayHm(cut)}` });
        else if (h >= 0 && h <= X.cutoffWarnH) out.push({ code: "cutoff", severity: h <= X.cutoffRedH ? "error" : "warn",
          text: `${v.vessel} ${kind} cutoff in ${fmtMin(h * 60)} (${v.terminal} ${dayHm(cut)}), container not at port yet${eta ? ` · ${etaTxt}, should make it` : ""}`, short: `⚓ Cutoff in ${fmtMin(h * 60)}${eta ? " · ETA " + hhmm(eta) + " ✓" : ""}` });
      }
    }
    // 2. Genset mismatch
    if (!finished && r.cols.genset !== undefined && /^Y/i.test(rowVal(r, "genset"))) {
      const tr = findTransporter(X.transporters, rowVal(r, "transporter"), n && n.owner, r.tracking && r.tracking.owner);
      if (tr && tr.genset === "NO") out.push({ code: "genset", severity: "error", text: `Genset required but ${tr.name} has no genset (DATA - TRANSPORTER)`, short: `🔌 Needs genset, ${tr.name} has none` });
    }
    // (No "phone silent" alert: many routes have no signal, so it only cluttered the screen.
    //  The last ping is still shown in the row details.)
    // 4. Driver on two active jobs at once
    if (r.tracking && (ls.stage === "allocated" || ls.stage === "moving")) {
      const others = (X.activeByDriver.get(compact(r.tracking.driver)) || []).filter((t) => t !== r.tracking);
      if (others.length) {
        const list = [...new Set(others.map((t) => t.instruction))].join(", ");
        out.push({ code: "dbldriver", severity: "warn", text: `${ls.who || r.tracking.driver} is also allocated to ${list}, so one load will wait`, short: `👥 Driver also on ${list}` });
      }
    }
    // 5. Container on two live jobs
    if (r.key && !finished) {
      const planDup = (X.planByContainer.get(r.key) || []).filter((o) => o !== r && o.base !== r.base);
      const nexlDup = (X.nexlByContainer.get(r.key) || []).filter((c) => baseInstr(c.instruction) !== r.base && !/complete/i.test(c.moveStatus || "")
        && (X.instrById.get(c.instruction) || {}).state !== "Completed");
      const list = [...new Set([...planDup.map((o) => o.id), ...nexlDup.map((c) => c.instruction)])];
      if (list.length) out.push({ code: "dblcontainer", severity: "warn", text: `Container ${r.container} is also on live instruction ${list.join(", ")}`, short: `🔁 Also on ${list.join(", ")}` });
    }
    return out;
  }

  /**
   * When will this load reach the port? Uses the truck's live position and today's typical
   * stop-to-stop time (X.legMin). Not started: +30 min to get going. No truck yet: +60 min to allocate.
   */
  function predictEta(r, ls, X) {
    const leg = X.legMin * 60000, now = X.now.getTime();
    if (ls.stage === "moving" && !ls.lastAt) return new Date(now + Math.max(ls.remaining || 1, 1) * leg); // loaded, no stop times
    if (ls.stage === "moving" && ls.lastAt) {
      const rem = Math.max(ls.remaining, 1);
      const stuck = ls.alerts.some((x) => x.code === "stuck") || now - ls.lastAt.getTime() > 2 * leg;
      if (stuck) return new Date(now + rem * leg); // not moving: every remaining leg still to drive
      return new Date(Math.max(ls.lastAt.getTime() + rem * leg, now + (rem - 1) * leg + 30 * 60000));
    }
    if (ls.stage === "allocated") return new Date(now + 30 * 60000 + Math.max(ls.remaining || 2, 1) * leg);
    if (!ls.stage || ls.stage === "other") return new Date(now + 60 * 60000 + 2 * leg);
    return null;
  }

  /** Wording on a vessel cutoff card. Only "missed" when containers really are not at port. */
  function cutoffLabel(c) {
    const past = c.hoursLeft < 0, n = c.open.length;
    return { big: past ? (n ? `${n} missed` : "✓ Closed") : n ? fmtMin(c.hoursLeft * 60) : "✓ All in", allIn: !n, past };
  }

  function cutoffSummary(rowRecs, X, rows) {
    const by = new Map();
    for (const r of rowRecs) {
      if (!r.cutoff) continue;
      const d = rows[r.tab + "|" + r.row];
      const done = d && (d.leg.stage === "delivered" || d.leg.stage === "completed");
      const k = r.cutoff.vessel + "|" + r.cutoff.kind;
      if (!by.has(k)) by.set(k, { ...r.cutoff, total: 0, done: 0, open: [] });
      const g = by.get(k);
      g.total++;
      if (done) g.done++; else g.open.push({ tab: r.tab, row: r.row, id: r.id, container: r.container, eta: r.cutoff.eta, late: r.cutoff.late });
    }
    for (const g of by.values()) g.atRisk = g.open.filter((o) => o.late).length;
    return [...by.values()].filter((g) => g.hoursLeft > -24).sort((a, b) => a.at - b.at);
  }

  /** Map config field patterns to column indexes using the header row. */
  function detectColumns(headerValues, fieldPatterns, overrides) {
    const cols = {};
    const heads = headerValues.map((h) => clean(h));
    for (const [field, pats] of Object.entries(fieldPatterns)) {
      if (overrides && overrides[field]) { cols[field] = letterToIndex(overrides[field]); continue; }
      const idx = heads.findIndex((h) => pats.some((p) => p.test(h)));
      if (idx >= 0) cols[field] = idx;
    }
    return cols;
  }
  function letterToIndex(l) {
    let n = 0;
    for (const ch of String(l).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  root.NexlMatcher = { compare, detectColumns, colLetter, letterToIndex, compact, tokens, nameMatch, refMatch, vesselMatch, legStatus, parseRoute, durMin, fmtMin, issueKey, issueFp, instrKey, instrFp, cutoffLabel, shortIssue, parseStackDates, parseTransporters, cellDate };
})(typeof window !== "undefined" ? window : globalThis);
