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
  function nameMatch(a, b) {
    const ca = compact(a), cb = compact(b);
    if (!ca || !cb) return true;
    if (ca === cb || ca.includes(cb) || cb.includes(ca)) return true;
    // Abbreviations: "IOEC" = "Indian Ocean Export Company"
    const initials = (s) => clean(s).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean).map((w) => w[0]).join("");
    if (ca.length >= 2 && (initials(b).startsWith(ca) || initials(a).startsWith(cb))) return true;
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
  function compare(sheetTabs, nexl) {
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
        if (!key) { g.openSlots++; stats.openSlots++; continue; }
        stats.rowsChecked++;

        const entry = { container: cont, id, sheet: cellRef("container"), nexl: null, tracking: trackByContainer.get(key) || null, issues: [] };
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
            add("error", "container", cont, "", `Not on this instruction in Nexl. Latest Nexl record: ${latest.instruction} (${latest.customer}, ${latest.status || "?"})`);
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
        const rowText = new Set(r.values.flatMap((v) => tokens(v, 3)));

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

    const notOnPlan = [...instrById.values()].filter((i) => !groups.has(i.base));
    const order = { error: 0, warn: 1, info: 2 };
    issues.sort((a, b) => order[a.severity] - order[b.severity] || String(a.instruction).localeCompare(String(b.instruction)));
    const groupList = [...groups.values()].map((g) => ({ ...g, tabs: [...g.tabs] })).sort((a, b) => b.base.localeCompare(a.base));
    return { issues, groups: groupList, notOnPlan, stats };
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

  root.NexlMatcher = { compare, detectColumns, colLetter, letterToIndex, compact, tokens, nameMatch, refMatch, vesselMatch };
})(typeof window !== "undefined" ? window : globalThis);
