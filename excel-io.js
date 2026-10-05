/*
 * Nexl Check — Excel side: reads only the rows that matter from the plan tabs, jumps to cells,
 * and writes the "NEXL CHECK" discrepancy tab. It never edits the plan tabs themselves.
 */
(function (root) {
  "use strict";
  const M = root.NexlMatcher;
  const P = root.NexlParsers;

  /**
   * Reads the header row + every row whose instruction belongs to `bases` (a Set of base instruction numbers).
   * Returns [{name, compare, cols, rows:[{row, values}], missing?, error?}]
   */
  async function readPlanTabs(cfg, bases) {
    return Excel.run(async (ctx) => {
      const out = [];
      for (const tab of cfg.tabs) {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(tab.name);
        await ctx.sync();
        if (ws.isNullObject) { out.push({ name: tab.name, missing: true, rows: [], cols: {}, compare: tab.compare }); continue; }

        const used = ws.getUsedRangeOrNullObject(true);
        used.load("rowCount,columnCount,rowIndex,columnIndex");
        const head = ws.getRangeByIndexes(cfg.headerRow - 1, 0, 1, cfg.maxColumns);
        head.load("values");
        await ctx.sync();
        if (used.isNullObject) { out.push({ name: tab.name, rows: [], cols: {}, compare: tab.compare }); continue; }

        const cols = M.detectColumns(head.values[0], cfg.fields, tab.overrides);
        if (cols.instruction === undefined || cols.container === undefined) {
          out.push({ name: tab.name, rows: [], cols, compare: tab.compare, error: "Couldn't find the INSTRUCTION or CONTAINER column in row " + cfg.headerRow });
          continue;
        }
        const lastRow = used.rowIndex + used.rowCount; // exclusive, 0-based
        const first = cfg.headerRow; // 0-based index of first data row
        const nCols = Math.min(cfg.maxColumns, Math.max(used.columnIndex + used.columnCount, 1));

        // 1) Instruction column only, in chunks.
        const wanted = [];
        for (let r = first; r < lastRow; r += 5000) {
          const n = Math.min(5000, lastRow - r);
          const rng = ws.getRangeByIndexes(r, cols.instruction, n, 1);
          rng.load("values");
          await ctx.sync();
          rng.values.forEach((v, i) => {
            const id = P.normInstr(v[0]);
            if (id && bases.has(P.baseInstr(id))) wanted.push(r + i);
          });
        }
        // 2) Full rows, as contiguous blocks.
        const blocks = [];
        for (const r of wanted) {
          const b = blocks[blocks.length - 1];
          if (b && r === b.start + b.len) b.len++; else blocks.push({ start: r, len: 1 });
        }
        const loaded = blocks.map((b) => { const rng = ws.getRangeByIndexes(b.start, 0, b.len, nCols); rng.load("values"); return { b, rng }; });
        if (loaded.length) await ctx.sync();
        const rows = [];
        for (const { b, rng } of loaded) rng.values.forEach((vals, i) => rows.push({ row: b.start + i + 1, values: vals }));
        out.push({ name: tab.name, compare: tab.compare, cutoff: tab.cutoff || null, cols, rows });
      }
      return out;
    });
  }

  async function goTo(ref) {
    if (!ref || !ref.tab) return;
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(ref.tab);
      ws.activate();
      ws.getRange((ref.col || "A") + ref.row).select();
      await ctx.sync();
    });
  }

  const safe = (v) => { const s = v == null ? "" : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; };
  const SEV_FILL = { error: "#FFD7D5", warn: "#FFF1C2", info: "#DDEBFF" };
  const SEV_LABEL = { error: "Error", warn: "Check", info: "Info" };
  const FIELD_LABEL = { container: "Container", instruction: "Instruction", seal: "Seal", booking: "Booking ref", loadRef: "Load ref",
    vessel: "Vessel", customer: "Customer", transporter: "Transporter", driver: "Driver", progress: "Progress" };

  let lastSignature = "";

  /** Writes / refreshes the NEXL CHECK tab. Rewrites the list only when it changed. */
  async function writeCheckTab(cfg, issues, firstSeen, stamp) {
    const list = issues;
    const sig = JSON.stringify(list.map((i) => [i.severity, i.ref && i.ref.tab, i.ref && i.ref.row, i.field, i.sheet, i.nexl, i.message]));
    await Excel.run(async (ctx) => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject(cfg.checkTabName);
      await ctx.sync();
      const created = ws.isNullObject;
      if (created) { ws = ctx.workbook.worksheets.add(cfg.checkTabName); }
      ws.getRange("A1").values = [[`Nexl Check — last run ${stamp}. ${list.length} item(s). This tab is rewritten by the Nexl Check add-in; edits here are lost.`]];
      ws.getRange("A1").format.font.bold = true;
      if (!created && sig === lastSignature) { await ctx.sync(); return; }

      const old = ws.getUsedRangeOrNullObject(true);
      await ctx.sync();
      if (!old.isNullObject) { const body = ws.getRange("A3:L" + Math.max(3, 3 + 5000)); body.clear(); }

      const header = ["Severity", "Tab", "Cell", "Instruction", "Container", "Field", "Sheet value", "Nexl value", "Issue", "First seen", "Go to"];
      ws.getRange("A3:K3").values = [header];
      const hdr = ws.getRange("A3:K3").format;
      hdr.font.bold = true; hdr.font.color = "#FFFFFF"; hdr.fill.color = "#24292F";

      if (list.length) {
        const vals = list.map((i) => {
          const cell = i.ref ? `${i.ref.col || "A"}${i.ref.row}` : "";
          return [SEV_LABEL[i.severity], i.ref ? i.ref.tab : "", cell, i.instruction, i.container, FIELD_LABEL[i.field] || i.field,
            i.sheet, i.nexl, i.message, firstSeen(i) || ""].map(safe);
        });
        const n = vals.length;
        ws.getRange(`A4:J${3 + n}`).values = vals;
        ws.getRange(`K4:K${3 + n}`).formulas = list.map((i) => [i.ref ? `=HYPERLINK("#'${i.ref.tab.replace(/'/g, "''")}'!${i.ref.col || "A"}${i.ref.row}","Go")` : ""]);
        list.forEach((i, k) => { ws.getRange(`A${4 + k}`).format.fill.color = SEV_FILL[i.severity]; });
      }
      ws.freezePanes.freezeRows(3);
      ws.getRange("A3:K" + (3 + Math.max(list.length, 1))).format.autofitColumns();
      ws.getRange("A1").format.columnWidth = 70;
      await ctx.sync();
      lastSignature = sig;
    });
  }

  // ---------------------------------------------------------------------------
  // On-sheet indicators: two status columns at the end of each plan tab + notes on wrong cells.
  // Only the add-in's own columns and its own notes (text starting "NEXL CHECK") are ever changed.
  // ---------------------------------------------------------------------------
  const LEVEL_STYLE = {
    error: { fill: "#FFC7CE", font: "#9C0006" },
    warn: { fill: "#FFEB9C", font: "#7F4F00" },
    ok: { fill: "#C6EFCE", font: "#006100" },
  };
  const NOTE_PREFIX = "NEXL CHECK";
  const colCache = new Map(); // tab -> step column index

  async function findStatusColumns(ctx, ws, cfg) {
    const H = cfg.headerRow - 1;
    const head = ws.getRangeByIndexes(H, 0, 1, 120);
    head.load("values");
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("rowIndex,rowCount");
    await ctx.sync();
    const heads = head.values[0].map((v) => String(v).trim().toUpperCase());
    const existing = heads.indexOf(cfg.statusColumns.stepHeader.toUpperCase());
    if (existing >= 0) return { c: existing, created: false };
    const lastRow = used.isNullObject ? cfg.headerRow : used.rowIndex + used.rowCount;
    let c = heads.reduce((last, h, i) => (h ? i : last), -1) + 1;
    // Skip columns that hold unlabelled data (e.g. IMPORTS P.E column P).
    for (let tries = 0; tries < 30; tries++, c++) {
      const probe = ws.getRangeByIndexes(H, c, Math.max(lastRow - H, 1), 2);
      probe.load("values");
      await ctx.sync();
      if (probe.values.every((r) => r[0] === "" && r[1] === "")) return { c, created: true };
    }
    throw new Error("Couldn't find two empty columns for the status columns");
  }

  /**
   * rowStatus: [{tab,row,step,alert,level}] for every plan row in scope.
   * Rows not in rowStatus (instructions no longer active in Nexl) keep their last status.
   */
  async function writeStatusColumns(cfg, rowStatus) {
    const byTab = new Map();
    for (const r of rowStatus) { if (!byTab.has(r.tab)) byTab.set(r.tab, []); byTab.get(r.tab).push(r); }
    let written = 0;
    await Excel.run(async (ctx) => {
      for (const [tab, rows] of byTab) {
        const ws = ctx.workbook.worksheets.getItem(tab);
        const { c, created } = await findStatusColumns(ctx, ws, cfg);
        colCache.set(tab, c);
        const H = cfg.headerRow - 1;
        const hdr = ws.getRangeByIndexes(H, c, 1, 2);
        hdr.values = [[cfg.statusColumns.stepHeader, cfg.statusColumns.alertHeader]];
        hdr.format.font.bold = true; hdr.format.font.color = "#FFFFFF"; hdr.format.fill.color = "#203864";
        hdr.format.horizontalAlignment = "Center";

        // Current contents of our two columns for the rows we own, read as one block.
        const minR = Math.min(...rows.map((r) => r.row)), maxR = Math.max(...rows.map((r) => r.row));
        const cur = ws.getRangeByIndexes(minR - 1, c, maxR - minR + 1, 2);
        cur.load("values");
        await ctx.sync();
        for (const r of rows) {
          const old = cur.values[r.row - minR];
          const want = [r.step || "", r.alert || ""];
          if (old[0] === want[0] && old[1] === want[1]) continue;
          const rng = ws.getRangeByIndexes(r.row - 1, c, 1, 2);
          rng.values = [want];
          const a = ws.getRangeByIndexes(r.row - 1, c + 1, 1, 1);
          const st = LEVEL_STYLE[r.level];
          if (st) { a.format.fill.color = st.fill; a.format.font.color = st.font; }
          else { a.format.fill.clear(); a.format.font.color = "#000000"; }
          written++;
        }
        if (created) { // only size the columns the first time, so people can resize them
          ws.getRangeByIndexes(0, c, 1, 1).format.columnWidth = 260;
          ws.getRangeByIndexes(0, c + 1, 1, 1).format.columnWidth = 300;
        }
        await ctx.sync();
      }
    });
    return written;
  }

  const notesSupported = () => {
    try { return Office.context.requirements.isSetSupported("ExcelApi", "1.18"); } catch (e) { return false; }
  };

  /** Adds/updates/removes the add-in's own notes. issues: those with ref.col on a plan tab. */
  async function syncNotes(cfg, issues, stamp) {
    if (!notesSupported()) return { supported: false };
    const want = new Map(); // "tab|A1" -> text
    for (const i of issues) {
      if (!i.ref || !i.ref.col || i.severity === "info" || i.field === "progress") continue;
      const k = `${i.ref.tab}|${i.ref.col}${i.ref.row}`;
      const line = i.nexl && i.field !== "container" ? `${i.message}. Nexl has: ${String(i.nexl).split(" | ")[0]}` : i.message;
      want.set(k, want.has(k) ? want.get(k) + "\n• " + line : `${NOTE_PREFIX} (${stamp}):\n• ${line}`);
    }
    const tabs = [...new Set(cfg.tabs.map((t) => t.name))];
    let added = 0, removed = 0;
    await Excel.run(async (ctx) => {
      for (const tab of tabs) {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(tab);
        await ctx.sync();
        if (ws.isNullObject) continue;
        const notes = ws.notes;
        notes.load("items/content");
        await ctx.sync();
        const locs = notes.items.map((n) => { const l = n.getLocation(); l.load("address"); return { n, l }; });
        await ctx.sync();
        const existing = new Map(); // A1 -> {note, ours}
        for (const { n, l } of locs) {
          const a1 = l.address.split("!").pop().replace(/\$/g, "");
          existing.set(a1, { n, ours: String(n.content || "").startsWith(NOTE_PREFIX) });
        }
        for (const [a1, e] of existing) {
          if (!e.ours) continue;
          const text = want.get(`${tab}|${a1}`);
          if (!text) { e.n.delete(); removed++; }
          else if (e.n.content.replace(/\(.*?\)/, "") !== text.replace(/\(.*?\)/, "")) e.n.content = text;
        }
        for (const [k, text] of want) {
          const [t, a1] = k.split("|");
          if (t !== tab || existing.has(a1)) continue; // never touch someone else's note
          ws.notes.add(ws.getRange(a1), text);
          added++;
        }
        await ctx.sync();
      }
    });
    return { supported: true, added, removed };
  }

  // ---------------------------------------------------------------------------
  // Fill blanks: writes a Nexl value into a cell ONLY if the cell is still empty right now.
  // ---------------------------------------------------------------------------
  async function fillBlanks(list) {
    let filled = 0, skipped = 0;
    await Excel.run(async (ctx) => {
      const cells = list.map((f) => { const r = ctx.workbook.worksheets.getItem(f.tab).getRange(f.col + f.row); r.load("values"); return { f, r }; });
      await ctx.sync();
      for (const { f, r } of cells) {
        if (String(r.values[0][0]).trim() !== "") { skipped++; continue; }
        r.values = [[f.value]];
        filled++;
      }
      await ctx.sync();
    });
    return { filled, skipped };
  }

  // ---------------------------------------------------------------------------
  // Shared snoozes ("Seen – I'm on it"): a hidden NEXL_ACK sheet so the whole team sees them.
  // ---------------------------------------------------------------------------
  const ACK_SHEET = "NEXL_ACK";
  async function readAcks() {
    const out = {};
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(ACK_SHEET);
      await ctx.sync();
      if (ws.isNullObject) return;
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("values");
      await ctx.sync();
      if (used.isNullObject) return;
      for (const [key, by, until, text] of used.values.slice(1)) {
        const t = Date.parse(until);
        if (key && t > Date.now()) out[key] = { by, until: t, text };
      }
    });
    return out;
  }
  async function writeAck(key, by, minutes, text) {
    await Excel.run(async (ctx) => {
      let ws = ctx.workbook.worksheets.getItemOrNullObject(ACK_SHEET);
      await ctx.sync();
      if (ws.isNullObject) {
        ws = ctx.workbook.worksheets.add(ACK_SHEET);
        ws.visibility = "Hidden";
        ws.getRange("A1:D1").values = [["Key", "By", "Until", "What"]];
      }
      const used = ws.getUsedRange(true);
      used.load("values,rowCount");
      await ctx.sync();
      const now = Date.now();
      // Keep live entries only (drop expired ones and any older entry for this key).
      const keep = used.values.slice(1).filter((r) => r[0] && r[0] !== key && Date.parse(r[2]) > now);
      if (minutes > 0) keep.push([key, by, new Date(now + minutes * 60000).toISOString(), text || ""]);
      ws.getRange(`A2:D${Math.max(used.rowCount, 2) + 1}`).clear("Contents");
      if (keep.length) ws.getRange(`A2:D${keep.length + 1}`).values = keep;
      await ctx.sync();
    });
  }

  // ---------------------------------------------------------------------------
  // Selection tracking: tell the panel which plan row the user clicked.
  // ---------------------------------------------------------------------------
  const selHandlers = [];
  async function watchSelection(tabNames, cb) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) return false;
    await Excel.run(async (ctx) => {
      for (const name of tabNames) {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
        await ctx.sync();
        if (ws.isNullObject) continue;
        selHandlers.push(ws.onSelectionChanged.add(async (ev) => {
          const m = /(\d+)/.exec(String(ev.address).split(":")[0]);
          if (m) cb(name, +m[1]);
        }));
      }
      await ctx.sync();
    });
    return true;
  }

  // ---------------------------------------------------------------------------
  // Auto-open: Office opens the panel whenever this workbook is opened (setting saved in the file).
  // Needs the manifest's TaskpaneId to be Office.AutoShowTaskpaneWithDocument.
  // ---------------------------------------------------------------------------
  function getAutoOpen() {
    try { return !!Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument"); } catch (e) { return false; }
  }
  function setAutoOpen(on) {
    return new Promise((resolve) => {
      try {
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", !!on);
        Office.context.document.settings.saveAsync(() => resolve(true));
      } catch (e) { resolve(false); }
    });
  }

  // ---------------------------------------------------------------------------
  // Extra reference tabs: STACK DATES (cutoffs) and DATA - TRANSPORTER (gensets)
  // ---------------------------------------------------------------------------
  async function readTabByName(ctx, wanted) {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const ws = sheets.items.find((w) => w.name.trim().toUpperCase() === wanted.trim().toUpperCase());
    if (!ws) return null;
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("values,rowIndex,columnIndex");
    await ctx.sync();
    if (used.isNullObject) return [];
    // Re-base to A1 so row/column positions match the sheet.
    const grid = [];
    used.values.forEach((r, i) => { grid[used.rowIndex + i] = Array(used.columnIndex).fill("").concat(r); });
    return Array.from({ length: grid.length }, (_, i) => grid[i] || []);
  }
  async function readExtras(cfg) {
    const out = {};
    await Excel.run(async (ctx) => {
      out.stack = await readTabByName(ctx, cfg.stackDatesTab);
      out.transporters = await readTabByName(ctx, cfg.transportersTab);
    });
    return out;
  }

  /** All values of one tab (used range re-based to A1), or null if the tab doesn't exist. */
  async function readTabValues(name) {
    let out = null;
    await Excel.run(async (ctx) => { out = await readTabByName(ctx, name); });
    return out;
  }

  LEVEL_STYLE.ack = { fill: "#DDEBF7", font: "#1F4E79" };

  root.ExcelIO = { readPlanTabs, goTo, writeCheckTab, writeStatusColumns, syncNotes, fillBlanks, readAcks, writeAck, watchSelection, getAutoOpen, setAutoOpen, readExtras, readTabValues };
})(window);
