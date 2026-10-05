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
        out.push({ name: tab.name, compare: tab.compare, cols, rows });
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
    vessel: "Vessel", customer: "Customer", transporter: "Transporter", driver: "Driver" };

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

  root.ExcelIO = { readPlanTabs, goTo, writeCheckTab };
})(window);
