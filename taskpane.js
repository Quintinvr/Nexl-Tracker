/*
 * Nexl Check — task pane controller (v1.2).
 */
(function () {
  "use strict";
  const CFG = window.NEXL_CONFIG;
  const VERSION = "2.5.0";
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ---------- settings (per user, this browser) ----------
  const SKEY = "nexlcheck.settings.v1";
  const defaults = { auto: true, minutes: CFG.refreshMinutes, writeTab: true, writeInfo: false, region: CFG.region, disabledTabs: [],
    statusCols: true, notes: true, notStarted: CFG.notStartedMinutes, stuck: CFG.stuckMinutes, name: "", snoozeMin: 60, splash: true, autoOpenSet: false, popUp: true, waDrvSheet: "", waDrvFieldsBy: {},
    cutoffH: CFG.cutoffWarnHours, silent: CFG.pingSilentMinutes,
    waEmoji: true, waFields: { container: false, seal: false, tare: false, booking: false } };
  let settings = { ...defaults };
  try { settings = { ...defaults, ...JSON.parse(localStorage.getItem(SKEY) || "{}") }; } catch (e) { /* storage unavailable */ }
  const saveSettings = () => { try { localStorage.setItem(SKEY, JSON.stringify(settings)); } catch (e) { /* ignore */ } };

  // ---------- state ----------
  const state = { panelId: Math.random().toString(36).slice(2) + Date.now().toString(36), isWriter: false, writerName: "", busy: false, result: null, lastSyncAt: 0, firstSeen: new Map(), view: "now", input: null, acks: {}, selected: null, fillSel: new Set(), photo: new Map(), slips: {}, sealOk: new Set(),
    extras: {}, waGrid: null, waUnticked: new Set(), waEdits: {}, waClient: null };
  const fmtTime = (d) => d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  const fmtStamp = (d) => d.toLocaleString([], { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
  const SEV_ORDER = (i) => (i.severity === "error" ? 0 : i.code === "notstarted" ? 1 : i.code === "stuck" ? 2 : 3);

  // ---------- small UI helpers ----------
  function pill(id, ok, text) {
    const el = $(id);
    el.className = "pill " + (ok === true ? "ok" : ok === false ? "bad" : "");
    el.querySelector("span").textContent = text;
  }
  function banner(html, kind) {
    const b = $("banner");
    if (!html) { b.hidden = true; return; }
    b.hidden = false;
    b.innerHTML = html;
    b.className = "banner " + (kind === "warn" ? "warn" : "");
  }
  function progress(frac, text) {
    const p = $("progress");
    if (text) $("splashStatus").textContent = text;
    if (frac === null) { p.hidden = true; return; }
    p.hidden = false;
    p.querySelector("div").style.width = Math.round(frac * 100) + "%";
    p.querySelector("span").textContent = text;
  }
  let toastT;
  function toast(msg) {
    const t = $("toast");
    t.textContent = msg; t.hidden = false;
    clearTimeout(toastT); toastT = setTimeout(() => (t.hidden = true), 3500);
  }
  function modal(title, html) {
    $("modalTitle").textContent = title;
    $("modalBody").innerHTML = html;
    $("modal").hidden = false;
  }
  const closeModal = () => ($("modal").hidden = true);

  function errorMessage(e) {
    switch (e && e.code) {
      case "NOT_INSTALLED":
        return "The <b>Nexl Check Bridge</b> browser extension isn't reachable. Open this workbook in <b>Excel for the web</b> in Chrome or Edge with the extension installed, then reload.";
      case "NOT_LOGGED_IN":
        return 'Not logged in to Nexl. <a href="https://controller.nexl.online/" target="_blank" rel="noopener">Open Nexl Controller</a>, log in, then press Sync now.';
      case "BRIDGE_RELOADED":
        return "The bridge extension was updated. Reload this Excel tab.";
      case "LAYOUT_CHANGED":
        return "Nexl's screens look different from what Nexl Check expects. The parser may need updating.";
      case "TIMEOUT":
        return "Nexl took too long to respond. It will retry on the next sync.";
      default:
        return "Sync failed: " + esc((e && e.message) || e);
    }
  }

  // ---------- start-up animation ----------
  const splashStart = Date.now();
  let splashGone = false;
  function hideSplash(force) {
    if (splashGone) return;
    const wait = force ? 0 : Math.max(0, 2600 - (Date.now() - splashStart));
    setTimeout(() => {
      splashGone = true;
      const s = $("splash");
      s.classList.add("out");
      setTimeout(() => (s.hidden = true), 450);
    }, wait);
  }

  // ---------- sync ----------
  function matchOpts() {
    return { now: new Date(), notStartedMinutes: settings.notStarted, stuckMinutes: settings.stuck, acks: state.acks,
      extras: state.extras, cutoffWarnHours: settings.cutoffH, pingSilentMinutes: settings.silent, slips: state.slips, legHistory: loadLegs() };
  }

  async function sync() {
    if (state.busy) return;
    state.busy = true;
    $("syncBtn").disabled = true;
    banner(null);
    try {
      progress(0.05, "Connecting to bridge…");
      const ping = await NexlClient.ping();
      if (!ping.ok) { pill("bridgePill", false, "Bridge missing"); throw Object.assign(new Error("bridge"), { code: "NOT_INSTALLED" }); }
      pill("bridgePill", true, "Bridge");

      progress(0.12, "Reading Nexl instructions…");
      let instructions;
      try { instructions = await NexlClient.getInstructions(settings.region); }
      catch (e) { if (e.code === "NOT_LOGGED_IN") pill("nexlPill", false, "Nexl: log in"); throw e; }
      pill("nexlPill", true, "Nexl");
      const bases = new Set(instructions.map((i) => i.base));

      progress(0.22, "Reading plan tabs…");
      const tabsCfg = CFG.tabs.filter((t) => !settings.disabledTabs.includes(t.name));
      // Light on Excel: STACK DATES / DATA - TRANSPORTER are re-read at most every 30 min, and the full
      // PE CITRUS grid for WhatsApp only while the WhatsApp tab is open (or when you open it).
      const t0 = performance.now();
      const needExtras = !state.extrasAt || Date.now() - state.extrasAt > 30 * 60000;
      const [sheetTabs, acks, extras] = await Promise.all([ExcelIO.readPlanTabs({ ...CFG, tabs: tabsCfg }, bases), ExcelIO.readAcks().catch(() => ({})),
        needExtras ? ExcelIO.readExtras(CFG).catch(() => null) : Promise.resolve(null)]);
      state.acks = acks;
      if (extras) { state.extras = extras; state.extrasAt = Date.now(); }
      if (state.view === "wa") state.waGrid = await ExcelIO.readTabValues(CFG.whatsAppTab).catch(() => state.waGrid);
      else state.waGrid = null; // read again when the WhatsApp tab is opened
      state.perf = { read: Math.round(performance.now() - t0) };
      const tabWarnings = sheetTabs.filter((t) => t.missing || t.error).map((t) => `${esc(t.name)}: ${t.missing ? "tab not found" : esc(t.error)}`);
      const usable = sheetTabs.filter((t) => !t.missing && !t.error);

      const onSheet = new Set();
      for (const t of usable) for (const r of t.rows) { const id = NexlParsers.normInstr(r.values[t.cols.instruction]); if (id) onSheet.add(NexlParsers.baseInstr(id)); }
      const ids = instructions.filter((i) => onSheet.has(i.base)).map((i) => i.id);

      const containers = await NexlClient.getContainers(ids, (d, n) => progress(0.25 + 0.55 * (d / n), `Reading Nexl containers ${d}/${n}…`));
      progress(0.82, "Reading driver tracking…");
      const tracking = await NexlClient.getTracking();
      const nexl = { instructions, containers, tracking };

      let res = NexlMatcher.compare(usable, nexl, matchOpts());
      const unknown = [...new Set(res.issues.filter((i) => i.field === "container" && i.sheet && !i.nexl).map((i) => i.container))].slice(0, 25);
      if (unknown.length) {
        progress(0.9, `Searching Nexl for ${unknown.length} unknown container(s)…`);
        nexl.search = await NexlClient.searchContainers(unknown);
        res = NexlMatcher.compare(usable, nexl, matchOpts());
      }
      state.input = { usable, nexl };

      const now = new Date();
      for (const i of res.issues) if (!state.firstSeen.has(i.key)) state.firstSeen.set(i.key, now);
      state.result = res;
      state.lastSyncAt = Date.now();
      $("lastSync").textContent = "Synced " + fmtTime(now);
      if (state.perf) $("versions").title = `Last sync: reading the sheet ${state.perf.read} ms`;
      render();
      hideSplash();
      runPhotoChecks();
      runSlipChecks();
      popOnNewErrors();
      checkForUpdate();
      rememberLegs(res.legSamples);

      // Only ONE panel writes to the workbook (see ExcelIO.claimWriter); everyone else just reads.
      const w = await ExcelIO.claimWriter(state.panelId, settings.name || "a team member").catch(() => ({ isWriter: false, holder: "?" }));
      state.isWriter = w.isWriter; state.writerName = w.holder;
      showWriter();
      if (!state.isWriter) { if (tabWarnings.length) banner(tabWarnings.join("<br>"), "warn"); return; }

      if (settings.writeTab) {
        progress(0.96, `Updating ${CFG.checkTabName} tab…`);
        const list = res.issues.filter((i) => !i.bypass && (settings.writeInfo || i.severity !== "info"));
        try {
          await ExcelIO.writeCheckTab(CFG, list, (i) => { const d = state.firstSeen.get(i.key); return d ? fmtStamp(d) : ""; }, fmtStamp(now));
        } catch (e) { tabWarnings.push(`Couldn't update the ${esc(CFG.checkTabName)} tab (${esc(e.message)}). The panel is still up to date.`); }
      }
      await writeColumns(tabWarnings);
      if (settings.notes) {
        progress(0.98, "Updating cell notes…");
        try {
          const n = await ExcelIO.syncNotes(CFG, res.issues.filter((i) => !i.bypass), fmtTime(now));
          if (!n.supported) $("notesHint").hidden = false;
        } catch (e) { tabWarnings.push(`Couldn't update cell notes (${esc(e.message)}).`); }
      }
      if (tabWarnings.length) banner(tabWarnings.join("<br>"), "warn");
    } catch (e) {
      banner(errorMessage(e));
      hideSplash(true);
    } finally {
      progress(null);
      state.lastSyncAt = Date.now(); // also after failures, so auto-refresh waits a full interval before retrying
      state.busy = false;
      $("syncBtn").disabled = false;
    }
  }

  function showWriter() {
    const el = $("writerPill"); if (!el) return;
    el.hidden = false;
    el.textContent = state.isWriter ? "✍ updates the sheet" : `📖 read-only`;
    el.title = state.isWriter ? "This panel writes NEXL STEP / NEXL ALERT, notes and the NEXL CHECK tab for the whole team."
      : `${state.writerName || "Another panel"} writes the sheet. This panel only reads, so Excel isn't overloaded.`;
  }
  async function writeColumns(warnings) {
    if (!state.isWriter) return;
    if (!settings.statusCols || !state.result || !state.result.rowStatus.length) return;
    progress(0.97, "Updating NEXL STEP / NEXL ALERT columns…");
    try { await ExcelIO.writeStatusColumns(CFG, state.result.rowStatus); }
    catch (e) { (warnings || []).push(`Couldn't update the status columns (${esc(e.message)}).`); }
  }

  /** Re-run the comparison on the last Nexl data (after a snooze or fill) without fetching again. */
  async function recompute() {
    if (!state.input) return;
    state.result = NexlMatcher.compare(state.input.usable, state.input.nexl, matchOpts());
    render();
    await writeColumns();
    progress(null);
  }

  // ---------- actions ----------
  async function ensureName() {
    if (settings.name) return settings.name;
    return new Promise((resolve) => {
      modal("Who's on it?", `<p class="small">Your name is shown to the team next to alerts you snooze or bypass.</p>
        <input id="nameIn" type="text" maxlength="24" placeholder="Your name" class="full">
        <div class="row-actions"><button id="nameOk" class="primary small-btn" type="button">Save</button></div>`);
      const go = () => { const v = $("nameIn").value.trim(); if (!v) return; settings.name = v; $("sName").value = v; saveSettings(); closeModal(); resolve(v); };
      $("nameOk").onclick = go;
      $("nameIn").onkeydown = (e) => { if (e.key === "Enter") { e.preventDefault(); go(); } };
      $("nameIn").focus();
    });
  }

  async function snooze(issue, minutes) {
    const by = await ensureName();
    const m = minutes == null ? settings.snoozeMin : minutes;
    try {
      const wide = !m && issue.ack && /^INSTR\|/.test(issue.ack.key || "");
      await ExcelIO.writeAck(!m && issue.ack && issue.ack.key ? issue.ack.key : issue.key, by, m, issue.message);
      state.acks = await ExcelIO.readAcks();
      toast(m ? `Snoozed for ${m} min — the team sees "👀 ${by}"` : wide ? `Bypass undone on every row of ${issue.instruction}` : issue.bypass ? "Bypass undone — the alert is back" : "Snooze removed");
      await recompute();
      if (!m && issue.bypass && settings.notes && state.isWriter) await ExcelIO.syncNotes(CFG, state.result.issues.filter((x) => !x.bypass), fmtTime(new Date())).catch(() => {});
    } catch (e) { toast("Couldn't save the snooze: " + e.message); }
  }

  // Bypass = "this is OK / resolved": the alert is cleared for everyone (status column, notes, NEXL CHECK)
  // and stays cleared until the sheet or Nexl value changes, then it is flagged again.
  const BYPASS_REASONS = ["Sheet is correct", "Nexl will be corrected", "Agreed with client / transporter", "Known exception"];
  async function bypass(list) {
    list = list.filter((i) => !i.ack || !i.bypass);
    if (!list.length) return;
    const by = await ensureName();
    const what = list.length === 1 ? esc(list[0].message) : `${list.length} alerts on this row`;
    // Same alert on other rows of this instruction (e.g. vessel differs on every row of 86888)?
    const one = list.length === 1 ? list[0] : null;
    const iKey = one && one.instruction && one.field !== "progress" ? NexlMatcher.instrKey(one) : null;
    const same = iKey ? state.result.issues.filter((x) => !x.ack && x.instruction && NexlMatcher.instrKey(x) === iKey && NexlMatcher.instrFp(x) === NexlMatcher.instrFp(one)) : [];
    const base = one ? String(one.instruction).split(".")[0] : "";
    modal("Bypass alert", `<p class="small"><b>${what}</b></p>
      <p class="small muted">The alert is cleared for the whole team and comes back only if the sheet or Nexl value changes.</p>
      ${iKey ? `<label class="small wide-opt"><input type="checkbox" id="bpWide" ${same.length > 1 ? "checked" : ""}>
        Bypass this <b>${esc(FIELD[one.field] || one.field)}</b> alert on <b>every row of instruction ${esc(base)}</b>${same.length > 1 ? ` (${same.length} rows now)` : ""} where the sheet says <b>${esc(one.sheet || "(blank)")}</b>, including rows added later</label>` : ""}
      <div class="reasons">${BYPASS_REASONS.map((r, k) => `<label class="small"><input type="radio" name="bpWhy" value="${esc(r)}" ${k === 0 ? "checked" : ""}> ${esc(r)}</label>`).join("")}
      <label class="small"><input type="radio" name="bpWhy" value=""> Other:</label></div>
      <input id="bpText" type="text" maxlength="80" placeholder="Note (optional)" class="full">
      <div class="row-actions"><button id="bpOk" class="primary small-btn" type="button">✔ Bypass</button>
      <button id="bpCancel" class="ghost small-btn" type="button">Cancel</button></div>`);
    $("bpCancel").onclick = closeModal;
    $("bpText").onkeydown = (e) => { if (e.key === "Enter") { e.preventDefault(); $("bpOk").click(); } };
    $("bpText").oninput = () => { if ($("bpText").value) { const o = document.querySelector('input[name="bpWhy"][value=""]'); if (o && !document.querySelector('input[name="bpWhy"]:checked').value) o.checked = true; } };
    $("bpOk").onclick = async () => {
      const pick = (document.querySelector('input[name="bpWhy"]:checked') || {}).value || "";
      const note = $("bpText").value.trim();
      const reason = [pick, note].filter(Boolean).join(" — ").replace(/^[=+\-@]/, "'$&");
      if (!reason) { $("bpText").focus(); toast("Add a short note for 'Other'."); return; }
      const wideChecked = !!(iKey && $("bpWide") && $("bpWide").checked);
      closeModal();
      try {
        const wide = wideChecked;
        if (wide) await ExcelIO.writeAck(iKey, by, 0, reason, { kind: "bypass", fp: NexlMatcher.instrFp(one) });
        else for (const i of list) await ExcelIO.writeAck(i.key, by, 0, reason, { kind: "bypass", fp: NexlMatcher.issueFp(i) });
        state.acks = await ExcelIO.readAcks();
        toast(wide ? `✔ Bypassed on every row of ${base}${same.length > 1 ? ` (${same.length} rows)` : ""}` : `✔ Bypassed ${list.length === 1 ? "alert" : list.length + " alerts"} — the team sees "✔ ${by}"`);
        await recompute();
        if (settings.notes && state.isWriter) await ExcelIO.syncNotes(CFG, state.result.issues.filter((x) => !x.bypass), fmtTime(new Date())).catch(() => {});
      } catch (e) { toast("Couldn't save the bypass: " + e.message); }
    };
  }

  // Which Nexl screen an instruction lives on (Active / Completed), from the last sync.
  function screenFor(id) {
    const list = (state.input && state.input.nexl && state.input.nexl.instructions) || [];
    const b = String(id || "").split(".")[0];
    const hit = list.find((x) => x.id === String(id)) || list.find((x) => x.base === b);
    return hit && /complete/i.test(hit.state || "") ? "completed" : "active";
  }
  async function openNexl(word, filter, instruction) {
    const w = String(word || instruction || "").trim();
    if (!w) return;
    const ins = /^\d{3,8}(\.\d{1,2})?$/.test(String(instruction || "")) ? String(instruction) : null;
    toast(ins ? `Opening ${ins} in Nexl…` : `Searching Nexl for ${w}…`);
    const r = await NexlClient.openInNexl(filter, w, ins, ins ? screenFor(ins) : null);
    if (r.ok) {
      if (ins && r.how === "search") toast(`${ins} isn't on Active, Completed or Finance — showing Nexl's search instead.`);
      return;
    }
    if (r.error === "NOT_LOGGED_IN") { toast("Log in to Nexl first, then try again."); return; }
    // Older bridge (1.0) or blocked: copy + open Nexl so the user can paste into Nexl's search.
    try { await navigator.clipboard.writeText(w); } catch (e) { /* ignore */ }
    window.open("https://controller.nexl.online/php/landing.php", "_blank", "noopener");
    toast(`Copied ${w} — paste it into Nexl's search. (Update the bridge extension for one-click search.)`);
  }

  async function doFill(list) {
    const blocked = list.filter((f) => !fillAllowed(f));
    list = list.filter(fillAllowed);
    if (blocked.length) toast(`${blocked.length} container number(s) skipped: photo not confirmed yet.`);
    if (!list.length) return;
    try {
      const r = await ExcelIO.fillBlanks(list);
      toast(`Filled ${r.filled} cell(s)${r.skipped ? `, skipped ${r.skipped} that already had a value` : ""}.`);
      state.fillSel.clear();
      setTimeout(sync, 400);
    } catch (e) { toast("Couldn't fill: " + e.message); }
  }

  // ---------- rendering ----------
  const FIELD = { container: "Container", instruction: "Instruction", seal: "Seal", booking: "Booking ref", loadRef: "Load ref",
    vessel: "Vessel", customer: "Customer", transporter: "Transporter", driver: "Driver", progress: "Now" };

  // ---------- container photo check ----------
  // A container number from the Nexl app is only offered with "Apply" once the driver's
  // CONTAINER photo in Nexl has been read and shows the same number.
  const photoKey = (f) => `${f.nexlRowId || "-"}|${f.value}`;
  const photoOf = (f) => (f.needsPhoto ? state.photo.get(photoKey(f)) || { status: "checking" } : null);
  const fillAllowed = (f) => !f.needsPhoto || (f.photoKind === "seal" ? state.sealOk.has(photoKey(f)) : photoOf(f).status === "match");
  const PHOTO_BADGE = {
    checking: ["⏳", "Checking photo…", "chk"], match: ["✓", "Photo matches", "ok"], mismatch: ["✗", "Photo shows a different number", "bad"],
    unreadable: ["?", "Photo unclear — check by eye", "warn"], invalid: ["✗", "App number looks mistyped", "bad"],
    manual: ["👁", "Check the seal photo", "chk"], sealok: ["✓", "Seal photo checked", "ok"],
    nophoto: ["📷", "No photo uploaded yet", "warn"], norow: ["?", "Can't find this container's photos in Nexl", "warn"], error: ["!", "Photo check failed", "warn"],
  };
  function photoBadge(f) {
    const p = f.photoKind === "seal" && state.sealOk.has(photoKey(f)) ? { status: "sealok" } : photoOf(f); if (!p) return "";
    const [ic, txt, cls] = PHOTO_BADGE[p.status] || PHOTO_BADGE.error;
    return `<span class="pbadge ${cls}" title="${esc(p.detail || txt)}">${ic} ${esc(txt)}</span>`;
  }
  function runPhotoChecks() {
    if (!state.result || !window.NexlPhotoCheck) return;
    for (const f of state.result.fills) {
      if (!f.needsPhoto) continue;
      const k = photoKey(f);
      if (state.photo.has(k)) continue;
      if (!f.nexlRowId) { state.photo.set(k, { status: "norow", detail: "Nexl didn't give a row for this container, so its photos can't be opened" }); continue; }
      state.photo.set(k, { status: "checking" });
      (f.photoKind === "seal" ? NexlPhotoCheck.sealPhoto(f.nexlRowId).catch((e) => ({ status: "error", detail: String(e.message || e) })) : NexlPhotoCheck.check(f)).then((r) => {
        state.photo.set(k, r);
        if (r.status === "error") setTimeout(() => state.photo.get(k) === r && state.photo.delete(k), 60000); // retry on a later sync
        renderFills();
        if (state.selected) showDetail(state.selected.tab, state.selected.row, true);
      });
    }
  }
  // Leg times remembered on this PC (last 300 stop-to-stop times) so "Will it make the vessel?" learns your routes.
  const LEG_KEY = "nexlcheck.legs.v1";
  function loadLegs() { try { return JSON.parse(localStorage.getItem(LEG_KEY) || "[]"); } catch (e) { return []; } }
  function rememberLegs(samples) {
    if (!samples || !samples.length) return;
    try {
      const have = loadLegs(), ids = new Set(have.map((x) => x.id));
      const merged = have.concat(samples.filter((x) => !ids.has(x.id)).map((x) => ({ id: x.id, m: x.m }))).slice(-300);
      localStorage.setItem(LEG_KEY, JSON.stringify(merged));
    } catch (e) { /* storage off: fine, falls back to today's data */ }
  }
  // Port slips: for rows with an empty status cell, look for a port slip photo in Nexl.
  async function runSlipChecks() {
    const res = state.result;
    if (!res || !window.NexlPhotoCheck || !res.slipCandidates || !res.slipCandidates.length) return;
    const list = res.slipCandidates.slice(0, 40);
    let found = 0, i = 0;
    const work = async () => {
      while (i < list.length) {
        const c = list[i++];
        try { const up = await NexlPhotoCheck.uploads(c.rowId); if (up.psli && !state.slips[c.key]) { state.slips[c.key] = { path: up.psli }; found++; } }
        catch (e) { if (e.code === "NOT_LOGGED_IN") return; }
      }
    };
    await Promise.all([work(), work(), work()]);
    if (found && state.input) {
      state.result = NexlMatcher.compare(state.input.usable, state.input.nexl, matchOpts());
      render();
      runPhotoChecks();
    }
  }
  async function showSlip(f) {
    modal(`Port slip · ${f.container || f.instruction}`, `<p class="small">Loading the port slip…</p>`);
    try {
      const img = await NexlPhotoCheck.image(f.slipPath);
      modal(`Port slip · ${f.container || f.instruction}`, `<img class="cphoto" src="${esc(img)}" alt="Port slip">
        <p class="small">The driver uploaded a port slip, so this ${f.slipKind === "collected" ? "import was <b>collected from the port</b>" : "export is <b>stacked at the port</b>"}.</p>
        ${f.value && f.col ? `<div class="row-actions"><button id="slApply" class="primary small-btn" type="button">Put ${esc(f.value)} in ${esc(f.tab)} ${esc(f.col + f.row)}</button></div>` : ""}`);
      if ($("slApply")) $("slApply").onclick = () => { closeModal(); doFill([f]); };
    } catch (e) { modal("Port slip", `<p class="small">Couldn't load the photo: ${esc(e.message || e)}</p>`); }
  }
  function showSealPhoto(f) {
    const p = photoOf(f) || {};
    modal(`Seal photo · ${f.container || f.instruction}`, `${p.photo ? `<img class="cphoto" src="${esc(p.photo)}" alt="Seal photo">` : `<p class="small">${esc(p.detail || "No photo loaded.")}</p>`}
      <p class="small">Nexl says the seal is <b class="mono big-seal">${esc(f.value)}</b>. Does the photo show this number?</p>
      <div class="row-actions">${p.photo ? `<button id="seYes" class="primary small-btn" type="button">✓ Yes — put it in ${esc(f.tab)} ${esc(f.col + f.row)}</button>
      <button id="seNo" class="ghost small-btn" type="button">✗ No / can't tell</button>` : ""}</div>`);
    if ($("seYes")) $("seYes").onclick = () => { state.sealOk.add(photoKey(f)); closeModal(); doFill([f]); };
    if ($("seNo")) $("seNo").onclick = () => { closeModal(); toast("Left blank. Check the seal with the driver or in Nexl."); };
  }
  // ---------- Check photo & fix: sheet value ≠ Nexl (seal, or a container with a typo / missing digit) ----------
  // Shows the driver's photo next to both numbers (differences highlighted) and lets the controller pick
  // the right one, or type what the photo shows, and puts it in the sheet cell.
  function diffMark(a, b) { // HTML of a with the characters that aren't in b highlighted
    a = String(a || ""); b = String(b || "");
    const L = Array.from({ length: a.length + 1 }, () => Array(b.length + 1).fill(0));
    for (let i = a.length - 1; i >= 0; i--) for (let j = b.length - 1; j >= 0; j--)
      L[i][j] = a[i].toUpperCase() === b[j].toUpperCase() ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    let i = 0, j = 0, out = "";
    while (i < a.length) {
      if (j < b.length && a[i].toUpperCase() === b[j].toUpperCase()) { out += esc(a[i]); i++; j++; }
      else if (j < b.length && L[i][j + 1] > L[i + 1][j]) { out += `<span class="dgap" title="missing here">▾</span>`; j++; }
      else { out += `<mark>${esc(a[i])}</mark>`; i++; }
    }
    if (j < b.length) out += `<span class="dgap" title="missing here">▾</span>`;
    return out;
  }
  const isoHint = (v) => {
    const c = String(v || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!/^[A-Z]{4}\d{7}$/.test(c)) return `<span class="pbadge bad">✗ ${c.length} characters (should be 4 letters + 7 digits)</span>`;
    return NexlPhotoCheck.iso6346Valid(c) ? `<span class="pbadge ok">✓ valid number</span>` : `<span class="pbadge bad">✗ check digit wrong</span>`;
  };
  async function fixWithPhoto(i) {
    if (!i) return;
    const isSeal = i.fix === "seal" || i.field === "seal";
    const what = isSeal ? "Seal" : "Container";
    const d = state.result.rows[i.ref.tab + "|" + i.ref.row];
    const rowId = i.nexlRowId || (d && d.nexl && d.nexl.rowId) || "";
    const cell = `${i.ref.tab} ${i.ref.col}${i.ref.row}`;
    const nexlVal = String(i.nexl || "").split(" | ")[0];
    const title = `Check & fix ${what.toLowerCase()} · ${cell}`;
    modal(title, `<p class="small">Loading the ${what.toLowerCase()} photo from Nexl…</p>`);
    let r = { status: "norow", detail: "Can't find this container's photos in Nexl" };
    if (rowId && window.NexlPhotoCheck) {
      r = await (isSeal ? NexlPhotoCheck.sealPhoto(rowId) : NexlPhotoCheck.check({ value: nexlVal, nexlRowId: rowId }))
        .catch((e) => ({ status: "error", detail: String(e.message || e) }));
    }
    if ($("modal").hidden) return; // closed while loading
    let verdict = "";
    if (!isSeal && r.photo) {
      const sheetOnPhoto = r.text ? NexlPhotoCheck.verifyContainer(r.text, i.sheet).status === "match" : false;
      verdict = r.status === "match" ? `<span class="pbadge ok">📷 Photo matches Nexl</span>`
        : sheetOnPhoto ? `<span class="pbadge ok">📷 Photo matches the sheet</span>`
        : r.status === "mismatch" ? `<span class="pbadge bad">📷 Photo shows ${esc(r.seen || "another number")}</span>`
        : `<span class="pbadge warn">📷 Couldn't read the number clearly — check the photo by eye</span>`;
    }
    const best = !isSeal && r.status === "match" ? nexlVal : "";
    modal(title, `${r.photo ? `<img class="cphoto zoomable" id="fxImg" src="${esc(r.photo)}" alt="${what} photo" title="Click to zoom">` : `<p class="small pbadge warn">📷 ${esc(r.detail || "No photo uploaded yet.")}</p>`}
      ${verdict ? `<p class="small">${verdict}</p>` : ""}
      <div class="fxvals">
        <div><span class="muted small">On the sheet</span><b class="mono big-seal">${diffMark(i.sheet, nexlVal)}</b>${isSeal ? "" : isoHint(i.sheet)}</div>
        <div><span class="muted small">In Nexl</span><b class="mono big-seal">${diffMark(nexlVal, i.sheet)}</b>${isSeal ? "" : isoHint(nexlVal)}</div>
      </div>
      <p class="small">Which one does the photo show?</p>
      <div class="row-actions">
        <button id="fxNexl" class="${best || isSeal ? "primary" : "ghost"} small-btn" type="button">✓ Nexl is right — put ${esc(nexlVal)} in ${esc(i.ref.col + i.ref.row)}</button>
        <button id="fxSheet" class="ghost small-btn" type="button">Sheet is right (${esc(i.sheet)})</button>
      </div>
      <p class="small muted" style="margin:10px 0 4px">Photo shows something else? Type it:</p>
      <div class="fxtype"><input id="fxVal" class="mono" value="${esc(nexlVal)}" autocomplete="off" spellcheck="false"><button id="fxPut" class="ghost small-btn" type="button">Put in sheet</button></div>
      <p class="small" id="fxWarn"></p>
      <div class="row-actions"><button id="fxOpen" class="link xs" type="button">Open in Nexl ↗</button></div>`);
    if ($("fxImg")) $("fxImg").onclick = () => $("fxImg").classList.toggle("zoom");
    const put = async (v) => {
      closeModal();
      const ok = await ExcelIO.replaceCell({ tab: i.ref.tab, row: i.ref.row, col: i.ref.col, value: v }, i.sheet).catch(() => false);
      toast(ok ? `${what} in ${cell} changed to ${v}.` : "The cell changed since the last sync — not overwritten. Sync and try again.");
      if (ok) setTimeout(sync, 400);
    };
    $("fxNexl").onclick = () => put(nexlVal);
    $("fxSheet").onclick = async () => {
      closeModal();
      const by = await ensureName();
      await ExcelIO.writeAck(i.key, by, 0, `${what} photo shows ${i.sheet} (Nexl needs correcting)`, { kind: "bypass", fp: NexlMatcher.issueFp(i) });
      state.acks = await ExcelIO.readAcks();
      toast(`✔ Kept ${i.sheet} on the sheet. Remember to correct Nexl.`);
      await recompute();
    };
    let warned = "";
    $("fxPut").onclick = () => {
      const v = $("fxVal").value.trim().toUpperCase();
      if (!v) return;
      const c = v.replace(/[^A-Z0-9]/g, "");
      const bad = !isSeal && !(/^[A-Z]{4}\d{7}$/.test(c) && NexlPhotoCheck.iso6346Valid(c));
      if (bad && warned !== v) { warned = v; $("fxWarn").innerHTML = `${isoHint(v)} Check it, or press <b>Put in sheet</b> again to use it anyway.`; return; }
      put(isSeal ? v : c);
    };
    $("fxOpen").onclick = () => openNexl(i.container || nexlVal, "container", i.instruction);
  }
  const sealIssuePhoto = fixWithPhoto; // older name
  const fixable = (i) => i && i.fix && !i.bypass && i.severity === "error"; // "I'm on it" can still fix; a bypass means it's OK as is
  function renderFixes() {
    const res = state.result; if (!res) return;
    const list = res.issues.filter(fixable);
    $("fixBox").hidden = !list.length;
    $("nFix").textContent = list.length;
    $("fixList").innerHTML = list.map((i, k) => `<div class="fill"><span>
      <button class="link go-cell" data-go="${k}" type="button" title="Go to this cell on the sheet">📍 ${esc(i.ref.tab)} ${esc(i.ref.col + i.ref.row)}</button> · ${i.fix === "seal" ? "Seal" : "Container"}
      <span class="mono">${esc(i.sheet)}</span> → Nexl <b class="mono">${esc(String(i.nexl).split(" | ")[0])}</b> <span class="muted">(${esc(i.instruction)})</span>
      <br><button class="primary xs" data-fx="${k}" type="button">📷 Check photo & fix</button></span></div>`).join("");
    $("fixList").querySelectorAll("[data-fx]").forEach((b) => (b.onclick = () => fixWithPhoto(list[+b.dataset.fx])));
    $("fixList").querySelectorAll("[data-go]").forEach((b) => (b.onclick = () => { const i = list[+b.dataset.go]; ExcelIO.goTo(i.ref).catch(() => {}); showDetail(i.ref.tab, i.ref.row, true); }));
  }
  function showPhoto(f) {
    if (f.photoKind === "seal") return showSealPhoto(f);
    const p = photoOf(f) || {};
    modal(`Photo · ${f.value}`, `${p.photo ? `<img class="cphoto" src="${esc(p.photo)}" alt="Container photo">` : `<p class="small">No photo loaded.</p>`}
      <p class="small">${photoBadge(f)}</p><p class="small">${esc(p.detail || "")}</p>
      <p class="small muted">App number: <span class="mono">${esc(f.value)}</span>${p.seen ? ` · read on photo: <span class="mono">${esc(p.seen)}</span>` : ""}</p>
      <div class="row-actions">${fillAllowed(f) ? `<button id="phApply" class="primary small-btn" type="button">Apply to ${esc(f.tab)} ${esc(f.col + f.row)}</button>` : ""}
      <button id="phNexl" class="ghost small-btn" type="button">Open in Nexl ↗</button></div>`);
    if ($("phApply")) $("phApply").onclick = () => { closeModal(); doFill([f]); };
    $("phNexl").onclick = () => openNexl(f.instruction, "instruction", f.instruction);
  }
  function renderFills() {
    const res = state.result; if (!res) return;
    const fills = res.fills;
    $("fillBox").hidden = !fills.length;
    $("nFill").textContent = fills.length;
    $("fillList").innerHTML = fills.map((f, k) => {
      const ok = fillAllowed(f);
      if (!ok) state.fillSel.delete(k);
      const p = photoOf(f);
      return `<div class="fill ${f.needsPhoto ? "photo" : ""}">${ok ? `<input type="checkbox" data-k="${k}" ${state.fillSel.has(k) ? "checked" : ""}>` : `<span class="nocb">🔒</span>`}
      <span><button class="link go-cell" data-go="${k}" type="button" title="Go to this cell on the sheet">📍 ${esc(f.tab)} ${esc(f.col + f.row)}</button> · ${esc(FIELD[f.field] || f.field)} ← <button class="link mono go-cell" data-go="${k}" type="button" title="Go to this cell on the sheet">${esc(f.value)}</button>
      <span class="muted">(${esc(f.instruction)} ${esc(f.container || "")})</span>
      ${f.needsPhoto ? `<br>${photoBadge(f)} ${p && p.photo ? `<button class="link xs" data-ph="${k}" type="button">${f.photoKind === "seal" && !ok ? "Check seal photo" : "View photo"}</button>` : ""}
        ${ok ? `<button class="primary xs" data-apply="${k}" type="button">Apply</button>` : ""}` : ""}
      ${f.slipPath ? `<br><span class="pbadge ok">📄 Port slip uploaded</span> <button class="link xs" data-slip="${k}" type="button">View slip</button>` : ""}</span></div>`;
    }).join("");
    $("fillList").querySelectorAll("input").forEach((cb) => cb.onchange = () => { const k = +cb.dataset.k; cb.checked ? state.fillSel.add(k) : state.fillSel.delete(k); });
    $("fillList").querySelectorAll("[data-ph]").forEach((b) => (b.onclick = () => showPhoto(fills[+b.dataset.ph])));
    $("fillList").querySelectorAll("[data-go]").forEach((b) => (b.onclick = (e) => {
      e.preventDefault();
      const f = fills[+b.dataset.go];
      ExcelIO.goTo({ tab: f.tab, row: f.row, col: f.col }).catch(() => {});
      showDetail(f.tab, f.row, true);
    }));
    const ready = fills.filter((f) => f.needsPhoto && fillAllowed(f)).length;
    $("nReady").textContent = ready ? `· ${ready} container${ready === 1 ? "" : "s"} ready to apply ✓` : "";
    $("fillList").querySelectorAll("[data-slip]").forEach((b) => (b.onclick = () => showSlip(fills[+b.dataset.slip])));
    $("fillList").querySelectorAll("[data-apply]").forEach((b) => (b.onclick = () => doFill([fills[+b.dataset.apply]])));
  }

  // ---------- self-update: open panels switch to a new version as soon as GitHub has it ----------
  async function checkForUpdate() {
    try {
      const r = await fetch(`version.json?t=${Date.now()}`, { cache: "no-store" });
      if (!r.ok) return;
      const v = (await r.json()).version;
      if (!v || v === VERSION) return;
      const u = new URL(location.href);
      if (u.searchParams.get("v") === v) return; // already tried this version: don't loop
      banner(`🔄 Nexl Check <b>v${esc(v)}</b> is available — updating in a moment… <button id="updNow" class="primary xs" type="button">Update now</button>`, "info");
      const go = () => { u.searchParams.set("v", v); location.replace(u.toString()); };
      $("updNow").onclick = go;
      setTimeout(() => { if (!state.busy) go(); }, 8000);
    } catch (e) { /* offline: try next sync */ }
  }

  // ---------- team-aligned sync schedule ----------
  const SLOT_CHOICES = [5, 10, 15, 30];
  const slotMinutes = () => (SLOT_CHOICES.includes(+settings.minutes) ? +settings.minutes : 5);
  function nextSlot(from = Date.now()) {
    const ms = slotMinutes() * 60000;
    return Math.floor(from / ms) * ms + ms; // next whole 5-minute mark on the clock
  }
  function showNextSync() {
    const el = $("nextSync"); if (!el) return;
    el.textContent = settings.auto && state.nextSlot ? `· next ${fmtTime(new Date(state.nextSlot))}` : "";
  }

  // ---------- hide panel (shared runtime keeps syncing while hidden) ----------
  const canHide = () => !!(window.Office && Office.addin && typeof Office.addin.hide === "function");
  state.hidden = false;
  async function hidePanel() {
    if (!canHide()) {
      modal("Hide the panel", `<p class="small">To hide the panel and keep Nexl Check syncing, Excel needs the updated add-in file once:</p>
        <ol class="small"><li>Home > Add-ins > More add-ins > My add-ins > "..." next to Nexl Check > Remove.</li>
        <li>Upload <b>NexlCheck-Excel.xml</b> from the new team pack again.</li></ol>
        <p class="small muted">Until then you can close the panel with its ✕ (syncing stops while it is closed).</p>`);
      return;
    }
    try {
      state.knownErrors = openErrorKeys();
      await Office.addin.hide();
      state.hidden = true;
    } catch (e) { toast("Couldn't hide the panel: " + (e.message || e)); }
  }
  function openErrorKeys() {
    const res = state.result;
    return new Set(res ? res.issues.filter((i) => !i.ack && i.severity === "error").map((i) => i.key) : []);
  }
  // While hidden: a NEW red error re-opens the panel (Settings > "Pop up on new errors").
  async function popOnNewErrors() {
    if (!state.hidden || !settings.popUp || !canHide() || !state.knownErrors) return;
    const now = openErrorKeys();
    const fresh = [...now].filter((k) => !state.knownErrors.has(k));
    state.knownErrors = now;
    if (!fresh.length) return;
    try {
      await Office.addin.showAsTaskpane();
      state.hidden = false;
      const first = state.result.issues.find((i) => i.key === fresh[0]);
      toast(`🔔 ${fresh.length} new error${fresh.length > 1 ? "s" : ""}${first ? ": " + (first.short || first.message) : ""}`);
    } catch (e) { /* ignore */ }
  }

  function render() {
    const res = state.result;
    if (!res) return;
    const open = res.issues.filter((i) => !i.ack && i.severity !== "info");
    const rows = Object.values(res.rows);
    $("tErr").textContent = open.filter((i) => i.severity === "error").length;
    $("tWarn").textContent = open.filter((i) => i.severity === "warn").length;
    $("tRoad").textContent = rows.filter((d) => d.leg.stage === "moving" || d.leg.stage === "allocated").length;
    $("tOk").textContent = res.stats.matched;

    const tabSel = $("fTab"), cur = tabSel.value;
    const tabs = [...new Set(res.issues.map((i) => i.ref && i.ref.tab).filter(Boolean))];
    tabSel.innerHTML = '<option value="">All tabs</option>' + tabs.map((t) => `<option${t === cur ? " selected" : ""}>${esc(t)}</option>`).join("");
    renderNow();
    renderIssues();
    renderLive();
    renderWhatsApp();
    if (state.selected) showDetail(state.selected.tab, state.selected.row, true);
  }

  const ackTag = (i) => (i.bypass ? "✔ " : "👀 ") + esc(i.ack.by) + (/^INSTR\|/.test(i.ack.key || "") ? " (whole instr.)" : "");
  function ackButtons(i) {
    if (i.ack) return i.bypass ? `<button class="ghost xs" data-act="unsnooze" title="${esc(i.ack.text || "")}">↩ Undo bypass</button>`
      : `<button class="ghost xs" data-act="unsnooze">Un-snooze</button>`;
    return `<button class="ghost xs" data-act="snooze">👀 I'm on it</button><button class="ghost xs" data-act="bypass" title="Mark as resolved / OK as is">✔ Bypass</button>`;
  }
  function issueCard(i, k, opts = {}) {
    const where = i.ref ? `${esc(i.ref.tab)} · ${esc((i.ref.col || "") + i.ref.row)}` : "";
    const since = state.firstSeen.get(i.key);
    const word = /^\(|^$/.test(i.container || "") ? i.instruction : i.container;
    const filter = word === i.instruction ? "instruction" : "container";
    if (opts.compact) {
      const extra = i.field !== "progress" && i.field !== "container" && i.field !== "instruction" && i.nexl ? ` · Nexl: <b>${esc(String(i.nexl).split(" | ")[0])}</b>` : "";
      return `<div class="card slim ${i.ack ? "acked" : i.severity}" data-k="${k}">
        <div class="what">${esc(i.short || NexlMatcher.shortIssue(i))}</div>
        <div class="where"><span class="mono">${esc(i.instruction)} · ${esc(i.container)}</span><span>${where}</span></div>
        ${extra ? `<div class="small">${extra}</div>` : ""}
        <div class="card-actions">
          <button class="ghost xs" data-act="go">Go</button>
          ${fixable(i) ? `<button class="ghost xs" data-act="sealphoto">📷 Check photo & fix</button>` : ""}
          ${ackButtons(i)}
          <button class="ghost xs" data-act="nexl" data-word="${esc(word)}" data-filter="${filter}">Nexl ↗</button>
          <span class="muted small since">${i.ack ? ackTag(i) : since ? fmtTime(since) : ""}</span>
        </div></div>`;
    }
    return `<div class="card ${i.ack ? "acked" : i.severity}" data-k="${k}">
      <div class="where"><span>${where}</span><span>${i.ack ? ackTag(i) : since ? "since " + fmtTime(since) : ""}</span></div>
      <div class="what">${esc(i.message)}</div>
      ${i.bypass && i.ack.text ? `<div class="small bypass-why">✔ Bypassed by ${esc(i.ack.by)}: ${esc(i.ack.text)}</div>` : ""}
      <dl class="vals">
        <dt>Instr.</dt><dd class="mono">${esc(i.instruction)}</dd>
        <dt>Container</dt><dd class="mono">${esc(i.container)}</dd>
        ${i.field === "progress" ? `<dt>Now</dt><dd>${esc(i.nexl)}</dd>` : i.field !== "container" && i.field !== "instruction" ? `<dt>${esc(FIELD[i.field] || i.field)}</dt><dd>sheet <b>${esc(i.sheet || "—")}</b> · Nexl <b>${esc(i.nexl || "—")}</b></dd>` : ""}
      </dl>
      <div class="card-actions">
        <button class="ghost xs" data-act="go">Go to row</button>
        ${fixable(i) ? `<button class="ghost xs" data-act="sealphoto">📷 Check photo & fix</button>` : ""}
        ${opts.noSnooze && !i.ack ? "" : ackButtons(i)}
        <button class="ghost xs" data-act="nexl" data-word="${esc(word)}" data-filter="${filter}">Nexl ↗</button>
      </div></div>`;
  }
  function wireCards(container, list) {
    container.querySelectorAll(".card").forEach((c) => {
      const i = list[+c.dataset.k];
      c.addEventListener("click", (e) => {
        const act = e.target.dataset && e.target.dataset.act;
        if (act === "snooze") { e.stopPropagation(); snooze(i); return; }
        if (act === "bypass") { e.stopPropagation(); bypass([i]); return; }
        if (act === "sealphoto") { e.stopPropagation(); sealIssuePhoto(i); return; }
        if (act === "unsnooze") { e.stopPropagation(); snooze(i, 0); return; }
        if (act === "nexl") { e.stopPropagation(); openNexl(e.target.dataset.word, e.target.dataset.filter, i.instruction); return; }
        if (i.ref) { ExcelIO.goTo(i.ref).catch(() => {}); showDetail(i.ref.tab, i.ref.row); }
      });
    });
  }

  function renderCutoffs() {
    const box = $("cutoffBox");
    // Past cutoffs: keep the card 12 h if something missed it, but only 2 h when everything made it.
    const list = (state.result.cutoffs || []).filter((c) => c.hoursLeft <= Math.max(settings.cutoffH, 24) && c.hoursLeft > (c.open.length ? -12 : -2));
    box.hidden = !list.length;
    box.innerHTML = list.map((c) => {
      const cls = c.open.length === 0 ? "ok" : c.hoursLeft < 0 || c.atRisk ? "err" : c.hoursLeft <= 3 ? "err" : c.hoursLeft <= settings.cutoffH ? "warn" : "";
      const when = c.at.toLocaleString([], { weekday: "short", hour: "2-digit", minute: "2-digit" });
      const lab = NexlMatcher.cutoffLabel(c), past = lab.past, n = c.open.length, left = lab.big;
      return `<div class="cut ${cls}"><div><b>⚓ ${esc(c.vessel)}</b> <span class="muted small">${esc(c.terminal)} · ${esc(c.kind)} · ${esc(when)}</span></div>
        <div class="cutr"><span class="big">${esc(left)}</span><span class="small">${c.total - c.open.length}/${c.total} at port</span>
        ${!n ? `<span class="small okc">all made it${past ? "" : ` · closes in ${esc(NexlMatcher.fmtMin(c.hoursLeft * 60))}`}</span>` : c.atRisk && !past ? `<span class="small risk">🚨 ${c.atRisk} likely to miss</span>` : !past ? `<span class="small okc">✓ on track</span>` : ""}</div></div>`;
    }).join("");
  }

  function renderNow() {
    const res = state.result;
    renderCutoffs();
    const need = res.issues.filter((i) => !i.ack && (i.severity === "error" || i.severity === "warn")).sort((a, b) => SEV_ORDER(a) - SEV_ORDER(b));
    $("nNeed").textContent = need.length ? `(${need.length})` : "";
    const nl = $("needList");
    nl.innerHTML = need.length ? need.slice(0, 150).map((i, k) => issueCard(i, k, { compact: true })).join("") : `<p class="empty ok">✓ Nothing needs attention right now.</p>`;
    wireCards(nl, need);

    renderFills();
    renderFixes();

    // On the road
    const road = Object.values(res.rows).filter((d) => d.leg.stage === "moving" || d.leg.stage === "allocated")
      .sort((a, b) => (a.leg.stage === b.leg.stage ? 0 : a.leg.stage === "moving" ? -1 : 1));
    $("nRoad").textContent = road.length ? `(${road.length})` : "";
    $("roadList").innerHTML = road.length ? road.map((d, k) => `<div class="rrow ${d.leg.alerts.length ? "warn" : ""}" data-k="${k}">
        <span class="mono">${esc(d.container || d.id)}</span><span class="muted small">${esc(d.tab)} r${d.row}</span>
        <span class="rstep">${esc(d.leg.short)}${d.leg.alerts.length ? " · " + esc(d.leg.alerts.map((a) => a.short).join(" ")) : ""}</span></div>`).join("")
      : `<p class="empty">No trucks on the road for the checked instructions.</p>`;
    $("roadList").querySelectorAll(".rrow").forEach((r) => r.onclick = () => { const d = road[+r.dataset.k]; ExcelIO.goTo(d.ref).catch(() => {}); showDetail(d.tab, d.row); });

    // Snoozed
    const sn = res.issues.filter((i) => i.ack);
    $("snoozeBox").hidden = !sn.length;
    $("nSnooze").textContent = `(${sn.length})`;
    $("snoozeList").innerHTML = sn.map((i, k) => issueCard(i, k, { compact: true })).join("");
    wireCards($("snoozeList"), sn);
  }

  function renderIssues() {
    const res = state.result;
    const el = $("issueList");
    if (!res) return;
    const sev = $("fSeverity").value, tab = $("fTab").value;
    const list = res.issues.filter((i) =>
      (sev === "all" || (sev === "problems" ? i.severity !== "info" : i.severity === sev)) && (!tab || (i.ref && i.ref.tab === tab)));
    if (!list.length) { el.innerHTML = `<p class="empty">${res.issues.length ? "Nothing for this filter." : "✓ Everything on the checked tabs matches Nexl."}</p>`; return; }
    el.innerHTML = list.slice(0, 400).map((i, k) => issueCard(i, k, { noSnooze: i.severity === "info" })).join("")
      + (list.length > 400 ? `<p class="empty">…and ${list.length - 400} more (see the ${esc(CFG.checkTabName)} tab).</p>` : "");
    wireCards(el, list);
  }

  function podBadge(s) {
    const v = (s || "").toLowerCase();
    const cls = /verified|accepted|complete/.test(v) ? "ok" : /progress|pending|outstanding|awaiting/.test(v) ? "warn" : /reject|query|fail/.test(v) ? "err" : "";
    return s ? `<span class="badge ${cls}">POD ${esc(s)}</span>` : `<span class="badge">No POD</span>`;
  }

  function renderLive() {
    const res = state.result;
    const el = $("liveList");
    if (!res) return;
    const q = $("liveSearch").value.trim().toUpperCase();
    const showDone = $("showDone").checked;
    const refs = [];
    const html = res.groups.map((g) => {
      const instr = g.nexl[0] || {};
      const rows = g.containers.filter((c) => {
        const done = c.leg ? c.leg.stage === "completed" : c.nexl && /complete/i.test(c.nexl.moveStatus);
        if (!showDone && !q && done && !c.issues.some((i) => i.severity !== "info")) return false;
        if (!q) return true;
        const n = c.nexl || {};
        return [c.container, c.id, n.driver, n.owner, g.base, instr.customer, c.tracking && c.tracking.driver].join(" ").toUpperCase().includes(q);
      });
      if (!rows.length) return "";
      const st = [...new Set(g.nexl.map((i) => i.state))].join(" / ");
      const nIssues = g.containers.reduce((a, c) => a + c.issues.filter((i) => i.severity !== "info" && !i.ack).length, 0);
      return `<details class="group"${q || nIssues ? " open" : ""}>
        <summary><div class="t"><span>${esc(g.base)} · ${esc(instr.customer || "")}</span><span class="badge ${st === "Completed" ? "ok" : "info"}">${esc(st || "Nexl")}</span></div>
          <div class="s">${esc([instr.vessel && instr.vessel !== "0" ? instr.vessel : "", instr.type].filter(Boolean).join(" · "))}</div>
          <div class="s">${rows.length} shown of ${g.containers.length}${g.openSlots ? ` · ${g.openSlots} open slot(s)` : ""}${nIssues ? ` · <b style="color:var(--err)">${nIssues} issue(s)</b>` : ""}</div></summary>
        <div class="s route">${esc(instr.route || "")}${instr.completion ? ` · Done / in progress / not started / remaining: ${esc(instr.completion)}` : ""}</div>
        ${rows.map((c) => {
          const n = c.nexl || {};
          const k = refs.push(c.sheet || g.firstRef) - 1;
          const bad = c.issues.filter((i) => i.severity !== "info" && !i.ack).length;
          return `<div class="crow" data-k="${k}">
            <span class="mono">${esc(c.container)}${c.id !== g.base ? ` <span class="muted">(${esc(c.id)})</span>` : ""}</span>
            <span>${c.pending ? '<span class="badge info">Allocated</span>' : c.nexl ? podBadge(n.podStatus) : '<span class="badge err">Not in Nexl</span>'}${bad ? ` <span class="badge err">${bad}</span>` : ""}${!c.sheet ? ' <span class="badge err">Not on sheet</span>' : ""}</span>
            ${c.leg && c.leg.short ? `<span class="sub step">${esc(c.leg.short)}</span>` : c.step ? `<span class="sub step">${esc(c.step)}</span>` : ""}
          </div>`;
        }).join("")}
      </details>`;
    }).join("");
    el.innerHTML = html || `<p class="empty">${q ? "No match." : showDone ? "No Nexl instructions on the checked tabs." : "No active jobs. Tick 'Show completed jobs' to see finished ones."}</p>`;
    el.querySelectorAll(".crow").forEach((r) => r.addEventListener("click", () => { const ref = refs[+r.dataset.k]; if (ref) { ExcelIO.goTo(ref).catch(() => {}); showDetail(ref.tab, ref.row); } }));

    const nop = res.notOnPlan;
    $("notOnPlan").hidden = !nop.length;
    $("nopCount").textContent = `(${nop.length})`;
    $("nopList").innerHTML = nop.map((i) => `<div><b class="mono">${esc(i.id)}</b> ${esc(i.customer)} · ${esc(i.type)} · ${esc(i.state)}${i.vessel && i.vessel !== "0" ? " · " + esc(i.vessel) : ""}</div>`).join("");
  }

  // ---------- WhatsApp client updates (copy only) ----------
  async function copyText(text, what) {
    try { await navigator.clipboard.writeText(text); }
    catch (e) {
      const ta = document.createElement("textarea"); ta.value = text; document.body.appendChild(ta); ta.select();
      document.execCommand("copy"); ta.remove();
    }
    if (what) toast(`${what} copied. Paste it into WhatsApp.`);
  }
  function waClients() {
    const grid = state.waGrid;
    if (!grid || !grid.length) return [];
    const cols = NexlMatcher.detectColumns(grid[CFG.headerRow - 1] || [], CFG.fields);
    const loads = NexlWhatsApp.buildLoads(grid, cols, state.result ? state.result.rows : {}, CFG.whatsAppTab);
    state.waAll = loads;
    const pick = waDatePick(loads);
    return NexlWhatsApp.byClient(pick === "*" ? loads : loads.filter((l) => l.date === pick));
  }
  const localToday = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
  // Which load date to build the update for: the user's pick, else today, else every date.
  function waDatePick(loads) {
    const today = localToday();
    const dates = new Set(loads.map((l) => l.date));
    if (state.waDate != null && (state.waDate === "*" || dates.has(state.waDate))) return state.waDate;
    return dates.has(today) ? today : "*";
  }
  function renderWaDates() {
    const loads = state.waAll || [], sel = $("waDate");
    const count = new Map();
    for (const l of loads) count.set(l.date, (count.get(l.date) || 0) + 1);
    const today = localToday();
    const keys = [...count.keys()].sort((a, b) => (!a ? 1 : !b ? -1 : a.localeCompare(b)));
    const pick = waDatePick(loads);
    sel.innerHTML = `<option value="*">All dates (${loads.length})</option>` + keys.map((k) =>
      `<option value="${esc(k)}" ${k === pick ? "selected" : ""}>${esc(NexlWhatsApp.dateLabel(k))}${k === today ? " · today" : ""} (${count.get(k)})</option>`).join("");
    if (pick === "*") sel.value = "*";
    sel.parentElement.hidden = !loads.length;
  }
  const waOpts = () => ({ emoji: settings.waEmoji, fields: settings.waFields });

  // ---------- Driver messages (WhatsApp > Driver) ----------
  // Container / Booking on the Collect line: on by default for imports and full exports, off for PE CITRUS.
  function drvFields(sheet) {
    const saved = (settings.waDrvFieldsBy || {})[sheet];
    if (saved) return saved;
    const on = sheet !== CFG.whatsAppTab;
    return { container: on, booking: on };
  }
  function driverSource() {
    const sheets = CFG.driverSheets || [{ name: CFG.whatsAppTab, fallback: {} }];
    const pick = sheets.find((x) => x.name === settings.waDrvSheet) || sheets[0];
    if (pick.name === CFG.whatsAppTab) return { pick, loads: state.waAll || [], dated: true };
    // Other sheets: the rows the last sync read (rows of instructions that are live in Nexl).
    const t = state.input && state.input.usable.find((x) => x.name === pick.name);
    const jobs = t ? NexlWhatsApp.driverJobs(t.name, t.rows, t.cols, state.result ? state.result.rows : {}, pick.fallback) : [];
    return { pick, loads: jobs, dated: false };
  }
  function renderDriver() {
    const src = driverSource();
    const sel0 = $("waDrvSheet");
    sel0.innerHTML = (CFG.driverSheets || []).map((x) => `<option value="${esc(x.name)}">${esc(x.label || x.name)}</option>`).join("");
    sel0.value = src.pick.name;
    $("waDate").parentElement.style.display = src.dated ? "" : "none";
    const fl = drvFields(src.pick.name);
    $("waDrvCont").checked = !!fl.container; $("waDrvBook").checked = !!fl.booking;
    const day = src.dated ? waDatePick(src.loads) : "*";
    const loads = src.loads.filter((l) => day === "*" || l.date === day);
    const q = ($("waDrvSearch").value || "").trim().toUpperCase();
    const list = loads.filter((l) => !q || [l.loadRef, l.driver, l.instruction, l.client, l.container].join(" ").toUpperCase().includes(q))
      .sort((a, b) => (a.driver || "~").localeCompare(b.driver || "~") || String(a.loadRef).localeCompare(String(b.loadRef)));
    const short = (x) => String(x || "?").replace(/\s*\((POL|VIA|POD)\)\s*/g, "");
    $("waDrvList").innerHTML = list.length ? list.map((l) => `<button type="button" class="drv ${state.waDrvSel === l.id ? "sel" : ""}" data-id="${esc(l.id)}">
        <b class="mono">${esc(l.loadRef)}</b> <span class="muted small">${esc(l.instruction || "")}</span> · <b>${esc(l.driver || "no driver")}</b>${l.portBooking ? ` <span class="small">· 🎫 ${esc(l.portBooking)}</span>` : ""}
        <span class="small drv-route">${esc(short(l.route.collect))} → ${esc((l.route.packing || []).map(short).join(" → ") || "?")} → ${esc(short(l.route.dropoff))}</span></button>`).join("")
      : `<p class="empty">No loads for this date.</p>`;
    $("waDrvList").querySelectorAll(".drv").forEach((b) => (b.onclick = () => { state.waDrvSel = b.dataset.id; renderDriver(); showDrvBox(); }));
    const sel = src.loads.find((l) => l.id === state.waDrvSel);
    $("waDrvBox").hidden = !sel; $("waDrvHint").hidden = !!sel;
    if (sel) {
      $("waDrvTitle").textContent = `${sel.loadRef} · ${sel.driver || "no driver on the sheet"}`;
      $("waDrvSrc").textContent = sel.route.source === "nexl" ? `route from Nexl ${sel.instruction}` : "route from the sheet (not found in Nexl)";
      const key = [sel.id, settings.waEmoji, fl.container, fl.booking].join("|");
      if (state.waDrvMsgFor !== key) { $("waDrvMsg").value = NexlWhatsApp.driverMessage(sel, { emoji: settings.waEmoji, fields: fl }); state.waDrvMsgFor = key; }
    }
  }
  // Bring the driver message (at the top of the tab) into view and flash it.
  function showDrvBox() {
    const b = $("waDrvBox");
    if (b.hidden) return;
    b.scrollIntoView({ block: "start", behavior: "smooth" });
    b.classList.remove("flash"); void b.offsetWidth; b.classList.add("flash");
  }
  // A row clicked on the sheet while WhatsApp > Driver is open: show that load's driver message.
  function driverFromSheet(tab, row) {
    if (state.view !== "wa" || state.waMode !== "driver") return false;
    const sheet = (CFG.driverSheets || []).find((x) => x.name === tab);
    if (!sheet) return false;
    if (settings.waDrvSheet !== tab) { settings.waDrvSheet = tab; saveSettings(); }
    const id = `${tab}|${row}`;
    const src = driverSource();
    if (!src.loads.some((l) => l.id === id)) { toast(`Row ${row} on ${tab} isn't a load with a Nexl instruction yet.`); return true; }
    state.waDrvSel = id;
    $("waDrvSearch").value = "";
    if (src.dated) { const l = src.loads.find((x) => x.id === id); if (l && state.waDate !== "*") state.waDate = l.date; } // show that load's date in the list
    renderWhatsApp(); showDrvBox();
    return true;
  }
  function setWaMode(m) {
    state.waMode = m;
    document.querySelectorAll(".wa-mode button").forEach((b) => b.classList.toggle("active", b.dataset.mode === m));
    const drv = m === "driver";
    $("waDriver").hidden = !drv; $("waList").hidden = drv;
    $("waClient").parentElement.style.display = drv ? "none" : "";
    if (!drv) $("waDate").parentElement.style.display = "";
    document.querySelectorAll(".wa-opts [data-f]").forEach((cb) => (cb.parentElement.style.display = drv ? "none" : ""));
    const lbl = document.querySelector("#view-wa > .wa-opts > span.muted"); if (lbl) lbl.style.display = drv ? "none" : "";
    $("waHelp").hidden = drv;
    renderWhatsApp();
  }

  function renderWhatsApp() {
    const el = $("waList");
    $("waTab").textContent = CFG.whatsAppTab;
    $("waEmoji").checked = settings.waEmoji;
    document.querySelectorAll(".wa-opts [data-f]").forEach((cb) => (cb.checked = !!settings.waFields[cb.dataset.f]));
    const allClients = waClients();
    renderWaDates();
    if (state.waMode === "driver") { renderDriver(); return; }
    // Client drop-down: one client at a time keeps it clean (or "All clients").
    if (allClients.length && state.waClient !== "*" && !allClients.some((c) => c.key === state.waClient)) state.waClient = allClients[0].key;
    const cs = $("waClient");
    cs.innerHTML = `<option value="*">All clients (${allClients.length})</option>` + allClients.map((c) =>
      `<option value="${esc(c.key)}">${esc(c.client)} — ${c.loads.length} load${c.loads.length === 1 ? "" : "s"}</option>`).join("");
    cs.value = state.waClient || "*";
    cs.parentElement.hidden = !allClients.length;
    const clients = state.waClient === "*" ? allClients : allClients.filter((c) => c.key === state.waClient);
    if (!clients.length) { el.innerHTML = `<p class="empty">No loads found on the ${esc(CFG.whatsAppTab)} tab${state.waAll && state.waAll.length ? " for this date" : ""}.</p>`; return; }
    el.innerHTML = clients.map((c, k) => {
      const picked = c.loads.filter((l) => !state.waUnticked.has(l.id));
      const msg = state.waEdits[c.key] != null ? state.waEdits[c.key] : picked.length ? NexlWhatsApp.formatMessage(picked, waOpts()) : "";
      const vessels = [...new Set(c.loads.map((l) => l.vessel || "VESSEL TBC"))];
      return `<div class="wa" data-k="${k}">
        <div class="wa-head"><b>${esc(c.client)}</b><span class="muted small">${picked.length}/${c.loads.length} load(s) selected</span></div>
        ${vessels.map((v) => `<div class="wa-vessel">${settings.waEmoji ? "🚢 " : ""}${esc(v)}</div>` + c.loads.filter((l) => (l.vessel || "VESSEL TBC") === v).map((l) => `
          <div class="wa-load${l.live ? " live" : ""}">
            <input type="checkbox" data-id="${esc(l.id)}" ${state.waUnticked.has(l.id) ? "" : "checked"} aria-label="Include ${esc(l.loadRef)}">
            <span class="wa-ref mono">${esc(l.loadRef)}</span>
            <span class="wa-st">${l.hasComment ? esc(l.status) : `<span class="wa-none">no comment yet</span>`}${l.nexlHint ? `<br><span class="wa-rem" title="Nexl status (not included in the update)">Nexl: ${esc(l.nexlHint)}</span>` : ""}</span>
            <button class="ghost xs wa-one" data-id="${esc(l.id)}" type="button" title="Copy an update for just this load">Copy</button>
          </div>`).join("")).join("")}
        <textarea class="full wa-msg" rows="${Math.min(12, msg.split("\n").length + 1)}" placeholder="Tick at least one load">${esc(msg)}</textarea>
        <div class="row-actions">
          <button class="wa-copy primary small-btn" type="button" ${picked.length ? "" : "disabled"}>Copy update (${picked.length} load${picked.length === 1 ? "" : "s"})</button>
          <button class="wa-all ghost xs" type="button">${picked.length === c.loads.length ? "Untick all" : "Tick all"}</button>
          ${state.waEdits[c.key] != null ? '<button class="wa-reset ghost xs" type="button" title="Discard your edits and rebuild from live status">↺ Rebuild</button>' : ""}
        </div></div>`;
    }).join("");
    el.querySelectorAll(".wa").forEach((box) => {
      const c = clients[+box.dataset.k];
      const ta = box.querySelector(".wa-msg");
      ta.oninput = () => (state.waEdits[c.key] = ta.value);
      box.querySelectorAll(".wa-load input").forEach((cb) => (cb.onchange = () => {
        cb.checked ? state.waUnticked.delete(cb.dataset.id) : state.waUnticked.add(cb.dataset.id);
        delete state.waEdits[c.key];
        renderWhatsApp();
      }));
      box.querySelectorAll(".wa-one").forEach((b) => (b.onclick = () => {
        const l = c.loads.find((x) => x.id === b.dataset.id);
        copyText(NexlWhatsApp.formatMessage([l], waOpts()), `Update for ${l.loadRef}`);
      }));
      box.querySelector(".wa-copy").onclick = () => copyText(ta.value, `${c.client} update`);
      box.querySelector(".wa-all").onclick = () => {
        const all = c.loads.every((l) => !state.waUnticked.has(l.id));
        c.loads.forEach((l) => (all ? state.waUnticked.add(l.id) : state.waUnticked.delete(l.id)));
        delete state.waEdits[c.key];
        renderWhatsApp();
      };
      const reset = box.querySelector(".wa-reset");
      if (reset) reset.onclick = () => { delete state.waEdits[c.key]; renderWhatsApp(); };
    });
  }
  function initWhatsAppOptions() {
    $("waEmoji").addEventListener("change", () => { settings.waEmoji = $("waEmoji").checked; saveSettings(); state.waEdits = {}; renderWhatsApp(); });
    document.querySelectorAll(".wa-opts [data-f]").forEach((cb) => cb.addEventListener("change", () => {
      settings.waFields = { ...settings.waFields, [cb.dataset.f]: cb.checked }; saveSettings(); state.waEdits = {}; renderWhatsApp();
    }));
  }

  // ---------- selected row details ----------
  function showDetail(tab, row, quiet) {
    const box = $("detail");
    const d = state.result && state.result.rows[tab + "|" + row];
    state.selected = d ? { tab, row } : null;
    if (!d) { box.hidden = true; return; }
    const n = d.nexl || {}, t = d.tracking, ins = d.instr || {};
    const stops = d.leg.stops || [];
    const stepper = stops.length ? `<ol class="stepper">${stops.map((s, i) => {
      const cur = !s.done && (i === 0 || stops[i - 1].done);
      return `<li class="${s.done ? "done" : cur ? "cur" : ""}"><span class="dot"></span><span class="nm">${esc(s.name)}</span><span class="tm">${s.time ? esc(s.time) : cur ? "next" : ""}</span></li>`;
    }).join("")}</ol>` : "";
    const issues = (d.issues || []).filter((i) => i.severity !== "info");
    const word = d.container || d.id;
    box.hidden = false;
    box.innerHTML = `<div class="dhead"><b>${esc(tab)} · row ${row}</b><button class="ghost xs" id="dClose" aria-label="Close">✕</button></div>
      <div class="dtitle"><span class="mono">${esc(d.container || "(no container yet)")}</span> · ${esc(d.id)} ${ins.customer ? "· " + esc(ins.customer) : ""}</div>
      <div class="muted small">${esc([ins.vessel && ins.vessel !== "0" ? ins.vessel : "", ins.type, ins.state].filter(Boolean).join(" · "))}</div>
      <div class="dstep">${esc(d.leg.step || "No live tracking for this row.")}</div>
      ${stepper}
      <dl class="vals">
        ${t || n.driver ? `<dt>Driver</dt><dd>${esc((t && t.driver) || n.driver)}</dd>` : ""}
        ${n.owner || (t && t.owner) ? `<dt>Transporter</dt><dd>${esc(n.owner || t.owner)}</dd>` : ""}
        ${n.seal ? `<dt>Seal</dt><dd class="mono">${esc(n.seal)}</dd>` : ""}
        ${n.moveStatus ? `<dt>Move</dt><dd>${esc(n.moveStatus)}${n.start ? ` · ${esc(n.start)} → ${esc(n.end || "…")}` : ""}</dd>` : ""}
        ${n.podStatus ? `<dt>POD</dt><dd>${podBadge(n.podStatus)}</dd>` : ""}
        ${t && t.lastPing ? `<dt>Last ping</dt><dd>${esc(t.lastPing)} ago</dd>` : ""}
        ${d.cutoff && d.cutoff.eta ? `<dt>Port ETA</dt><dd class="${d.cutoff.late ? "late" : ""}">~${esc(fmtTime(d.cutoff.eta))} · cutoff ${esc(fmtTime(d.cutoff.at))} ${d.cutoff.late ? "🚨 likely to miss" : "✓"}</dd>` : ""}
        ${d.slip ? `<dt>Port slip</dt><dd>📄 uploaded <button class="link xs" id="dSlip" type="button">view</button></dd>` : ""}
      </dl>
      ${issues.length ? `<div class="dissues">${issues.map((i) => `<div class="di ${i.ack ? "acked" : i.severity}">${i.ack ? ackTag(i) + ": " : ""}${esc(i.message)}${i.bypass && i.ack.text ? ` <span class="muted">(${esc(i.ack.text)})</span>` : ""}</div>`).join("")}</div>` : ""}
      ${d.fills.length ? `<div class="dfills">${d.fills.map((f, k) => f.needsPhoto
        ? `<div class="dphoto">${photoOf(f).photo ? `<img class="thumb" data-ph="${k}" src="${esc(photoOf(f).photo)}" alt="Container photo" title="Click to enlarge">` : ""}
           <div>${f.photoKind === "seal" ? "Nexl seal" : "App container"}: <span class="mono">${esc(f.value)}</span><br>${photoBadge(f)}<br>
           ${fillAllowed(f) ? `<button class="primary xs" data-fill="${k}" type="button">Apply to ${esc(f.col + f.row)}</button>` : f.photoKind === "seal" && photoOf(f).photo ? `<button class="primary xs" data-ph="${k}" type="button">Check seal photo</button>` : `<span class="muted small">Apply unlocks once the photo matches.</span>`}
           ${photoOf(f).photo ? ` <button class="link xs" data-ph="${k}" type="button">View photo</button>` : ""}</div></div>`
        : f.slipPath ? `<div class="dphoto"><div>📄 Port slip uploaded → <b>${esc(f.value)}</b><br><button class="primary xs" data-fill="${k}" type="button">Put ${esc(f.value)} in ${esc(f.col + f.row)}</button> <button class="link xs" data-slipd="${k}" type="button">View slip</button></div></div>`
        : `<button class="ghost xs" data-fill="${k}">Fill ${esc(FIELD[f.field] || f.field)} ← ${esc(f.value)}</button>`).join("")}</div>` : ""}
      <div class="row-actions">
        <button class="primary small-btn" id="dNexl" type="button">Open in Nexl ↗</button>
        ${issues.some(fixable) ? `<button class="ghost small-btn" id="dSeal" type="button">📷 Check photo & fix</button>` : ""}
        ${issues.some((i) => !i.ack) ? `<button class="ghost small-btn" id="dSnooze" type="button">👀 I'm on it</button><button class="ghost small-btn" id="dBypass" type="button">✔ Bypass</button>` : ""}
      </div>`;
    $("dClose").onclick = () => { box.hidden = true; state.selected = null; };
    $("dNexl").onclick = () => openNexl(word, d.container ? "container" : "instruction", d.id);
    if ($("dSnooze")) $("dSnooze").onclick = async () => { for (const i of issues.filter((x) => !x.ack)) await snooze(i); };
    if ($("dSlip")) $("dSlip").onclick = () => showSlip({ slipPath: d.slip.path, container: d.container, instruction: d.id, value: "", tab, row, col: "" , slipKind: /IMPORT/i.test(tab) ? "collected" : "stacked" });
    if ($("dSeal")) $("dSeal").onclick = () => fixWithPhoto(issues.find(fixable));
    if ($("dBypass")) $("dBypass").onclick = () => bypass(issues.filter((x) => !x.ack));
    box.querySelectorAll("[data-fill]").forEach((b) => (b.onclick = () => doFill([d.fills[+b.dataset.fill]])));
    box.querySelectorAll("[data-ph]").forEach((b) => (b.onclick = () => showPhoto(d.fills[+b.dataset.ph])));
    box.querySelectorAll("[data-slipd]").forEach((b) => (b.onclick = () => showSlip(d.fills[+b.dataset.slipd])));
    if (!quiet) box.scrollIntoView({ block: "start", behavior: "smooth" });
  }

  // ---------- wiring ----------
  function showView(v) {
    state.view = v;
    if (v === "wa" && !state.waGrid) { // load the WhatsApp data only when someone actually opens the tab
      ExcelIO.readTabValues(CFG.whatsAppTab).then((g) => { state.waGrid = g; if (state.view === "wa") renderWhatsApp(); }).catch(() => {});
    }
    document.querySelectorAll(".tabs button").forEach((b) => b.classList.toggle("active", b.dataset.view === v));
    for (const id of ["now", "issues", "live", "wa", "settings"]) $("view-" + id).hidden = id !== v;
    $("detail").classList.toggle("off", v === "wa" || v === "settings");
  }

  function initSettingsUI() {
    const bind = (id, key, kind = "checked", parse = (x) => x) => {
      $(id)[kind] = settings[key];
      $(id).addEventListener("change", () => { settings[key] = parse($(id)[kind]); saveSettings(); });
    };
    bind("sName", "name", "value", (v) => v.trim());
    bind("sSplash", "splash");
    bind("sPopUp", "popUp");
    bind("sAuto", "auto");
    bind("sMinutes", "minutes", "value", (v) => (SLOT_CHOICES.includes(+v) ? +v : 5));
    $("sMinutes").addEventListener("change", () => { state.nextSlot = nextSlot(); showNextSync(); });
    $("sAuto").addEventListener("change", () => { state.nextSlot = nextSlot(); showNextSync(); });
    bind("sSnoozeMin", "snoozeMin", "value", (v) => Math.max(10, Math.min(480, +v || 60)));
    bind("sStatusCols", "statusCols");
    bind("sNotes", "notes");
    bind("sWriteTab", "writeTab");
    bind("sWriteInfo", "writeInfo");
    bind("sNotStarted", "notStarted", "value", (v) => Math.max(5, +v || CFG.notStartedMinutes));
    bind("sStuck", "stuck", "value", (v) => Math.max(15, +v || CFG.stuckMinutes));
    bind("sRegion", "region", "value");
    bind("sCutoffH", "cutoffH", "value", (v) => Math.max(2, Math.min(72, +v || CFG.cutoffWarnHours)));
    $("sTabName").textContent = CFG.checkTabName;
    $("sAutoOpen").checked = ExcelIO.getAutoOpen();
    $("sAutoOpen").addEventListener("change", async () => {
      const ok = await ExcelIO.setAutoOpen($("sAutoOpen").checked);
      toast(ok ? ($("sAutoOpen").checked ? "Nexl Check will open with this workbook." : "Auto-open switched off.") : "Couldn't change auto-open.");
    });
    $("sTabs").innerHTML = CFG.tabs.map((t, k) =>
      `<label><span>${esc(t.name)}</span><input type="checkbox" data-k="${k}" ${settings.disabledTabs.includes(t.name) ? "" : "checked"}></label>`).join("");
    $("sTabs").addEventListener("change", (e) => {
      const t = CFG.tabs[+e.target.dataset.k];
      settings.disabledTabs = settings.disabledTabs.filter((n) => n !== t.name);
      if (!e.target.checked) settings.disabledTabs.push(t.name);
      saveSettings();
    });
  }

  function initSplash() {
    const s = $("splash");
    if (!settings.splash) { s.hidden = true; splashGone = true; return; }
    s.addEventListener("click", () => hideSplash(true));
    setTimeout(() => hideSplash(true), 15000); // never block the panel for long
  }

  Office.onReady(async () => {
    initSplash();
    initSettingsUI();
    initWhatsAppOptions();
    $("syncBtn").addEventListener("click", sync);
    $("modalClose").addEventListener("click", closeModal);
    $("modal").addEventListener("click", (e) => { if (e.target.id === "modal") closeModal(); });
    document.querySelectorAll(".tabs button").forEach((b) => b.addEventListener("click", () => showView(b.dataset.view)));
    $("fSeverity").addEventListener("change", renderIssues);
    $("fTab").addEventListener("change", renderIssues);
    $("liveSearch").addEventListener("input", renderLive);
    $("showDone").addEventListener("change", renderLive);
    $("fillAll").addEventListener("click", () => { (state.result ? state.result.fills : []).forEach((f, k) => fillAllowed(f) && state.fillSel.add(k)); renderFills(); });
    $("fillSel").addEventListener("click", () => doFill([...state.fillSel].map((k) => state.result.fills[k]).filter(Boolean)));
    $("hideBtn").addEventListener("click", hidePanel);
    $("waDate").addEventListener("change", () => { state.waDate = $("waDate").value; state.waEdits = {}; renderWhatsApp(); });
    document.querySelectorAll(".wa-mode button").forEach((b) => b.addEventListener("click", () => setWaMode(b.dataset.mode)));
    $("waDrvSearch").addEventListener("input", renderDriver);
    $("waDrvSheet").addEventListener("change", () => { settings.waDrvSheet = $("waDrvSheet").value; saveSettings(); state.waDrvSel = null; renderDriver(); });
    const setDrvField = (f, v) => { const n = driverSource().pick.name; settings.waDrvFieldsBy = { ...settings.waDrvFieldsBy, [n]: { ...drvFields(n), [f]: v } }; saveSettings(); renderDriver(); };
    $("waDrvCont").addEventListener("change", () => setDrvField("container", $("waDrvCont").checked));
    $("waDrvBook").addEventListener("change", () => setDrvField("booking", $("waDrvBook").checked));
    $("waDrvCopy").addEventListener("click", () => copyText($("waDrvMsg").value, "Driver message"));
    $("waClient").addEventListener("change", () => { state.waClient = $("waClient").value; renderWhatsApp(); });
    if (canHide() && Office.addin.onVisibilityModeChanged) {
      Office.addin.onVisibilityModeChanged((a) => {
        state.hidden = a.visibilityMode !== "Taskpane";
        if (state.hidden) state.knownErrors = openErrorKeys();
      }).catch(() => {});
    }
    document.querySelectorAll(".tile").forEach((t) => t.addEventListener("click", () => {
      const g = t.dataset.go;
      if (g === "jobs") { showView("live"); return; }
      if (g === "error") { $("fSeverity").value = "error"; showView("issues"); renderIssues(); return; }
      showView("now");
      const target = g === "road" ? $("roadHead") : $("needList");
      target.scrollIntoView({ behavior: "smooth" });
    }));

    // Auto-open: switch it on once, the first time this user opens the panel (they can turn it off in Settings).
    if (!settings.autoOpenSet) {
      settings.autoOpenSet = true; saveSettings();
      if (!ExcelIO.getAutoOpen()) { await ExcelIO.setAutoOpen(true); $("sAutoOpen").checked = true; }
    }

    const p = await NexlClient.ping();
    pill("bridgePill", p.ok, p.ok ? "Bridge" : "Bridge missing");
    $("versions").textContent = `Add-in v${VERSION}` + (p.ok ? ` · Bridge v${p.version}` : "");
    ExcelIO.watchSelection([...new Set([...CFG.tabs.map((t) => t.name), ...(CFG.driverSheets || []).map((t) => t.name)])], (tab, row) => { if (!driverFromSheet(tab, row)) showDetail(tab, row, true); }).catch(() => {});
    sync();
    checkForUpdate();

    // Team-aligned auto-refresh: everyone syncs on the same clock slots (e.g. 10:00, 10:05, 10:10 ...),
    // so all panels and the sheet show the same picture at the same time. Keeps running while the panel is hidden.
    state.nextSlot = nextSlot();
    showNextSync();
    setInterval(() => {
      if (!settings.auto) return;
      // Readers don't all hit Excel in the same second: the writer goes on the mark, readers 10-90 s later.
      const delay = state.isWriter ? 0 : (state.readerOffset = state.readerOffset || 10000 + Math.floor(Math.random() * 80000));
      if (Date.now() < state.nextSlot + delay) return;
      // A hidden panel that isn't the writer only needs every 3rd slot (15 min).
      if (state.hidden && !state.isWriter && (state.skipHidden = ((state.skipHidden || 0) + 1) % 3) !== 1) { state.nextSlot = nextSlot(); showNextSync(); return; }
      if (state.busy) return; // a sync is running: catch the slot as soon as it finishes
      state.nextSlot = nextSlot();
      showNextSync();
      sync().finally(showNextSync);
    }, 5000);
  });
})();
