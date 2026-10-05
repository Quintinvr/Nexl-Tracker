/*
 * Nexl Check — task pane controller (v1.2).
 */
(function () {
  "use strict";
  const CFG = window.NEXL_CONFIG;
  const VERSION = "1.2.0";
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ---------- settings (per user, this browser) ----------
  const SKEY = "nexlcheck.settings.v1";
  const defaults = { auto: true, minutes: CFG.refreshMinutes, writeTab: true, writeInfo: false, region: CFG.region, disabledTabs: [],
    statusCols: true, notes: true, notStarted: CFG.notStartedMinutes, stuck: CFG.stuckMinutes, name: "", snoozeMin: 60, splash: true, autoOpenSet: false };
  let settings = { ...defaults };
  try { settings = { ...defaults, ...JSON.parse(localStorage.getItem(SKEY) || "{}") }; } catch (e) { /* storage unavailable */ }
  const saveSettings = () => { try { localStorage.setItem(SKEY, JSON.stringify(settings)); } catch (e) { /* ignore */ } };

  // ---------- state ----------
  const state = { busy: false, result: null, lastSyncAt: 0, firstSeen: new Map(), view: "now", input: null, acks: {}, selected: null, fillSel: new Set() };
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
    return { now: new Date(), notStartedMinutes: settings.notStarted, stuckMinutes: settings.stuck, acks: state.acks };
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
      const [sheetTabs, acks] = await Promise.all([ExcelIO.readPlanTabs({ ...CFG, tabs: tabsCfg }, bases), ExcelIO.readAcks().catch(() => ({}))]);
      state.acks = acks;
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
      render();
      hideSplash();

      if (settings.writeTab) {
        progress(0.96, `Updating ${CFG.checkTabName} tab…`);
        const list = res.issues.filter((i) => settings.writeInfo || i.severity !== "info");
        try {
          await ExcelIO.writeCheckTab(CFG, list, (i) => { const d = state.firstSeen.get(i.key); return d ? fmtStamp(d) : ""; }, fmtStamp(now));
        } catch (e) { tabWarnings.push(`Couldn't update the ${esc(CFG.checkTabName)} tab (${esc(e.message)}). The panel is still up to date.`); }
      }
      await writeColumns(tabWarnings);
      if (settings.notes) {
        progress(0.98, "Updating cell notes…");
        try {
          const n = await ExcelIO.syncNotes(CFG, res.issues, fmtTime(now));
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

  async function writeColumns(warnings) {
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
      modal("Who's on it?", `<p class="small">Your name is shown to the team next to snoozed alerts.</p>
        <input id="nameIn" type="text" maxlength="24" placeholder="Your name" class="full">
        <div class="row-actions"><button id="nameOk" class="primary small-btn" type="button">Save</button></div>`);
      const go = () => { const v = $("nameIn").value.trim(); if (!v) return; settings.name = v; $("sName").value = v; saveSettings(); closeModal(); resolve(v); };
      $("nameOk").onclick = go;
      $("nameIn").onkeydown = (e) => e.key === "Enter" && go();
      $("nameIn").focus();
    });
  }

  async function snooze(issue, minutes) {
    const by = await ensureName();
    const m = minutes == null ? settings.snoozeMin : minutes;
    try {
      await ExcelIO.writeAck(issue.key, by, m, issue.message);
      state.acks = await ExcelIO.readAcks();
      toast(m ? `Snoozed for ${m} min — the team sees "👀 ${by}"` : "Snooze removed");
      await recompute();
    } catch (e) { toast("Couldn't save the snooze: " + e.message); }
  }

  async function openNexl(word, filter) {
    const w = String(word || "").trim();
    if (!w) return;
    const r = await NexlClient.openInNexl(filter, w);
    if (r.ok) return;
    if (r.error === "NOT_LOGGED_IN") { toast("Log in to Nexl first, then try again."); return; }
    // Older bridge (1.0) or blocked: copy + open Nexl so the user can paste into Nexl's search.
    try { await navigator.clipboard.writeText(w); } catch (e) { /* ignore */ }
    window.open("https://controller.nexl.online/php/landing.php", "_blank", "noopener");
    toast(`Copied ${w} — paste it into Nexl's search. (Update the bridge extension for one-click search.)`);
  }

  async function doFill(list) {
    if (!list.length) return;
    try {
      const r = await ExcelIO.fillBlanks(list);
      toast(`Filled ${r.filled} cell(s)${r.skipped ? `, skipped ${r.skipped} that already had a value` : ""}.`);
      state.fillSel.clear();
      setTimeout(sync, 400);
    } catch (e) { toast("Couldn't fill: " + e.message); }
  }

  // ---------- handover ----------
  function buildHandover() {
    const res = state.result;
    if (!res) return "Sync first.";
    const now = new Date();
    const rows = Object.values(res.rows);
    const open = res.issues.filter((i) => !i.ack && i.severity !== "info");
    const errs = open.filter((i) => i.severity === "error");
    const ns = open.filter((i) => i.code === "notstarted");
    const stuck = open.filter((i) => i.code === "stuck");
    const checks = open.filter((i) => i.severity === "warn" && i.field !== "progress");
    const moving = rows.filter((d) => d.leg.stage === "moving");
    const delivered = rows.filter((d) => d.leg.stage === "delivered" || (d.leg.stage === "completed" && d.nexl && !/verified/i.test(d.nexl.podStatus || "")));
    const snoozed = res.issues.filter((i) => i.ack);
    const where = (i) => (i.ref ? ` (${i.ref.tab} r${i.ref.row})` : "");
    const L = [];
    L.push(`NEXL HANDOVER — ${(settings.region || "").toUpperCase()} — ${fmtStamp(now)}${settings.name ? " — " + settings.name : ""}`);
    L.push(`${moving.length} on the road · ${ns.length} not started · ${stuck.length} stuck · ${errs.length} errors · ${delivered.length} delivered awaiting POD`);
    const sec = (title, list, fn) => { if (!list.length) return; L.push(""); L.push(`${title} (${list.length})`); list.forEach((x) => L.push("• " + fn(x))); };
    sec("⛔ ERRORS ON THE PLAN", errs, (i) => `${i.instruction} ${i.container}: ${i.message}${where(i)}`);
    sec("⏰ ALLOCATED, NOT STARTED", ns, (i) => `${i.instruction} ${i.container}: ${i.message.replace(/^Not started: /, "")}`);
    sec("🛑 STUCK", stuck, (i) => `${i.instruction} ${i.container}: ${i.message.replace(/^Stuck: /, "")}`);
    sec("🚚 ON THE ROAD", moving, (d) => `${d.id} ${d.container || ""}: ${d.leg.step}`);
    sec("📍 DELIVERED — POD OUTSTANDING", delivered, (d) => `${d.id} ${d.container || ""}: ${d.leg.step}`);
    sec("⚠ CHECK ON THE PLAN", checks, (i) => `${i.instruction} ${i.container}: ${i.message}${i.nexl ? " (Nexl: " + String(i.nexl).split(" | ")[0] + ")" : ""}${where(i)}`);
    sec("👀 BEING HANDLED", snoozed, (i) => `${i.instruction} ${i.container}: ${i.message} — ${i.ack.by}`);
    return L.join("\n");
  }
  function showHandover() {
    const text = buildHandover();
    modal("Shift handover", `<textarea id="hoText" class="full" rows="16" readonly>${esc(text)}</textarea>
      <div class="row-actions"><button id="hoCopy" class="primary small-btn" type="button">Copy</button>
      <span class="muted small">Paste into Teams, WhatsApp or email.</span></div>`);
    $("hoCopy").onclick = async () => {
      const ta = $("hoText");
      try { await navigator.clipboard.writeText(ta.value); }
      catch (e) { ta.select(); document.execCommand("copy"); }
      toast("Handover copied.");
    };
  }

  // ---------- rendering ----------
  const FIELD = { container: "Container", instruction: "Instruction", seal: "Seal", booking: "Booking ref", loadRef: "Load ref",
    vessel: "Vessel", customer: "Customer", transporter: "Transporter", driver: "Driver", progress: "Now" };

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
    if (state.selected) showDetail(state.selected.tab, state.selected.row, true);
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
          ${i.ack ? `<button class="ghost xs" data-act="unsnooze">Un-snooze</button>` : `<button class="ghost xs" data-act="snooze">👀 I'm on it</button>`}
          <button class="ghost xs" data-act="nexl" data-word="${esc(word)}" data-filter="${filter}">Nexl ↗</button>
          <span class="muted small since">${i.ack ? "👀 " + esc(i.ack.by) : since ? fmtTime(since) : ""}</span>
        </div></div>`;
    }
    return `<div class="card ${i.ack ? "acked" : i.severity}" data-k="${k}">
      <div class="where"><span>${where}</span><span>${i.ack ? "👀 " + esc(i.ack.by) : since ? "since " + fmtTime(since) : ""}</span></div>
      <div class="what">${esc(i.message)}</div>
      <dl class="vals">
        <dt>Instr.</dt><dd class="mono">${esc(i.instruction)}</dd>
        <dt>Container</dt><dd class="mono">${esc(i.container)}</dd>
        ${i.field === "progress" ? `<dt>Now</dt><dd>${esc(i.nexl)}</dd>` : i.field !== "container" && i.field !== "instruction" ? `<dt>${esc(FIELD[i.field] || i.field)}</dt><dd>sheet <b>${esc(i.sheet || "—")}</b> · Nexl <b>${esc(i.nexl || "—")}</b></dd>` : ""}
      </dl>
      <div class="card-actions">
        <button class="ghost xs" data-act="go">Go to row</button>
        ${i.ack ? `<button class="ghost xs" data-act="unsnooze">Un-snooze</button>` : opts.noSnooze ? "" : `<button class="ghost xs" data-act="snooze">👀 I'm on it</button>`}
        <button class="ghost xs" data-act="nexl" data-word="${esc(word)}" data-filter="${filter}">Nexl ↗</button>
      </div></div>`;
  }
  function wireCards(container, list) {
    container.querySelectorAll(".card").forEach((c) => {
      const i = list[+c.dataset.k];
      c.addEventListener("click", (e) => {
        const act = e.target.dataset && e.target.dataset.act;
        if (act === "snooze") { e.stopPropagation(); snooze(i); return; }
        if (act === "unsnooze") { e.stopPropagation(); snooze(i, 0); return; }
        if (act === "nexl") { e.stopPropagation(); openNexl(e.target.dataset.word, e.target.dataset.filter); return; }
        if (i.ref) { ExcelIO.goTo(i.ref).catch(() => {}); showDetail(i.ref.tab, i.ref.row); }
      });
    });
  }

  function renderNow() {
    const res = state.result;
    const need = res.issues.filter((i) => !i.ack && (i.severity === "error" || i.severity === "warn")).sort((a, b) => SEV_ORDER(a) - SEV_ORDER(b));
    $("nNeed").textContent = need.length ? `(${need.length})` : "";
    const nl = $("needList");
    nl.innerHTML = need.length ? need.slice(0, 150).map((i, k) => issueCard(i, k, { compact: true })).join("") : `<p class="empty ok">✓ Nothing needs attention right now.</p>`;
    wireCards(nl, need);

    // Fill blanks
    const fills = res.fills;
    $("fillBox").hidden = !fills.length;
    $("nFill").textContent = fills.length;
    $("fillList").innerHTML = fills.map((f, k) => `<label class="fill"><input type="checkbox" data-k="${k}" ${state.fillSel.has(k) ? "checked" : ""}>
      <span><b>${esc(f.tab)} ${esc(f.col + f.row)}</b> · ${esc(FIELD[f.field] || f.field)} ← <span class="mono">${esc(f.value)}</span>
      <span class="muted">(${esc(f.instruction)} ${esc(f.container || "")})</span></span></label>`).join("");
    $("fillList").querySelectorAll("input").forEach((cb) => cb.onchange = () => { const k = +cb.dataset.k; cb.checked ? state.fillSel.add(k) : state.fillSel.delete(k); });

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
      </dl>
      ${issues.length ? `<div class="dissues">${issues.map((i) => `<div class="di ${i.ack ? "acked" : i.severity}">${esc(i.ack ? "👀 " + i.ack.by + ": " : "")}${esc(i.message)}</div>`).join("")}</div>` : ""}
      ${d.fills.length ? `<div class="dfills">${d.fills.map((f, k) => `<button class="ghost xs" data-fill="${k}">Fill ${esc(FIELD[f.field] || f.field)} ← ${esc(f.value)}</button>`).join("")}</div>` : ""}
      <div class="row-actions">
        <button class="primary small-btn" id="dNexl" type="button">Open in Nexl ↗</button>
        ${issues.some((i) => !i.ack) ? `<button class="ghost small-btn" id="dSnooze" type="button">👀 I'm on it</button>` : ""}
      </div>`;
    $("dClose").onclick = () => { box.hidden = true; state.selected = null; };
    $("dNexl").onclick = () => openNexl(word, d.container ? "container" : "instruction");
    if ($("dSnooze")) $("dSnooze").onclick = async () => { for (const i of issues.filter((x) => !x.ack)) await snooze(i); };
    box.querySelectorAll("[data-fill]").forEach((b) => (b.onclick = () => doFill([d.fills[+b.dataset.fill]])));
    if (!quiet) box.scrollIntoView({ block: "start", behavior: "smooth" });
  }

  // ---------- wiring ----------
  function showView(v) {
    state.view = v;
    document.querySelectorAll(".tabs button").forEach((b) => b.classList.toggle("active", b.dataset.view === v));
    for (const id of ["now", "issues", "live", "settings"]) $("view-" + id).hidden = id !== v;
  }

  function initSettingsUI() {
    const bind = (id, key, kind = "checked", parse = (x) => x) => {
      $(id)[kind] = settings[key];
      $(id).addEventListener("change", () => { settings[key] = parse($(id)[kind]); saveSettings(); });
    };
    bind("sName", "name", "value", (v) => v.trim());
    bind("sSplash", "splash");
    bind("sAuto", "auto");
    bind("sMinutes", "minutes", "value", (v) => Math.max(2, Math.min(60, +v || CFG.refreshMinutes)));
    bind("sSnoozeMin", "snoozeMin", "value", (v) => Math.max(10, Math.min(480, +v || 60)));
    bind("sStatusCols", "statusCols");
    bind("sNotes", "notes");
    bind("sWriteTab", "writeTab");
    bind("sWriteInfo", "writeInfo");
    bind("sNotStarted", "notStarted", "value", (v) => Math.max(5, +v || CFG.notStartedMinutes));
    bind("sStuck", "stuck", "value", (v) => Math.max(15, +v || CFG.stuckMinutes));
    bind("sRegion", "region", "value");
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
    $("syncBtn").addEventListener("click", sync);
    $("handoverBtn").addEventListener("click", showHandover);
    $("modalClose").addEventListener("click", closeModal);
    $("modal").addEventListener("click", (e) => { if (e.target.id === "modal") closeModal(); });
    document.querySelectorAll(".tabs button").forEach((b) => b.addEventListener("click", () => showView(b.dataset.view)));
    $("fSeverity").addEventListener("change", renderIssues);
    $("fTab").addEventListener("change", renderIssues);
    $("liveSearch").addEventListener("input", renderLive);
    $("showDone").addEventListener("change", renderLive);
    $("fillAll").addEventListener("click", () => { (state.result ? state.result.fills : []).forEach((_, k) => state.fillSel.add(k)); renderNow(); });
    $("fillSel").addEventListener("click", () => doFill([...state.fillSel].map((k) => state.result.fills[k]).filter(Boolean)));
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
    ExcelIO.watchSelection(CFG.tabs.map((t) => t.name), (tab, row) => showDetail(tab, row, true)).catch(() => {});
    sync();

    setInterval(() => {
      if (!settings.auto || state.busy || document.visibilityState === "hidden") return;
      if (Date.now() - state.lastSyncAt >= settings.minutes * 60000) sync();
    }, 20000);
  });
})();
