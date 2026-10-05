/*
 * Nexl Check — task pane controller.
 */
(function () {
  "use strict";
  const CFG = window.NEXL_CONFIG;
  const VERSION = "1.1.0";
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ---------- settings (per user, this browser) ----------
  const SKEY = "nexlcheck.settings.v1";
  const defaults = { auto: true, minutes: CFG.refreshMinutes, writeTab: true, writeInfo: false, region: CFG.region, disabledTabs: [],
    statusCols: true, notes: true, notStarted: CFG.notStartedMinutes, stuck: CFG.stuckMinutes };
  let settings = { ...defaults };
  try { settings = { ...defaults, ...JSON.parse(localStorage.getItem(SKEY) || "{}") }; } catch (e) { /* storage unavailable */ }
  const saveSettings = () => { try { localStorage.setItem(SKEY, JSON.stringify(settings)); } catch (e) { /* ignore */ } };

  // ---------- state ----------
  const state = { busy: false, result: null, lastSyncAt: 0, firstSeen: new Map(), view: "issues", tileFilter: null };
  const issueKey = (i) => [i.ref && i.ref.tab, i.ref && i.ref.row, i.field, i.container, i.instruction].join("|");
  const fmtTime = (d) => d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  const fmtStamp = (d) => d.toLocaleString([], { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });

  // ---------- UI helpers ----------
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
    b.style.borderLeftColor = kind === "warn" ? "var(--warn)" : "var(--err)";
    b.style.background = kind === "warn" ? "var(--warn-bg)" : "var(--err-bg)";
  }
  function progress(frac, text) {
    const p = $("progress");
    if (frac === null) { p.hidden = true; return; }
    p.hidden = false;
    p.querySelector("div").style.width = Math.round(frac * 100) + "%";
    p.querySelector("span").textContent = text;
  }

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

  // ---------- sync ----------
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
      const sheetTabs = await ExcelIO.readPlanTabs({ ...CFG, tabs: tabsCfg }, bases);
      const tabWarnings = sheetTabs.filter((t) => t.missing || t.error).map((t) => `${esc(t.name)}: ${t.missing ? "tab not found" : esc(t.error)}`);
      const usable = sheetTabs.filter((t) => !t.missing && !t.error);

      const onSheet = new Set();
      for (const t of usable) for (const r of t.rows) { const id = NexlParsers.normInstr(r.values[t.cols.instruction]); if (id) onSheet.add(NexlParsers.baseInstr(id)); }
      const ids = instructions.filter((i) => onSheet.has(i.base)).map((i) => i.id);

      const containers = await NexlClient.getContainers(ids, (d, n) => progress(0.25 + 0.55 * (d / n), `Reading Nexl containers ${d}/${n}…`));
      progress(0.82, "Reading driver tracking…");
      const tracking = await NexlClient.getTracking();
      const nexl = { instructions, containers, tracking };

      const mopts = { now: new Date(), notStartedMinutes: settings.notStarted, stuckMinutes: settings.stuck };
      let res = NexlMatcher.compare(usable, nexl, mopts);
      const unknown = [...new Set(res.issues.filter((i) => i.field === "container" && i.sheet && !i.nexl).map((i) => i.container))].slice(0, 25);
      if (unknown.length) {
        progress(0.9, `Searching Nexl for ${unknown.length} unknown container(s)…`);
        nexl.search = await NexlClient.searchContainers(unknown);
        res = NexlMatcher.compare(usable, nexl, mopts);
      }

      const now = new Date();
      for (const i of res.issues) { const k = issueKey(i); if (!state.firstSeen.has(k)) state.firstSeen.set(k, now); }
      state.result = res;
      state.lastSyncAt = Date.now();
      $("lastSync").textContent = "Synced " + fmtTime(now);
      render();

      if (settings.writeTab) {
        progress(0.96, `Updating ${CFG.checkTabName} tab…`);
        const list = res.issues.filter((i) => settings.writeInfo || i.severity !== "info");
        try {
          await ExcelIO.writeCheckTab(CFG, list, (i) => { const d = state.firstSeen.get(issueKey(i)); return d ? fmtStamp(d) : ""; }, fmtStamp(now));
        } catch (e) {
          tabWarnings.push(`Couldn't update the ${esc(CFG.checkTabName)} tab (${esc(e.message)}). The panel is still up to date.`);
        }
      }
      if (settings.statusCols && res.rowStatus.length) {
        progress(0.97, "Updating NEXL STEP / NEXL ALERT columns…");
        try { await ExcelIO.writeStatusColumns(CFG, res.rowStatus); }
        catch (e) { tabWarnings.push(`Couldn't update the status columns (${esc(e.message)}).`); }
      }
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
    } finally {
      progress(null);
      state.lastSyncAt = Date.now(); // also after failures, so auto-refresh waits a full interval before retrying
      state.busy = false;
      $("syncBtn").disabled = false;
    }
  }

  // ---------- rendering ----------
  const FIELD = { container: "Container", instruction: "Instruction", seal: "Seal", booking: "Booking ref", loadRef: "Load ref",
    vessel: "Vessel", customer: "Customer", transporter: "Transporter", driver: "Driver", progress: "Now" };

  function render() {
    const res = state.result;
    if (!res) return;
    const s = res.stats;
    $("tChecked").textContent = s.rowsChecked + s.missingFromSheet;
    $("tOk").textContent = s.matched;
    $("tErr").textContent = res.issues.filter((i) => i.severity === "error").length;
    $("tWarn").textContent = res.issues.filter((i) => i.severity === "warn").length;

    const tabSel = $("fTab"), cur = tabSel.value;
    const tabs = [...new Set(res.issues.map((i) => i.ref && i.ref.tab).filter(Boolean))];
    tabSel.innerHTML = '<option value="">All tabs</option>' + tabs.map((t) => `<option${t === cur ? " selected" : ""}>${esc(t)}</option>`).join("");
    renderIssues();
    renderLive();
  }

  function renderIssues() {
    const res = state.result;
    const el = $("issueList");
    if (!res) return;
    const sev = $("fSeverity").value, tab = $("fTab").value;
    const list = res.issues.filter((i) =>
      (sev === "all" || (sev === "problems" ? i.severity !== "info" : i.severity === sev)) && (!tab || (i.ref && i.ref.tab === tab)));
    if (!list.length) {
      el.innerHTML = `<p class="empty">${res.issues.length ? "Nothing for this filter." : "✓ Everything on the checked tabs matches Nexl."}</p>`;
      return;
    }
    el.innerHTML = list.slice(0, 400).map((i, k) => {
      const where = i.ref ? `${esc(i.ref.tab)} · ${esc((i.ref.col || "") + i.ref.row)}` : "";
      const since = state.firstSeen.get(issueKey(i));
      return `<div class="card ${i.severity}" data-k="${k}" title="Go to cell">
        <div class="where"><span>${where}</span><span>${since ? "since " + fmtTime(since) : ""}</span></div>
        <div class="what">${esc(i.message)}</div>
        <dl class="vals">
          <dt>Instr.</dt><dd class="mono">${esc(i.instruction)}</dd>
          <dt>Container</dt><dd class="mono">${esc(i.container)}</dd>
          ${i.field === "progress" ? `<dt>Now</dt><dd>${esc(i.nexl)}</dd>` : i.field !== "container" ? `<dt>${esc(FIELD[i.field] || i.field)}</dt><dd>sheet <b>${esc(i.sheet || "—")}</b> · Nexl <b>${esc(i.nexl || "—")}</b></dd>` : ""}
        </dl></div>`;
    }).join("") + (list.length > 400 ? `<p class="empty">…and ${list.length - 400} more (see the ${esc(CFG.checkTabName)} tab).</p>` : "");
    el.querySelectorAll(".card").forEach((c) => c.addEventListener("click", () => ExcelIO.goTo(list[+c.dataset.k].ref).catch(() => {})));
  }

  function podBadge(s) {
    const v = (s || "").toLowerCase();
    const cls = /verified|accepted|complete/.test(v) ? "ok" : /progress|pending|outstanding/.test(v) ? "warn" : /reject|query|fail/.test(v) ? "err" : "";
    return s ? `<span class="badge ${cls}">POD ${esc(s)}</span>` : `<span class="badge">No POD</span>`;
  }

  function renderLive() {
    const res = state.result;
    const el = $("liveList");
    if (!res) return;
    const q = $("liveSearch").value.trim().toUpperCase();
    const refs = [];
    const html = res.groups.map((g) => {
      const instr = g.nexl[0] || {};
      const rows = g.containers.filter((c) => {
        if (!q) return true;
        const n = c.nexl || {};
        return [c.container, c.id, n.driver, n.owner, g.base, instr.customer].join(" ").toUpperCase().includes(q);
      });
      if (!rows.length) return "";
      const state_ = [...new Set(g.nexl.map((i) => i.state))].join(" / ");
      const nIssues = g.containers.reduce((a, c) => a + c.issues.filter((i) => i.severity !== "info").length, 0);
      const nPod = g.containers.filter((c) => c.nexl && /verified|accepted/i.test(c.nexl.podStatus)).length;
      const open = q || nIssues ? " open" : "";
      return `<details class="group"${open}>
        <summary><div class="t"><span>${esc(g.base)} · ${esc(instr.customer || "")}</span><span class="badge ${state_ === "Completed" ? "ok" : "info"}">${esc(state_ || "Nexl")}</span></div>
          <div class="s">${esc([instr.vessel && instr.vessel !== "0" ? instr.vessel : "", instr.type].filter(Boolean).join(" · "))}</div>
          <div class="s">${g.containers.length} container(s) · ${nPod} POD verified${g.openSlots ? ` · ${g.openSlots} open slot(s)` : ""}${nIssues ? ` · <b style="color:var(--err)">${nIssues} issue(s)</b>` : ""}</div></summary>
        <div class="s route">${esc(instr.route || "")}${instr.completion ? ` · Done / in progress / not started / remaining: ${esc(instr.completion)}` : ""}</div>
        ${rows.map((c) => {
          const n = c.nexl || {}, t = c.tracking;
          const k = refs.push(c.sheet || g.firstRef) - 1;
          const bad = c.issues.filter((i) => i.severity !== "info").length;
          const sub = [
            n.moveStatus && "Move " + n.moveStatus,
            n.driver && n.driver.split(" ").slice(0, 2).join(" "),
            n.owner,
            (n.start || n.end) && `${n.start || "?"} → ${n.end || "…"}`,
          ].filter(Boolean).join(" · ");
          const track = t ? [t.route, t.pickup && "Pick-up " + t.pickup, t.via && "Via " + t.via, t.dropoff && "Drop-off " + t.dropoff, t.lastPing && "Last ping " + t.lastPing + " ago"].filter(Boolean).join(" · ") : "";
          return `<div class="crow" data-k="${k}">
            <span class="mono">${esc(c.container)}${c.id !== g.base ? ` <span class="muted">(${esc(c.id)})</span>` : ""}</span>
            <span>${c.pending ? '<span class="badge info">Allocated</span>' : c.nexl ? podBadge(n.podStatus) : '<span class="badge err">Not in Nexl</span>'}${bad ? ` <span class="badge err">${bad} issue${bad > 1 ? "s" : ""}</span>` : ""}${!c.sheet ? ' <span class="badge err">Not on sheet</span>' : ""}</span>
            ${c.step ? `<span class="sub step">🚚 ${esc(c.step)}</span>` : ""}
            ${sub ? `<span class="sub">${esc(sub)}</span>` : ""}
            ${track ? `<span class="sub">📍 ${esc(track)}</span>` : ""}
          </div>`;
        }).join("")}
      </details>`;
    }).join("");
    el.innerHTML = html || `<p class="empty">${q ? "No match." : "No Nexl instructions on the checked tabs."}</p>`;
    el.querySelectorAll(".crow").forEach((r) => r.addEventListener("click", () => ExcelIO.goTo(refs[+r.dataset.k]).catch(() => {})));

    const nop = res.notOnPlan;
    $("notOnPlan").hidden = !nop.length;
    $("nopCount").textContent = `(${nop.length})`;
    $("nopList").innerHTML = nop.map((i) => `<div><b class="mono">${esc(i.id)}</b> ${esc(i.customer)} · ${esc(i.type)} · ${esc(i.state)}${i.vessel && i.vessel !== "0" ? " · " + esc(i.vessel) : ""}</div>`).join("");
  }

  // ---------- wiring ----------
  function showView(v) {
    state.view = v;
    document.querySelectorAll(".tabs button").forEach((b) => b.classList.toggle("active", b.dataset.view === v));
    for (const id of ["issues", "live", "settings"]) $("view-" + id).hidden = id !== v;
  }

  function initSettingsUI() {
    $("sAuto").checked = settings.auto;
    $("sMinutes").value = settings.minutes;
    $("sWriteTab").checked = settings.writeTab;
    $("sWriteInfo").checked = settings.writeInfo;
    $("sRegion").value = settings.region;
    $("sStatusCols").checked = settings.statusCols;
    $("sNotes").checked = settings.notes;
    $("sNotStarted").value = settings.notStarted;
    $("sStuck").value = settings.stuck;
    $("sTabName").textContent = CFG.checkTabName;
    $("sTabs").innerHTML = CFG.tabs.map((t, k) =>
      `<label><span>${esc(t.name)}</span><input type="checkbox" data-k="${k}" ${settings.disabledTabs.includes(t.name) ? "" : "checked"}></label>`).join("");
    const on = (id, ev, fn) => $(id).addEventListener(ev, () => { fn(); saveSettings(); });
    on("sAuto", "change", () => (settings.auto = $("sAuto").checked));
    on("sMinutes", "change", () => (settings.minutes = Math.max(2, Math.min(60, +$("sMinutes").value || CFG.refreshMinutes))));
    on("sWriteTab", "change", () => (settings.writeTab = $("sWriteTab").checked));
    on("sWriteInfo", "change", () => (settings.writeInfo = $("sWriteInfo").checked));
    on("sRegion", "change", () => (settings.region = $("sRegion").value));
    on("sStatusCols", "change", () => (settings.statusCols = $("sStatusCols").checked));
    on("sNotes", "change", () => (settings.notes = $("sNotes").checked));
    on("sNotStarted", "change", () => (settings.notStarted = Math.max(5, +$("sNotStarted").value || CFG.notStartedMinutes)));
    on("sStuck", "change", () => (settings.stuck = Math.max(15, +$("sStuck").value || CFG.stuckMinutes)));
    $("sTabs").addEventListener("change", (e) => {
      const t = CFG.tabs[+e.target.dataset.k];
      settings.disabledTabs = settings.disabledTabs.filter((n) => n !== t.name);
      if (!e.target.checked) settings.disabledTabs.push(t.name);
      saveSettings();
    });
  }

  Office.onReady(async () => {
    initSettingsUI();
    $("syncBtn").addEventListener("click", sync);
    document.querySelectorAll(".tabs button").forEach((b) => b.addEventListener("click", () => showView(b.dataset.view)));
    $("fSeverity").addEventListener("change", renderIssues);
    $("fTab").addEventListener("change", renderIssues);
    $("liveSearch").addEventListener("input", renderLive);
    document.querySelectorAll(".tile").forEach((t) => t.addEventListener("click", () => {
      const f = t.dataset.filter;
      if (f === "ok" || f === "all") { showView("live"); return; }
      $("fSeverity").value = f; showView("issues"); renderIssues();
    }));

    const p = await NexlClient.ping();
    pill("bridgePill", p.ok, p.ok ? "Bridge" : "Bridge missing");
    $("versions").textContent = `Add-in v${VERSION}` + (p.ok ? ` · Bridge v${p.version}` : "");
    sync();

    setInterval(() => {
      if (!settings.auto || state.busy || document.visibilityState === "hidden") return;
      if (Date.now() - state.lastSyncAt >= settings.minutes * 60000) sync();
    }, 20000);
  });
})();
