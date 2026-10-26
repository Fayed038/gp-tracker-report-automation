/* HB-Tracker admin (HQ) pages */
const P = HB.pages;
const allOfficers = () => q(sb.from("hb_officers").select("*").order("name"));
const allSims = () => q(sb.from("hb_sims").select("*").order("number"));

/* Dashboard */
P.dashboard = async el => {
  const [offs, sims, reqs, rep, leads] = await Promise.all([
    q(sb.from("hb_officers").select("id,region_id,verified").eq("status", "active")), allSims(),
    q(sb.from("hb_requests").select("id,created_by").eq("status", "pending")), latestReport(),
    q(sb.from("hb_profiles").select("*").eq("role", "lead").eq("status", "active"))]);
  let rows = [];
  if (rep) rows = await q(sb.from("hb_report_rows").select("region_id,issue,justified_at").eq("report_id", rep.id).not("issue", "is", null));
  const since = new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10);
  const month = await q(sb.from("hb_report_rows").select("region_id,justified_at,report_date").gte("report_date", since).not("issue", "is", null));
  const waiting = rows.filter(r => !r.justified_at).length;
  const pool = sims.filter(s => s.status === "hq_pool").length;
  const field = HB.regions.filter(r => r.kind !== "hq" || offs.some(o => o.region_id === r.id));
  const lastBackup = localStorage.getItem("hb_last_backup");
  const backupDue = !lastBackup || Date.now() - +lastBackup > 7 * 864e5;
  el.innerHTML = head(`Welcome, ${esc(HB.me.full_name || "Admin")}`, rep ? `Latest report: ${fmtDate(rep.report_date)}` : "No report uploaded yet", `<button id="bk">Download full backup</button><a href="#/report"><button class="primary">Upload today's raw data</button></a>`) +
    (backupDue ? `<div class="warn">Weekly backup reminder: download a full backup and keep it safe.</div>` : "") +
    stats([["Officers", offs.length, `${offs.filter(o => !o.verified).length} not yet cross-matched`], ["Justifications waiting", waiting, rep ? "latest report" : "", "#C2410C"], ["Requests to approve", reqs.length, "", "#1F5FD1"], ["SIMs in HQ pool", pool, "not returned by officers"]]) +
    `<div class="card"><h2>Regions</h2>${table(["Region", "Lead", "Officers", "Not cross-matched", "Latest justified", "30-day score"], field.map(r => {
      const lead = leads.find(l => l.region_id === r.id); const o = offs.filter(x => x.region_id === r.id);
      const lr = rows.filter(x => x.region_id === r.id); const m = month.filter(x => x.region_id === r.id);
      const score = m.length ? Math.round(100 * m.filter(x => x.justified_at).length / m.length) + "%" : "—";
      return tr([esc(r.name) + (r.kind !== "field" ? " " + pill(r.kind.toUpperCase(), r.kind === "hq" ? "p-navy" : "p-amber") : ""), lead ? esc(lead.full_name || lead.email) : `<span class="muted">no lead</span>`, o.length, o.filter(x => !x.verified).length || "—", lr.length ? `${lr.filter(x => x.justified_at).length}/${lr.length}` : "—", score]);
    }), "Add regions in Users → Regions.")}</div>`;
  $("#bk").onclick = fullBackup;
};
async function fullBackup() {
  try {
    const t = ["hb_regions", "hb_profiles", "hb_officers", "hb_sims", "hb_requests", "hb_reports", "hb_report_rows", "hb_messages", "hb_announcements", "hb_audit"];
    const sheets = {};
    for (const n of t) { let all = [], from = 0; for (;;) { const d = await q(sb.from(n).select("*").range(from, from + 999)); all = all.concat(d); if (d.length < 1000) break; from += 1000; } sheets[n.replace("hb_", "")] = all.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, v && typeof v === "object" ? JSON.stringify(v) : v]))); }
    downloadXlsx(`HB-Tracker_backup_${today()}.xlsx`, sheets);
    localStorage.setItem("hb_last_backup", Date.now()); rpc("hb_backup_logged").catch(() => {});
    toast("Backup downloaded");
  } catch (e) { fail(e); }
}

/* Cross-match */
P.crossmatch = async el => {
  const offs = await q(sb.from("hb_officers").select("*").eq("status", "active"));
  el.innerHTML = head("Cross-match", "Compare officer lists with a GP Team Tracker export, then finalize a region") +
    `<div class="card"><div class="filters"><label>Region<select id="cmR">${regionOptions("", { all: "All regions" })}</select></label><label>GP Team Tracker export (.xlsx)<input type="file" id="cmF" accept=".xlsx,.xls" multiple></label><button class="primary" id="cmGo" style="align-self:flex-end">Compare</button></div><p class="muted" style="margin:0">Numbers are matched first; names are compared ignoring MD/Mst prefixes and word order.</p></div><div id="cmOut"></div>`;
  $("#cmGo").onclick = async () => {
    const files = [...$("#cmF").files]; if (!files.length) return toast("Choose the export file first");
    const rid = +$("#cmR").value || null;
    const tt = {};
    for (const f of files) {
      const wb = await readExcel(f);
      for (const sn of wb.SheetNames) { const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true }); rows.slice(1).forEach(r => { const mk = GPReport.normMobile(r[1]); if (mk && !tt[mk]) tt[mk] = String(r[0] || "").split("\n")[0].split("_")[0].trim(); }); }
    }
    const list = offs.filter(o => !rid || o.region_id === rid);
    const res = []; const seen = new Set();
    list.forEach(o => {
      const mk = GPReport.normMobile(o.sim_number);
      if (mk && tt[mk] !== undefined) { seen.add(mk); res.push({ kind: GPReport.nameKey(tt[mk]) === GPReport.nameKey(o.name) ? "Match" : "Different", o, ttName: tt[mk], mk }); }
      else res.push({ kind: "Only in list", o, mk });
    });
    if (!rid) Object.keys(tt).filter(k => !seen.has(k)).forEach(k => res.push({ kind: "Only in TT", ttName: tt[k], mk: k }));
    S.cm = { res, rid };
    drawCm("all");
  };
  function drawCm(f) {
    const { res, rid } = S.cm; const cnt = k => res.filter(r => r.kind === k).length;
    const shown = res.filter(r => f === "all" ? r.kind !== "Match" : r.kind === f);
    const cls = { Match: "p-green", Different: "p-blue", "Only in list": "p-amber", "Only in TT": "p-red" };
    $("#cmOut").innerHTML = stats([["Match", cnt("Match"), "same in both", "#15803D"], ["Different", cnt("Different"), "name differs", "#1F5FD1"], ["Only in officer list", cnt("Only in list"), "add to Team Tracker", "#B45309"], ["Only in Team Tracker", cnt("Only in TT"), "remove or add officer", "#C2410C"]]) +
      `<div class="tabs">${[["all", "All differences"], ["Different", "Different"], ["Only in list", "Only in list"], ["Only in TT", "Only in TT"], ["Match", "Matched"]].map(([k, l]) => `<button data-f="${k}" class="${k === f ? "on" : ""}">${l}</button>`).join("")}</div>` +
      `<div class="card">${table(["Officer list", "GP Team Tracker", "Result", "Action"], shown.map((r, i) => tr([r.o ? `${esc(r.o.name)} · ${esc(r.o.sim_number || "no SIM")}<br><span class="muted">${esc(regionName(r.o.region_id))}</span>` : "—", r.ttName !== undefined ? `${esc(r.ttName)} · 0${r.mk}` : "—", pill(r.kind, cls[r.kind]),
        r.kind === "Different" ? `<button class="sm" data-useTT="${res.indexOf(r)}">Use TT name</button> <span class="muted">or fix in TT</span>` : r.kind === "Only in list" ? "Add to Team Tracker" : r.kind === "Only in TT" ? "Remove from Team Tracker, or add officer" : "—"])), "No differences.")}
      <div class="acts" style="margin-top:14px"><button id="cmCsv">Download Team Tracker fix list</button>${rid ? `<button class="primary" id="cmFin">Finalize ${esc(regionName(rid))}</button>` : `<span class="muted">Pick a region to finalize it.</span>`}</div></div>`;
    $("#cmOut").querySelectorAll("[data-f]").forEach(b => b.onclick = () => drawCm(b.dataset.f));
    $("#cmOut").querySelectorAll("[data-usett]").forEach(b => b.onclick = async () => { const r = res[+b.dataset.usett]; try { await q(sb.from("hb_officers").update({ name: r.ttName }).eq("id", r.o.id)); r.o.name = r.ttName; r.kind = "Match"; toast("Name updated"); drawCm(f); } catch (e) { fail(e); } });
    $("#cmCsv").onclick = () => downloadCsv(`TT_fix_list_${today()}.csv`, res.filter(r => r.kind !== "Match").map(r => ({ Action: { Different: "Rename in TT", "Only in list": "Add to TT", "Only in TT": "Remove from TT (or add officer)" }[r.kind], Name: r.o?.name || r.ttName, "Name in TT": r.ttName || "", Number: r.o?.sim_number || "0" + r.mk, Region: r.o ? regionName(r.o.region_id) : "" })));
    const fin = $("#cmFin"); if (fin) fin.onclick = async () => { if (!confirm(`Mark all active officers of ${regionName(rid)} as verified?`)) return; try { await q(sb.from("hb_officers").update({ verified: true }).eq("region_id", rid).eq("status", "active")); toast("Region finalized"); } catch (e) { fail(e); } };
  }
};

/* Daily report */
P.report = async el => {
  el.innerHTML = head("Daily Report", "Upload the raw Team Tracker export. The report is saved and each lead gets their own list.", `<button id="rpPdf" disabled>Download PDF</button><button id="rpCsv" disabled>Download CSV</button>`) +
    `<div class="card"><div class="drop"><b>Today's raw Team Tracker export</b><div class="muted">.xlsx · several files are merged into one report</div><input type="file" id="rpF" accept=".xlsx,.xls" multiple style="margin-top:10px"></div><div class="acts" style="margin-top:12px"><span id="rpInfo" class="muted"></span><span style="flex:1"></span><button class="primary" id="rpBuild">Build report &amp; send lists to leads</button></div></div><div id="rpOut"></div>`;
  $("#rpBuild").onclick = async () => {
    const files = [...$("#rpF").files]; if (!files.length) return toast("Choose the export file(s) first");
    $("#rpBuild").disabled = true;
    try {
      let acc; for (const f of files) { const wb = await readExcel(f); for (const sn of wb.SheetNames) acc = GPReport.loadTracker(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false }), acc); }
      const offs = await q(sb.from("hb_officers").select("*").eq("status", "active").order("name"));
      if (!offs.length) throw new Error("There are no officers in the database yet.");
      const zones = HB.regions.map((r, i) => ({ code: r.id, title: r.name, sort: i }));
      const mapped = offs.map(o => ({ ...o, report_zone: o.region_id, report_order: 0, report_name: o.name, report_designation: o.designation, report_base: o.territory, tracker_number: o.sim_number }));
      const rep = GPReport.buildReport(mapped, zones, acc, C);
      const d = rep.date; const ds = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
      const exists = await q(sb.from("hb_reports").select("id").eq("report_date", ds));
      if (exists.length) throw new Error(`A report for ${fmtDate(ds)} is already saved. Saved reports cannot be replaced; open it in Report History.`);
      const issues = rep.zones.flatMap(z => z.rows).filter(r => r.remarks).length;
      const saved = (await q(sb.from("hb_reports").insert({ report_date: ds, summary: { ...rep.stats, issues, files: files.map(f => f.name), unmatchedCl: rep.unmatchedCl } }).select()))[0];
      const rows = rep.zones.flatMap(z => z.rows.map(r => ({ report_id: saved.id, report_date: ds, officer_id: r.oid, region_id: r.region_id, officer_name: r.name, sim_number: r.mobile, checkpoints: r, issue: /Inactive/.test(r.remarks) ? null : (r.remarks || null) })));
      for (let i = 0; i < rows.length; i += 500) await q(sb.from("hb_report_rows").insert(rows.slice(i, i + 500)));
      S.lastRep = rep;
      $("#rpOut").innerHTML = stats([["On report", rep.stats.officers], ["Located", rep.stats.officers - rep.stats.cl_tracker_set - rep.stats.no_raw_record - rep.stats.absent_set, "", "#15803D"], ["Couldn't locate", rep.stats.cl_tracker_set, "sent to leads", "#C2410C"], ["Absent / no record", rep.stats.absent_set + rep.stats.no_raw_record, "sent to leads", "#C2410C"]]) +
        (rep.unmatchedCl.length ? `<div class="warn">Team Tracker "Couldn't locate" names not matched to any officer:\n${rep.unmatchedCl.map(esc).join(", ")}</div>` : "") +
        `<div class="card">${renderReportHtml(rep.zones, rep.date)}</div>`;
      $("#rpPdf").disabled = $("#rpCsv").disabled = false;
      $("#rpPdf").onclick = () => reportPdf(rep.zones, rep.date, `GPS_Report_${ds}.pdf`);
      $("#rpCsv").onclick = () => downloadCsv(`GPS_Report_${ds}.csv`, reportCsvRows(rep.zones));
      toast("Report saved. Leads can now justify their lists."); refreshBell();
    } catch (e) { fail(e); } finally { $("#rpBuild").disabled = false; }
  };
};

/* Justifications */
P.justify = async el => {
  const reps = await q(sb.from("hb_reports").select("id,report_date").order("report_date", { ascending: false }).limit(60));
  el.innerHTML = head("Justifications", "Reasons from leads for officers who could not be tracked") + (reps.length ? `<div class="filters"><select id="jR">${reps.map(r => `<option value="${r.id}">${fmtDate(r.report_date)}</option>`).join("")}</select><select id="jRg">${regionOptions("", { all: "All regions" })}</select><select id="jS"><option value="">All status</option><option value="w">Waiting</option><option value="a">Answered</option></select><button id="jX">Export</button></div><div id="jOut"></div>` : `<div class="card empty">No reports yet.</div>`);
  if (!reps.length) return;
  const draw = async () => {
    let rows = await q(sb.from("hb_report_rows").select("*").eq("report_id", +$("#jR").value).not("issue", "is", null).order("region_id"));
    const rg = +$("#jRg").value, st = $("#jS").value;
    rows = rows.filter(r => (!rg || r.region_id === rg) && (!st || (st === "w" ? !r.justified_at : !!r.justified_at)));
    S.jrows = rows;
    $("#jOut").innerHTML = stats([["Sent", rows.length], ["Answered", rows.filter(r => r.justified_at).length, "", "#15803D"], ["Waiting", rows.filter(r => !r.justified_at).length, "", "#C2410C"]]) +
      `<div class="card">${table(["Officer", "Region", "Issue", "Reason", "Answered", "HQ"], rows.map(r => tr([esc(r.officer_name), esc(regionName(r.region_id)), esc(r.issue), r.reason ? esc(r.reason) + (r.note ? `<br><span class="muted">${esc(r.note)}</span>` : "") : pill("Waiting", "p-amber"), fmtTime(r.justified_at),
        r.justified_at ? (r.accepted === true ? pill("Accepted", "p-green") : r.accepted === false ? pill("Not accepted", "p-red") : `<button class="sm" data-a="${r.id}">Accept</button> <button class="sm" data-n="${r.id}">Not OK</button>`) : ""])))}</div>`;
    $("#jOut").querySelectorAll("[data-a],[data-n]").forEach(b => b.onclick = async () => { try { await rpc("hb_accept_reason", { p_row: +(b.dataset.a || b.dataset.n), p_ok: !!b.dataset.a }); draw(); } catch (e) { fail(e); } });
  };
  ["#jR", "#jRg", "#jS"].forEach(s => $(s).onchange = draw);
  $("#jX").onclick = () => downloadCsv(`Justifications_${today()}.csv`, S.jrows.map(r => ({ Date: r.report_date, Region: regionName(r.region_id), Officer: r.officer_name, Number: r.sim_number, Issue: r.issue, Reason: r.reason || "", Note: r.note || "", Answered: r.justified_at || "", Accepted: r.accepted ?? "" })));
  draw();
};

/* Approvals */
P.approvals = async el => {
  const tab = S.apTab || "pending";
  const [reqs, profs] = await Promise.all([q(sb.from("hb_requests").select("*").eq("status", tab).order("created_at", { ascending: false }).limit(200)), q(sb.from("hb_profiles").select("user_id,full_name,email"))]);
  const who = id => { const p = profs.find(x => x.user_id === id); return p ? esc(p.full_name || p.email) : ""; };
  el.innerHTML = head("Approvals", "Nothing a lead asks for changes until you approve it") +
    `<div class="tabs">${["pending", "approved", "rejected"].map(t => `<button data-t="${t}" class="${t === tab ? "on" : ""}">${t[0].toUpperCase() + t.slice(1)}</button>`).join("")}</div>` +
    `<div class="card">${table(["From", "Region", "Request", "Details", "Sent", tab === "pending" ? "" : "Decision"], reqs.map(r => tr([who(r.created_by), esc(regionName(r.region_id)), esc(REQ[r.kind]), esc(reqSummary(r)) + (r.payload.note ? `<br><span class="muted">${esc(r.payload.note)}</span>` : ""), fmtTime(r.created_at),
      tab === "pending" ? `<div class="acts"><button class="sm" data-rej="${r.id}">Reject</button><button class="sm primary" data-ok="${r.id}">Approve</button>${r.kind === "submit_list" ? `<button class="sm" data-view="${r.id}">View list</button>` : ""}</div>` : `${reqPill(r.status)} ${fmtTime(r.decided_at)}${r.decision_note ? `<br><span class="muted">${esc(r.decision_note)}</span>` : ""}`])), "No requests.")}</div>`;
  el.querySelectorAll("[data-t]").forEach(b => b.onclick = () => { S.apTab = b.dataset.t; reload(); });
  el.querySelectorAll("[data-view]").forEach(b => b.onclick = () => { const r = reqs.find(x => x.id === +b.dataset.view), p = r.payload;
    const sec = (t, rows) => rows ? `<h4 style="margin:8px 0 4px">${t}</h4>${rows}` : "";
    dialog(`<h3>Officer list update — ${esc(regionName(r.region_id))}</h3>` +
      sec(`New officers (${(p.add||[]).length})`, (p.add||[]).length && table(["Name", "Designation", "Territory", "SIM", "Personal"], p.add.map(o => tr([esc(o.name), esc(o.designation), esc(o.territory), esc(o.sim_number), esc(o.personal_number)])))) +
      sec(`Changed (${(p.update||[]).length})`, (p.update||[]).length && table(["Officer", "Change", "Remarks"], p.update.map(u => tr([esc(u.officer_name), Object.entries(u.changes).map(([k, v]) => `${k}: ${esc(u.before[k] || "—")} → <b>${esc(v || "—")}</b>`).join("<br>"), esc(u.remarks)])))) +
      sec(`Left (${(p.left||[]).length})`, (p.left||[]).length && table(["Officer", "SIM", "Remarks"], p.left.map(o => tr([esc(o.officer_name), esc(o.sim_number), esc(o.remarks)])))) +
      sec(`Missing from the file (${(p.missing||[]).length}) — not changed`, (p.missing||[]).length && table(["Officer", "SIM"], p.missing.map(o => tr([esc(o.officer_name), esc(o.sim_number)])))) +
      `<p class="muted">On approval: new officers are added, changes are saved, officers marked Left leave the report and their SIM goes to the HQ pool for you to decide. Missing officers are not changed.</p>`, async () => {}); });
  el.querySelectorAll("[data-ok],[data-rej]").forEach(b => b.onclick = async () => {
    const ok = !!b.dataset.ok; const r = reqs.find(x => x.id === +(b.dataset.ok || b.dataset.rej));
    if (!ok) { dialog(`<h3>Reject request</h3><label>Reason for the lead (optional)<textarea name="n"></textarea></label>`, async f => { await rpc("hb_decide", { p_id: r.id, p_approve: false, p_note: f.n.value || null, p_sim: null }); toast("Rejected"); reload(); refreshBell(); }); return; }
    if (r.kind === "transfer" && !r.payload.to_region_id) {
      dialog(`<h3>Approve transfer</h3><p>${esc(reqSummary(r))}</p><label>Move to region<select name="g">${regionOptions("")}</select></label>`, async f => { await rpc("hb_decide", { p_id: r.id, p_approve: true, p_note: null, p_sim: f.g.value }); toast("Approved"); reload(); refreshBell(); });
      return;
    }
    if (needsSim(r)) {
      const free = await q(sb.from("hb_sims").select("*").in("status", ["hq_pool", "spare", "spare_inactive"]).eq("locked", false).order("number"));
      dialog(`<h3>Approve: ${esc(REQ[r.kind])}</h3><p>${esc(reqSummary(r))}</p><label>SIM to give ${r.kind === "new_joiner" ? "(optional)" : ""}<select name="s">${r.kind === "new_joiner" ? `<option value="">No SIM yet</option>` : ""}${free.map(s => `<option value="${s.number}">${s.number} · ${SIM_LABEL[s.status][0]}${s.region_id ? " · " + esc(regionName(s.region_id)) : ""}</option>`).join("")}</select></label>`, async f => { await rpc("hb_decide", { p_id: r.id, p_approve: true, p_note: null, p_sim: f.s.value || null }); toast("Approved"); reload(); refreshBell(); });
      return;
    }
    try { await rpc("hb_decide", { p_id: r.id, p_approve: true, p_note: null, p_sim: null }); toast("Approved"); reload(); refreshBell(); } catch (e) { fail(e); }
  });
};

/* Officers */
P.officers = async el => {
  const [offs, sims] = await Promise.all([allOfficers(), allSims()]);
  el.innerHTML = head("Officers", "Final database · all regions", `<button id="oImp">Import Excel</button><button id="oExp">Export Excel</button><button class="primary" id="oAdd">Add officer</button>`) +
    `<div class="filters"><input id="oQ" placeholder="Search name, number, territory"><select id="oR">${regionOptions("", { all: "All regions" })}</select><select id="oS"><option value="active">Active</option><option value="">All status</option><option value="left">Left</option><option value="suspended">Suspended</option></select><input type="file" id="oFile" accept=".xlsx,.xls" class="hidden"></div><div class="card" id="oOut"></div>`;
  const draw = () => {
    const s = $("#oQ").value.toLowerCase(), rg = +$("#oR").value, st = $("#oS").value;
    const list = offs.filter(o => (!rg || o.region_id === rg) && (!st || o.status === st) && (!s || [o.name, o.sim_number, o.territory, o.personal_number].join(" ").toLowerCase().includes(s)));
    $("#oOut").innerHTML = `<h2>${list.length} officers</h2>` + table(["Name", "Designation", "Region", "Territory", "Corporate SIM", "Personal", "Status"], list.map(o => {
      const sim = sims.find(x => x.number === o.sim_number);
      return tr([esc(o.name) + (o.verified ? "" : ` ${pill("not cross-matched", "p-amber")}`), esc(o.designation), esc(regionName(o.region_id)), esc(o.territory), esc(o.sim_number || "—") + (sim?.locked ? " " + pill("HQ locked", "p-navy") : ""), esc(o.personal_number), pill(o.status, o.status === "active" ? "p-blue" : o.status === "left" ? "p-red" : "")], `class="click" data-id="${o.id}"`);
    }));
    $("#oOut").querySelectorAll("[data-id]").forEach(t => t.onclick = () => officerDialog(offs.find(o => o.id === +t.dataset.id)));
  };
  ["#oQ", "#oR", "#oS"].forEach(s => $(s).oninput = draw);
  $("#oAdd").onclick = () => officerDialog(null);
  $("#oExp").onclick = () => downloadXlsx(`Officers_${today()}.xlsx`, { Officers: offs.map(o => ({ Name: o.name, Designation: o.designation, Region: regionName(o.region_id), Territory: o.territory, "Corporate SIM": o.sim_number, "Personal number": o.personal_number, Status: o.status, Verified: o.verified ? "yes" : "no" })) });
  $("#oImp").onclick = () => $("#oFile").click();
  $("#oFile").onchange = async () => {
    try {
      const list = parseOfficerSheet(sheetObjects(await readExcel($("#oFile").files[0])));
      if (!list.length) throw new Error("No rows with a Name column found.");
      const missing = [...new Set(list.map(o => o.region).filter(r => r && !HB.regions.some(x => x.name.toLowerCase() === r.toLowerCase())))];
      dialog(`<h3>Import ${list.length} officers</h3><p>Columns read: Name, Designation, Region, Territory, Corporate SIM, Personal number.</p>${missing.length ? `<p>New regions will be created: <b>${missing.map(esc).join(", ")}</b></p>` : ""}<label><span><input type="checkbox" name="hq"> Lock the SIMs of officers in region "HQ"</span></label><p class="muted">Officers are added as "not cross-matched". Existing officers are not changed.</p>`, async f => {
        for (const m of missing) HB.regions.push((await q(sb.from("hb_regions").insert({ name: m, kind: /^hq$/i.test(m) ? "hq" : /leaf/i.test(m) ? "leaf" : "field" }).select()))[0]);
        const rid = n => (HB.regions.find(x => x.name.toLowerCase() === (n || "").toLowerCase()) || {}).id;
        const noReg = list.filter(o => !rid(o.region)); if (noReg.length) throw new Error(`${noReg.length} rows have no Region.`);
        const exist = offs.filter(o => o.status === "active"); const dup = o => exist.some(e => (o.sim_number && e.sim_number === o.sim_number) || (e.name.toLowerCase() === o.name.toLowerCase() && e.region_id === rid(o.region)));
        const fresh = list.filter(o => !dup(o)); if (!fresh.length) throw new Error("Everyone in this file is already in HB-Tracker. Nothing to import.");
        const ins = await q(sb.from("hb_officers").insert(fresh.map(o => ({ name: o.name, designation: o.designation, region_id: rid(o.region), territory: o.territory, sim_number: o.sim_number || null, personal_number: o.personal_number }))).select());
        const simRows = ins.filter(o => o.sim_number).map(o => ({ number: o.sim_number, region_id: o.region_id, status: "active", officer_id: o.id, locked: f.hq.checked && HB.regions.find(r => r.id === o.region_id).kind === "hq" }));
        if (simRows.length) await q(sb.from("hb_sims").upsert(simRows, { onConflict: "number", ignoreDuplicates: true }));
        toast(`${ins.length} officers imported, ${list.length - fresh.length} already existed (skipped)`); reload();
      });
    } catch (e) { fail(e); } finally { $("#oFile").value = ""; }
  };
  draw();
};
async function officerDialog(o) {
  const hist = o ? await q(sb.from("hb_audit").select("*").eq("action", "hb_officers_update").filter("details->>id", "eq", String(o.id)).order("at", { ascending: false }).limit(20)) : [];
  dialog(`<h3>${o ? "Edit officer" : "Add officer"}</h3><div class="fgrid">
    <label>Name<input name="name" required value="${esc(o?.name)}"></label><label>Designation<input name="designation" value="${esc(o?.designation)}"></label>
    <label>Region<select name="region_id">${regionOptions(o?.region_id)}</select></label><label>Territory<input name="territory" value="${esc(o?.territory)}"></label>
    <label>Corporate SIM<input name="sim_number" value="${esc(o?.sim_number)}" placeholder="017XXXXXXXX"></label><label>Personal number<input name="personal_number" value="${esc(o?.personal_number)}"></label>
    ${o ? `<label>Status<select name="status">${["active", "left", "suspended"].map(s => `<option ${o.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>` : ""}</div>
    ${hist.length ? `<details><summary>History (${hist.length})</summary>${hist.map(h => `<div class="muted" style="font-size:12px">${fmtTime(h.at)} · ${esc(h.details.designation || "")} · ${esc(regionName(h.details.region_id))} · ${esc(h.details.sim_number || "")} · ${esc(h.details.status)}</div>`).join("")}</details>` : ""}`, async f => {
    const v = { name: f.name.value.trim(), designation: f.designation.value.trim(), region_id: +f.region_id.value, territory: f.territory.value.trim(), sim_number: normNum(f.sim_number.value) || null, personal_number: normNum(f.personal_number.value) || f.personal_number.value.trim() };
    if (f.sim_number.value && !v.sim_number) throw new Error("SIM number must be 11 digits");
    if (o) { v.status = f.status.value; await q(sb.from("hb_officers").update(v).eq("id", o.id)); if (v.status !== "active" && o.sim_number) await q(sb.from("hb_sims").update({ officer_id: null, status: "hq_pool" }).eq("number", o.sim_number)); }
    else { const n = (await q(sb.from("hb_officers").insert({ ...v, verified: true }).select()))[0]; if (v.sim_number) await q(sb.from("hb_sims").upsert({ number: v.sim_number, region_id: v.region_id, status: "active", officer_id: n.id }, { onConflict: "number" })); }
    if (o && v.sim_number && v.sim_number !== o.sim_number) await q(sb.from("hb_sims").upsert({ number: v.sim_number, region_id: v.region_id, status: "active", officer_id: o.id }, { onConflict: "number" }));
    toast("Saved"); reload();
  });
}

/* SIM register */
const SIM_TABS = [["all", "All corporate"], ["seq", "Sequential (017 → 013)"], ["region", "Region wise"], ["spare", "Active spare"], ["spare_inactive", "Inactive spare"], ["suspended", "Suspended"], ["hq", "HQ (locked)"]];
P.sims = async el => {
  const [sims, offs] = await Promise.all([allSims(), allOfficers()]);
  const tab = S.simTab || "all";
  const holder = s => { const o = offs.find(x => x.id === s.officer_id); return o ? o.name : ""; };
  const seqSort = (a, b) => (a.number.startsWith("017") ? 0 : a.number.startsWith("013") ? 1 : 2) - (b.number.startsWith("017") ? 0 : b.number.startsWith("013") ? 1 : 2) || a.number.localeCompare(b.number);
  const sheet = t => {
    let l = sims.filter(s => s.status !== "closed" || t === "all");
    if (t === "spare") l = l.filter(s => s.status === "spare"); if (t === "spare_inactive") l = l.filter(s => s.status === "spare_inactive");
    if (t === "suspended") l = l.filter(s => s.status === "suspended"); if (t === "hq") l = l.filter(s => s.locked);
    if (t === "seq") l = l.slice().sort(seqSort); if (t === "region") l = l.slice().sort((a, b) => regionName(a.region_id).localeCompare(regionName(b.region_id)) || seqSort(a, b));
    return l;
  };
  const list = sheet(tab);
  el.innerHTML = head("SIM Register", "Same sheets as the master Excel", `<button id="sAdd">Add SIMs</button><button id="sExp">Export this sheet</button><button class="primary" id="sMaster">Download master Excel</button>`) +
    `<div class="tabs">${SIM_TABS.map(([k, l]) => `<button data-t="${k}" class="${k === tab ? "on" : ""}">${l}</button>`).join("")}</div>` +
    `<div class="filters"><input id="sQ" placeholder="Search number or holder"></div><div class="card" id="sOut"></div>`;
  const rowsOf = l => l.map((s, i) => ({ "#": i + 1, Number: s.number, Series: s.number.slice(0, 3), Holder: holder(s), Region: regionName(s.region_id), Status: s.locked ? "HQ locked" : SIM_LABEL[s.status][0], Note: s.note || "" }));
  const draw = () => {
    const qs = $("#sQ").value.toLowerCase();
    const l = list.filter(s => !qs || (s.number + " " + holder(s)).toLowerCase().includes(qs));
    $("#sOut").innerHTML = `<h2>${l.length} SIMs</h2>` + table(["#", "Number", "Holder", "Region", "Status", "Note", ""], l.map((s, i) => tr([i + 1, s.number, esc(holder(s)) || "—", esc(regionName(s.region_id)), simPill(s), esc(s.note), s.locked ? `<span class="muted">cannot be changed</span>` : `<button class="sm" data-n="${s.number}">Change</button>`])));
    $("#sOut").querySelectorAll("[data-n]").forEach(b => b.onclick = () => simDialog(sims.find(s => s.number === b.dataset.n)));
  };
  $("#sQ").oninput = draw;
  el.querySelectorAll("[data-t]").forEach(b => b.onclick = () => { S.simTab = b.dataset.t; reload(); });
  $("#sExp").onclick = () => downloadXlsx(`SIMs_${tab}_${today()}.xlsx`, { [SIM_TABS.find(x => x[0] === tab)[1].replace(/[^\w ]/g, "")]: rowsOf(list) });
  $("#sMaster").onclick = () => { const sh = {}; SIM_TABS.forEach(([k, l]) => sh[l.replace(/[^\w ()]/g, "").replace("017  013", "017-013")] = rowsOf(sheet(k))); sh.Dashboard = Object.entries(SIM_LABEL).map(([k, v]) => ({ Status: v[0], Count: sims.filter(s => s.status === k && !s.locked).length })).concat([{ Status: "HQ locked", Count: sims.filter(s => s.locked).length }, { Status: "Total", Count: sims.length }]); downloadXlsx(`SIM_Master_${today()}.xlsx`, sh); };
  $("#sAdd").onclick = () => dialog(`<h3>Add SIMs</h3><label>Numbers (one per line)<textarea name="n" placeholder="01700000000"></textarea></label><div class="fgrid"><label>Region<select name="r"><option value="">HQ pool (no region)</option>${regionOptions("")}</select></label><label>Status<select name="s"><option value="spare">Active spare</option><option value="spare_inactive">Inactive spare</option><option value="suspended">Suspended</option><option value="hq_pool">HQ pool</option></select></label></div><label><span><input type="checkbox" name="l"> Lock as HQ number (can never be changed)</span></label>`, async f => {
    const nums = f.n.value.split(/\s+/).map(normNum).filter(Boolean); if (!nums.length) throw new Error("Enter at least one 11-digit number");
    await q(sb.from("hb_sims").insert(nums.map(n => ({ number: n, region_id: +f.r.value || null, status: f.s.value, locked: f.l.checked }))));
    toast(`${nums.length} SIMs added`); reload();
  });
  draw();
};
function simDialog(s) {
  dialog(`<h3>SIM ${s.number}</h3><div class="fgrid"><label>Status<select name="s">${Object.entries(SIM_LABEL).filter(([k]) => k !== "active" || s.officer_id).map(([k, v]) => `<option value="${k}" ${s.status === k ? "selected" : ""}>${v[0]}</option>`).join("")}</select></label><label>Region<select name="r"><option value="">None</option>${regionOptions(s.region_id)}</select></label></div><label>Note<input name="note" value="${esc(s.note)}"></label><label><span><input type="checkbox" name="l"> Lock as HQ number (permanent)</span></label><p class="muted">Suspended: the number stays but nobody can use it or be given it.</p>`, async f => {
    const v = { status: f.s.value, region_id: +f.r.value || null, note: f.note.value || null };
    if (v.status !== "active") v.officer_id = null;
    if (f.l.checked && !confirm("Locking is permanent. Continue?")) return;
    await q(sb.from("hb_sims").update(v).eq("number", s.number));
    if (f.l.checked) await q(sb.from("hb_sims").update({ locked: true }).eq("number", s.number));
    if (s.officer_id && v.status !== "active") await q(sb.from("hb_officers").update({ sim_number: null }).eq("id", s.officer_id));
    toast("Saved"); reload();
  });
}

/* HQ spare pool */
P.pool = async el => {
  const sims = await q(sb.from("hb_sims").select("*").in("status", ["hq_pool", "suspended"]).order("number"));
  const pool = sims.filter(s => s.status === "hq_pool"), susp = sims.filter(s => s.status === "suspended");
  el.innerHTML = head("HQ Spare Pool", "SIMs not returned by officers, and SIMs HQ holds") + stats([["Awaiting decision", pool.length, "", "#B45309"], ["Suspended", susp.length]]) +
    `<div class="card">${table(["Number", "Last region", "Note", "Since", "Decision"], pool.map(s => tr([s.number, esc(regionName(s.region_id)), esc(s.note), fmtDate(s.updated_at), `<div class="acts"><button class="sm" data-a="suspended" data-n="${s.number}">Suspend</button><button class="sm" data-a="give" data-n="${s.number}">Give to region</button><button class="sm danger" data-a="closed" data-n="${s.number}">Close</button></div>`])), "No SIMs waiting.")}</div>`;
  el.querySelectorAll("[data-a]").forEach(b => b.onclick = async () => {
    const n = b.dataset.n, a = b.dataset.a;
    if (a === "give") return dialog(`<h3>Give ${n} to a region</h3><label>Region<select name="r">${regionOptions("", { fieldOnly: false })}</select></label><p class="muted">It becomes that lead's spare SIM.</p>`, async f => { await q(sb.from("hb_sims").update({ status: "spare", region_id: +f.r.value, note: "given by HQ" }).eq("number", n)); toast("Done"); reload(); });
    if (a === "closed" && !confirm(`Close ${n}? It stays in the register as closed.`)) return;
    try { await q(sb.from("hb_sims").update({ status: a }).eq("number", n)); toast("Done"); reload(); } catch (e) { fail(e); }
  });
};

/* Report history */
P.history = async el => {
  const reps = await q(sb.from("hb_reports").select("*").order("report_date", { ascending: false }).limit(400));
  el.innerHTML = head("Report History", "Every daily report is saved permanently and cannot be edited", `<button id="hBk">Download full backup</button>`) +
    `<div class="card">${table(["Date", "Officers", "Not tracked", "Files", ""], reps.map(r => tr([fmtDate(r.report_date), r.summary?.officers ?? "", r.summary?.issues ?? "", esc((r.summary?.files || []).join(", ")), `<div class="acts"><button class="sm" data-o="${r.id}">Open</button><button class="sm" data-p="${r.id}">PDF</button><button class="sm" data-c="${r.id}">CSV</button></div>`])), "No reports yet.")}</div><div id="hOut"></div>`;
  $("#hBk").onclick = fullBackup;
  const load = async id => { const rows = await q(sb.from("hb_report_rows").select("*").eq("report_id", id).order("id")); return { zones: zonesFromRows(rows), date: reps.find(r => r.id === id).report_date }; };
  el.querySelectorAll("[data-o]").forEach(b => b.onclick = async () => { const { zones, date } = await load(+b.dataset.o); $("#hOut").innerHTML = `<div class="card">${renderReportHtml(zones, date)}</div>`; $("#hOut").scrollIntoView(); });
  el.querySelectorAll("[data-p]").forEach(b => b.onclick = async () => { const { zones, date } = await load(+b.dataset.p); reportPdf(zones, date, `GPS_Report_${date}.pdf`); });
  el.querySelectorAll("[data-c]").forEach(b => b.onclick = async () => { const { zones, date } = await load(+b.dataset.c); downloadCsv(`GPS_Report_${date}.csv`, reportCsvRows(zones)); });
};

/* Users, regions, activity */
P.users = async el => {
  const tab = S.uTab || "users";
  el.innerHTML = head("Users", "Who can sign in, their role and region, and everything they do", `<button class="primary" id="uInv">Invite user</button>`) +
    `<div class="tabs">${[["users", "All users"], ["regions", "Regions"], ["log", "Activity log"]].map(([k, l]) => `<button data-t="${k}" class="${k === tab ? "on" : ""}">${l}</button>`).join("")}</div><div id="uOut"></div>`;
  el.querySelectorAll("[data-t]").forEach(b => b.onclick = () => { S.uTab = b.dataset.t; reload(); });
  $("#uInv").onclick = () => dialog(`<h3>Invite a user</h3><div class="fgrid"><label>Full name<input name="n" required></label><label>Email<input name="e" type="email" required></label><label>Role<select name="r"><option value="lead">Region lead</option><option value="admin">Admin</option></select></label><label>Region (leads)<select name="g">${regionOptions("", { fieldOnly: true })}</select></label></div><p class="muted">Then tell them: open HB-Tracker, click "First time? Set your password", and use this email.</p>`, async f => {
    if (f.r.value === "admin" && !confirm("Admins can see and change everything. Continue?")) return;
    await rpc("hb_invite", { p_email: f.e.value.trim(), p_name: f.n.value.trim(), p_role: f.r.value, p_region: f.r.value === "lead" ? +f.g.value : null });
    toast("Invited"); reload();
  });
  const out = $("#uOut");
  if (tab === "users") {
    const [profs, invs] = await Promise.all([q(sb.from("hb_profiles").select("*").order("role").order("full_name")), q(sb.from("hb_invites").select("*"))]);
    out.innerHTML = `<div class="card">${table(["Name", "Email", "Role", "Region", "Last active", "Status", ""], profs.map(p => tr([esc(p.full_name), esc(p.email),
      p.is_owner ? "Admin (owner)" : `<select data-role="${p.user_id}"><option value="lead" ${p.role === "lead" ? "selected" : ""}>Region lead</option><option value="admin" ${p.role === "admin" ? "selected" : ""}>Admin</option></select>`,
      p.role === "admin" ? "All" : `<select data-reg="${p.user_id}">${regionOptions(p.region_id, { fieldOnly: true })}</select>`, fmtTime(p.last_seen),
      pill(p.status, p.status === "active" ? "p-green" : "p-red"),
      p.is_owner ? `<span class="muted">cannot be removed</span>` : `<div class="acts">${p.status === "active" ? `<button class="sm" data-dis="${p.user_id}">Disable</button>` : `<button class="sm" data-en="${p.user_id}">Enable</button>`}<button class="sm danger" data-rm="${p.user_id}">Remove</button></div>`])))}</div>` +
      (invs.length ? `<div class="card"><h2>Invited, not signed up yet</h2>${table(["Name", "Email", "Role", "Region"], invs.map(i => tr([esc(i.full_name), esc(i.email), i.role, esc(regionName(i.region_id))])))}</div>` : "") +
      `<div class="note">One lead per region. Removing or disabling someone signs them out on every device immediately; their region's data stays saved.</div>`;
    const set = (u, o) => rpc("hb_set_user", { p_user: u, p_role: o.role ?? null, p_region: o.region ?? null, p_status: o.status ?? null }).then(() => { toast("Saved"); reload(); }).catch(e => { fail(e); reload(); });
    out.querySelectorAll("[data-role]").forEach(s => s.onchange = () => { if (s.value === "admin" && !confirm("Make this user an Admin? They will see everything.")) return reload(); set(s.dataset.role, { role: s.value }); });
    out.querySelectorAll("[data-reg]").forEach(s => s.onchange = () => set(s.dataset.reg, { region: +s.value }));
    out.querySelectorAll("[data-dis]").forEach(b => b.onclick = () => set(b.dataset.dis, { status: "disabled" }));
    out.querySelectorAll("[data-en]").forEach(b => b.onclick = () => set(b.dataset.en, { status: "active" }));
    out.querySelectorAll("[data-rm]").forEach(b => b.onclick = () => { const p = profs.find(x => x.user_id === b.dataset.rm); dialog(`<h3>Remove ${esc(p.full_name || p.email)}?</h3><p>Their ID stops working immediately and they are signed out everywhere. Officers, reports, requests and chats stay saved.</p><label>Type REMOVE to confirm<input name="c" required></label>`, async f => { if (f.c.value !== "REMOVE") throw new Error("Type REMOVE"); await rpc("hb_set_user", { p_user: p.user_id, p_role: null, p_region: null, p_status: "removed" }); toast("Access removed"); reload(); }); });
  } else if (tab === "regions") {
    out.innerHTML = `<div class="card"><div class="acts" style="margin-bottom:12px"><input id="rgN" placeholder="New region name"><select id="rgK"><option value="field">Field region</option><option value="leaf">Leaf</option><option value="hq">HQ</option></select><button class="primary" id="rgAdd">Add region</button></div>${table(["Region", "Type"], HB.regions.map(r => tr([esc(r.name), r.kind])))}</div>`;
    $("#rgAdd").onclick = async () => { const n = $("#rgN").value.trim(); if (!n) return; try { await q(sb.from("hb_regions").insert({ name: n, kind: $("#rgK").value })); HB.regions = await q(sb.from("hb_regions").select("*").order("name")); toast("Region added"); reload(); } catch (e) { fail(e); } };
  } else {
    const [log, profs] = await Promise.all([q(sb.from("hb_audit").select("*").order("at", { ascending: false }).limit(300)), q(sb.from("hb_profiles").select("user_id,full_name,email"))]);
    const who = id => { const p = profs.find(x => x.user_id === id); return p ? p.full_name || p.email : ""; };
    const label = a => ({ hb_officers_insert: "Officer added", hb_officers_update: "Officer changed", hb_sims_insert: "SIM added", hb_sims_update: "SIM changed", hb_requests_insert: "Request sent", hb_requests_update: "Request decided", hb_reports_insert: "Report uploaded" }[a] || a.replace(/_/g, " "));
    const det = l => { const d = l.details || {}; if (l.action.startsWith("hb_officers")) return `${d.name} · ${d.status}`; if (l.action.startsWith("hb_sims")) return `${d.number} · ${d.status}`; if (l.action.startsWith("hb_requests")) return `${REQ[d.kind]} · ${d.status}`; if (l.action === "hb_reports_insert") return fmtDate(d.report_date); return Object.entries(d).filter(([, v]) => v !== null && typeof v !== "object").map(([k, v]) => `${k}: ${v}`).join(" · "); };
    out.innerHTML = `<div class="card">${table(["When", "User", "Region", "Action", "Details"], log.map(l => tr([fmtTime(l.at), esc(who(l.user_id)), esc(regionName(l.region_id)), esc(label(l.action)), esc(det(l))])))}<p class="muted">This log is permanent: nobody, including admins, can edit or delete it.</p></div>`;
  }
};

/* Messages */
P.messages = async el => {
  const [leads, msgs, anns] = await Promise.all([q(sb.from("hb_profiles").select("*").eq("role", "lead").neq("status", "removed").order("full_name")), q(sb.from("hb_messages").select("*").order("created_at")), q(sb.from("hb_announcements").select("*").order("created_at", { ascending: false }).limit(10))]);
  const cur = S.chatLead || leads[0]?.user_id;
  el.innerHTML = head("Messages", "One private conversation per lead. Leads cannot see or message each other.", `<button id="mAnn">Announcement to all leads</button>`) +
    (leads.length ? `<div class="chat"><div class="card convs">${leads.map(l => { const m = msgs.filter(x => x.lead_id === l.user_id); const un = m.filter(x => !x.read_at && x.sender !== HB.me.user_id).length; const last = m[m.length - 1]; return `<div class="conv ${l.user_id === cur ? "on" : ""}" data-l="${l.user_id}"><div style="flex:1;min-width:0"><b>${esc(l.full_name || l.email)}</b><div class="muted" style="font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(regionName(l.region_id))}${last ? " · " + esc(last.body.slice(0, 40)) : ""}</div></div>${un ? `<span class="badge">${un}</span>` : ""}</div>`; }).join("")}</div><div id="thread" class="card thread"></div></div>` : `<div class="card empty">No leads yet. Invite them from Users.</div>`) +
    (anns.length ? `<div class="card"><h2>Recent announcements</h2>${anns.map(a => `<div style="border-left:3px solid var(--blue);padding-left:12px;margin-bottom:10px"><b>${esc(a.title)}</b><div>${esc(a.body)}</div><div class="muted" style="font-size:12px">${fmtTime(a.created_at)}</div></div>`).join("")}</div>` : "");
  el.querySelectorAll("[data-l]").forEach(c => c.onclick = () => { S.chatLead = c.dataset.l; reload(); });
  $("#mAnn").onclick = () => dialog(`<h3>Announcement to all leads</h3><label>Title<input name="t" required></label><label>Message<textarea name="b"></textarea></label>`, async f => { await q(sb.from("hb_announcements").insert({ title: f.t.value, body: f.b.value })); toast("Sent"); reload(); });
  if (cur) { const l = leads.find(x => x.user_id === cur); drawThread($("#thread"), cur, `${l.full_name || l.email}`, `${regionName(l.region_id)} region · private conversation`, msgs.filter(m => m.lead_id === cur)); }
};
function drawThread(box, leadId, title, sub, msgs) {
  box.innerHTML = `<div class="th"><b>${esc(title)}</b><div class="muted">${esc(sub)}</div></div><div class="msgs" id="msgs">${msgs.map(m => `<div class="bub ${m.sender === HB.me.user_id ? "me" : ""}">${esc(m.body)}<small>${fmtTime(m.created_at)}${m.sender === HB.me.user_id && m.read_at ? " · seen" : ""}</small></div>`).join("") || `<div class="empty">No messages yet.</div>`}</div><form class="composer" id="cmp"><input id="cmpT" placeholder="Write a message" maxlength="4000" autocomplete="off"><button class="primary">Send</button></form>`;
  const m = $("#msgs"); m.scrollTop = m.scrollHeight;
  rpc("hb_mark_read", { p_lead: leadId }).then(refreshBell).catch(() => {});
  $("#cmp").onsubmit = async e => { e.preventDefault(); const t = $("#cmpT").value.trim(); if (!t) return; try { await q(sb.from("hb_messages").insert({ lead_id: leadId, body: t })); reload(); } catch (er) { fail(er); } };
}
