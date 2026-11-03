/* HB-Tracker admin extras: Team Tracker roster, scorecard, officer profile */

/* ---------- Team Tracker roster (latest upload is the truth) ---------- */
const TT = {
  async load() {
    let all = [], from = 0;
    for (;;) { const part = await q(sb.from("hb_tt").select("mobile,name,batch").range(from, from + 999)); all = all.concat(part); if (part.length < 1000) break; from += 1000; }
    if (!all.length) return { map: {}, batch: null };
    const batch = all.reduce((m, r) => r.batch > m ? r.batch : m, all[0].batch);
    const map = {}; all.filter(r => r.batch === batch).forEach(r => { const k = GPReport.normMobile(r.mobile); if (k) map[k] = r.name; });
    return { map, batch };
  },
  async save(tt) { // tt: {mobileKey: name}
    const batch = new Date().toISOString();
    const rows = Object.entries(tt).map(([k, name]) => ({ mobile: "0" + k, name, batch }));
    for (let i = 0; i < rows.length; i += 500) await q(sb.from("hb_tt").upsert(rows.slice(i, i + 500), { onConflict: "mobile" }));
    return batch;
  },
  status(roster, name, sim) {
    const k = GPReport.normMobile(sim);
    if (!roster.batch) return "";
    if (!k) return pill("No SIM", "");
    if (roster.map[k] === undefined) return pill("Not in Team Tracker", "p-red");
    const tn = roster.map[k];
    return GPReport.nameKey(tn) === GPReport.nameKey(name) ? pill("In Team Tracker", "p-green") : `${pill("Name differs", "p-blue")}<br><span class="muted" style="font-size:12px">TT: ${esc(tn.split("_")[0])}</span>`;
  },
  note(roster) { return roster.batch ? `<p class="muted" style="margin:4px 0">Team Tracker list of ${fmtDate(roster.batch)} (${Object.keys(roster.map).length} users). Upload a newer export in Cross-match to refresh it.</p>` : `<p class="muted">No Team Tracker list saved yet. Upload an export in Cross-match.</p>`; }
};

/* ---------- Monthly scorecard ---------- */
P.scorecard = async el => {
  const now = new Date(), ym = d => d.toISOString().slice(0, 7);
  const months = Array.from({ length: 12 }, (_, i) => ym(new Date(now.getFullYear(), now.getMonth() - i, 15)));
  el.innerHTML = head("Region Scorecard", "Tracking and justification performance by region for one month", `<select id="scM">${months.map(m => `<option value="${m}">${new Date(m + "-15").toLocaleDateString("en-GB", { month: "long", year: "numeric" })}</option>`).join("")}</select><button class="primary" id="scPdf">Download PDF</button>`) + `<div id="scOut"></div>`;
  const dlRow = (await q(sb.from("hb_settings").select("value").eq("key", "justify_deadline")))[0];
  const [dh, dm] = String(dlRow?.value || "18:00").split(":").map(Number);
  let html = "";
  const draw = async () => {
    const m = $("#scM").value, start = m + "-01", end = ym(new Date(+m.slice(0, 4), +m.slice(5, 7), 15)) + "-01";
    let rows = [], from = 0;
    for (;;) { const part = await q(sb.from("hb_report_rows").select("region_id,officer_id,officer_name,report_date,issue,sent_at,justified_at,accepted").gte("report_date", start).lt("report_date", end).range(from, from + 999)); rows = rows.concat(part); if (part.length < 1000) break; from += 1000; }
    const days = new Set(rows.map(r => r.report_date)).size;
    const regs = S.regions.filter(r => r.kind !== "hq");
    const pct = (a, b) => b ? Math.round(100 * a / b) + "%" : "—";
    const onTime = r => { if (!r.justified_at || !r.sent_at) return false; const s = new Date(r.sent_at), dl = new Date(s.getFullYear(), s.getMonth(), s.getDate(), dh, dm); if (s > dl) dl.setDate(dl.getDate() + 1); return new Date(r.justified_at) <= dl; };
    const lines = regs.map(g => {
      const x = rows.filter(r => r.region_id === g.id); const iss = x.filter(r => r.issue); const sent = iss.filter(r => r.sent_at);
      const ans = sent.filter(r => r.justified_at), ot = sent.filter(onTime), acc = ans.filter(r => r.accepted === true);
      return { g, checks: x.length, tracked: x.length - iss.length, iss: iss.length, sent: sent.length, ans: ans.length, ot: ot.length, acc: acc.length };
    }).filter(l => l.checks);
    const reps = {}; rows.filter(r => r.issue).forEach(r => { const k = r.officer_id || r.officer_name; (reps[k] = reps[k] || { name: r.officer_name, region_id: r.region_id, n: 0 }).n++; });
    const repeat = Object.values(reps).filter(r => r.n >= 3).sort((a, b) => b.n - a.n);
    const tot = lines.reduce((a, l) => ({ checks: a.checks + l.checks, tracked: a.tracked + l.tracked, sent: a.sent + l.sent, ans: a.ans + l.ans, ot: a.ot + l.ot }), { checks: 0, tracked: 0, sent: 0, ans: 0, ot: 0 });
    const title = new Date(m + "-15").toLocaleDateString("en-GB", { month: "long", year: "numeric" });
    html = `<div class="rep-zone"><div class="rep-title">Region Scorecard — ${esc(title)}</div><div class="rep-date">${days} report day${days === 1 ? "" : "s"}</div>
      <table class="rep" style="margin-top:6px"><thead><tr><th>Region</th><th>Officer-days</th><th>Tracked</th><th>Tracked %</th><th>Asked to justify</th><th>Answered</th><th>On time %</th><th>Accepted</th></tr></thead><tbody>
      ${lines.map(l => `<tr><td>${esc(l.g.name)}</td><td>${l.checks}</td><td>${l.tracked}</td><td><b>${pct(l.tracked, l.checks)}</b></td><td>${l.sent}</td><td>${l.ans}</td><td class="${l.sent && l.ot / l.sent < .8 ? "red" : ""}"><b>${pct(l.ot, l.sent)}</b></td><td>${l.acc}</td></tr>`).join("")}
      <tr><td><b>All regions</b></td><td><b>${tot.checks}</b></td><td><b>${tot.tracked}</b></td><td><b>${pct(tot.tracked, tot.checks)}</b></td><td><b>${tot.sent}</b></td><td><b>${tot.ans}</b></td><td><b>${pct(tot.ot, tot.sent)}</b></td><td></td></tr></tbody></table>
      ${repeat.length ? `<div class="rep-title" style="margin-top:14px">Officers untracked 3 or more times</div><table class="rep"><thead><tr><th>Officer</th><th>Region</th><th>Times untracked</th></tr></thead><tbody>${repeat.map(r => `<tr><td>${esc(r.name)}</td><td>${esc(regionName(r.region_id))}</td><td class="red">${r.n}</td></tr>`).join("")}</tbody></table>` : ""}
      <div class="wm">Tracked % = officer-days located at all checkpoints. On time = answered by ${String(dh).padStart(2, "0")}:${String(dm).padStart(2, "0")} on the day HQ sent the request. Generated ${fmtTime(new Date())} · HB-Tracker</div></div>`;
    $("#scOut").innerHTML = lines.length ? stats([["Report days", days], ["Tracked", pct(tot.tracked, tot.checks), "all regions", "#15803D"], ["Justified on time", pct(tot.ot, tot.sent), `${tot.sent} asked`, "#1F5FD1"], ["Repeat untracked", repeat.length, "3+ times this month", "#C2410C"]]) + `<div class="card">${html}</div>` : `<div class="card empty">No reports saved for ${esc(title)}.</div>`;
    $("#scPdf").disabled = !lines.length;
  };
  $("#scM").onchange = draw;
  $("#scPdf").onclick = () => { const box = document.createElement("div"); box.className = "pdfmode"; box.innerHTML = html; html2pdf().set({ margin: [10, 8, 10, 8], filename: `Region_Scorecard_${$("#scM").value}.pdf`, image: { type: "jpeg", quality: .95 }, html2canvas: { scale: 2 }, jsPDF: { unit: "mm", format: "a4", orientation: "landscape" }, pagebreak: { mode: ["css", "legacy"], avoid: "tr" } }).from(box).save(); };
  draw();
};

/* ---------- Officer profile ---------- */
async function officerProfile(o) {
  const since = new Date(Date.now() - 90 * 864e5).toISOString().slice(0, 10);
  const [rows, reqs, roster] = await Promise.all([
    q(sb.from("hb_report_rows").select("report_date,issue,reason,note,sent_at,justified_at,accepted").eq("officer_id", o.id).gte("report_date", since).order("report_date", { ascending: false })),
    q(sb.from("hb_requests").select("kind,status,created_at,decided_at,payload").eq("region_id", o.region_id).order("created_at", { ascending: false }).limit(200)),
    TT.load()
  ]);
  const mine = reqs.filter(r => String(r.payload?.officer_id) === String(o.id) || (r.payload?.update || []).some(u => String(u.officer_id) === String(o.id)) || (r.payload?.left || []).some(u => String(u.officer_id) === String(o.id)));
  const iss = rows.filter(r => r.issue), last30 = iss.filter(r => r.report_date >= new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10));
  dialog(`<h3>${esc(o.name)}</h3><p class="muted" style="margin-top:0">${esc(o.designation || "")} · ${esc(regionName(o.region_id))}${o.territory ? " · " + esc(o.territory) : ""}</p>
    <div class="fgrid"><div><b>Corporate SIM</b><br>${esc(o.sim_number || "—")}</div><div><b>Personal</b><br>${esc(o.personal_number || "—")}</div><div><b>Status</b><br>${esc(o.status)}</div><div><b>Team Tracker</b><br>${TT.status(roster, o.name, o.sim_number) || "—"}</div></div>
    <p style="margin:12px 0 4px"><b>Untracked:</b> ${last30.length} time${last30.length === 1 ? "" : "s"} in the last 30 days, ${iss.length} in 90 days ${last30.length >= 3 ? pill("Repeat", "p-red") : ""}</p>
    ${iss.length ? table(["Date", "Issue", "Reason", "HQ"], iss.slice(0, 15).map(r => tr([fmtDate(r.report_date), esc(r.issue), r.reason ? esc(r.reason) + (r.note ? `<br><span class="muted">${esc(r.note)}</span>` : "") : r.sent_at ? pill("Waiting", "p-amber") : '<span class="muted">not sent</span>', r.accepted === true ? pill("Accepted", "p-green") : r.accepted === false ? pill("Not OK", "p-red") : ""]))) : `<p class="muted">Tracked at every checkpoint in the last 90 days.</p>`}
    ${mine.length ? `<p style="margin:12px 0 4px"><b>Requests</b></p>${table(["Request", "Status", "Date"], mine.slice(0, 10).map(r => tr([esc(REQ[r.kind]), reqPill(r.status), fmtDate(r.decided_at || r.created_at)])))}` : ""}
    <p style="margin-top:12px"><button type="button" class="sm" id="opEdit">Edit officer</button></p>`, async () => {});
  const ok = $("#dlgOk"); if (ok) ok.textContent = "Close";
  $("#opEdit").onclick = () => { $("#dlg").close(); officerDialog(o); };
}

/* ---------- Stored files (Supabase Storage, admin only) ---------- */
const FILES = {
  KINDS: { daily_raw: "Daily report raw file", team_tracker: "Team Tracker list", track_history: "Track History", other: "Other" },
  async save(file, kind, report_date = null) {
    try {
      const d = new Date().toISOString().slice(0, 10), safe = file.name.replace(/[^\w.\-]+/g, "_");
      const path = `${kind}/${d}/${Date.now()}_${safe}`;
      const { error } = await sb.storage.from("hb-files").upload(path, file, { upsert: false, contentType: file.type || "application/octet-stream" });
      if (error) throw error;
      await q(sb.from("hb_files").insert({ kind, name: file.name, path, size: file.size, report_date }));
      return true;
    } catch (e) { console.error(e); toast("File could not be stored: " + (e.message || e), 5000); return false; }
  },
  async open(path, name) {
    const { data, error } = await sb.storage.from("hb-files").createSignedUrl(path, 120, { download: name });
    if (error) return fail(error);
    const a = document.createElement("a"); a.href = data.signedUrl; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  },
  size: n => n == null ? "" : n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB"
};

P.files = async el => {
  el.innerHTML = head("Files", "Every file uploaded to HB-Tracker is stored here. Delete what is no longer needed.", `<input type="file" id="fUp" multiple class="hidden"><button id="fAdd">Upload a file</button>`) +
    `<div class="filters"><select id="fK"><option value="">All types</option>${Object.entries(FILES.KINDS).map(([k, l]) => `<option value="${k}">${l}</option>`).join("")}</select><label>From<input type="date" id="fFrom"></label><label>To<input type="date" id="fTo"></label><input id="fQ" placeholder="Search file name"></div><div id="fOut"></div>`;
  let files = [];
  const loadAll = async () => { files = await q(sb.from("hb_files").select("*").order("uploaded_at", { ascending: false }).limit(2000)); draw(); };
  const draw = () => {
    const k = $("#fK").value, a = $("#fFrom").value, b = $("#fTo").value, s = $("#fQ").value.toLowerCase();
    const list = files.filter(f => (!k || f.kind === k) && (!a || f.uploaded_at.slice(0, 10) >= a) && (!b || f.uploaded_at.slice(0, 10) <= b) && (!s || f.name.toLowerCase().includes(s)));
    const total = list.reduce((t, f) => t + (f.size || 0), 0);
    $("#fOut").innerHTML = `<div class="card"><h2>${list.length} file${list.length === 1 ? "" : "s"} <span class="muted" style="font-weight:400;font-size:14px">· ${FILES.size(total)}</span></h2>${table(["File", "Type", "Report date", "Uploaded", "Size", ""], list.map(f => tr([esc(f.name), esc(FILES.KINDS[f.kind] || f.kind), f.report_date ? fmtDate(f.report_date) : "", fmtTime(f.uploaded_at), FILES.size(f.size), `<div class="acts"><button class="sm" data-dl="${f.id}">Download</button><button class="sm danger" data-del="${f.id}">Delete</button></div>`])), "No files stored yet. Files are saved automatically when you upload them in Daily Report, Cross-match or Track History.")}</div>` +
      `<div class="note">Deleting a file removes only the file. Saved daily reports stay in Report History.</div>`;
    $("#fOut").querySelectorAll("[data-dl]").forEach(x => x.onclick = () => { const f = files.find(y => y.id === +x.dataset.dl); FILES.open(f.path, f.name); });
    $("#fOut").querySelectorAll("[data-del]").forEach(x => x.onclick = () => { const f = files.find(y => y.id === +x.dataset.del); dialog(`<h3>Delete this file?</h3><p><b>${esc(f.name)}</b><br><span class="muted">${esc(FILES.KINDS[f.kind] || f.kind)} · ${fmtTime(f.uploaded_at)}</span></p><p>The file is removed permanently.${f.kind === "daily_raw" ? " The saved daily report itself is not affected." : ""}</p>`, async () => {
      const { error } = await sb.storage.from("hb-files").remove([f.path]); if (error) throw error;
      await q(sb.from("hb_files").delete().eq("id", f.id)); toast("File deleted"); loadAll(); }); });
  };
  ["#fK", "#fFrom", "#fTo", "#fQ"].forEach(s => $(s).oninput = draw);
  $("#fAdd").onclick = () => $("#fUp").click();
  $("#fUp").onchange = async () => { for (const f of $("#fUp").files) await FILES.save(f, "other"); toast("Stored"); loadAll(); };
  loadAll();
};

/* ---------- Report History with filters (replaces the simple list) ---------- */
P.history = async el => {
  const [reps, stored] = await Promise.all([q(sb.from("hb_reports").select("*").order("report_date", { ascending: false }).limit(1000)), q(sb.from("hb_files").select("id,name,path,report_date").eq("kind", "daily_raw")).catch(() => [])]);
  el.innerHTML = head("Report History", "Every daily report is saved permanently and cannot be edited", `<button id="hBk">Download full backup</button>`) +
    `<div class="filters"><label>From<input type="date" id="hFrom"></label><label>To<input type="date" id="hTo"></label><label>Region<select id="hR">${regionOptions("", { all: "All regions" })}</select></label><input id="hQ" placeholder="Officer name or number (search inside reports)"><button id="hGo">Search</button></div><div id="hList"></div><div id="hOut"></div>`;
  $("#hBk").onclick = fullBackup;
  const rid = () => +$("#hR").value || null;
  const load = async id => { let rows = await q(sb.from("hb_report_rows").select("*").eq("report_id", id).order("id")); if (rid()) rows = rows.filter(r => r.region_id === rid()); return { zones: zonesFromRows(rows), date: reps.find(r => r.id === id).report_date }; };
  const tag = () => rid() ? "_" + regionName(rid()).replace(/\W+/g, "_") : "";
  const drawList = () => {
    const a = $("#hFrom").value, b = $("#hTo").value;
    const list = reps.filter(r => (!a || r.report_date >= a) && (!b || r.report_date <= b));
    $("#hList").innerHTML = `<div class="card"><h2>${list.length} report${list.length === 1 ? "" : "s"}${rid() ? " · " + esc(regionName(rid())) + " only" : ""}</h2>${table(["Date", "Officers", "Not tracked", "Raw files", ""], list.map(r => { const fs = stored.filter(f => f.report_date === r.report_date); return tr([fmtDate(r.report_date), r.summary?.officers ?? "", r.summary?.issues ?? "", fs.length ? fs.map(f => `<a href="#" data-f="${f.id}">${esc(f.name)}</a>`).join("<br>") : esc((r.summary?.files || []).join(", ")), `<div class="acts"><button class="sm" data-o="${r.id}">Open</button><button class="sm" data-p="${r.id}">PDF</button><button class="sm" data-c="${r.id}">CSV</button></div>`]); }), "No reports in this period.")}</div>`;
    $("#hList").querySelectorAll("[data-f]").forEach(x => x.onclick = e => { e.preventDefault(); const f = stored.find(y => y.id === +x.dataset.f); FILES.open(f.path, f.name); });
    $("#hList").querySelectorAll("[data-o]").forEach(x => x.onclick = async () => { const { zones, date } = await load(+x.dataset.o); $("#hOut").innerHTML = `<div class="card">${zones.length ? renderReportHtml(zones, date) : "No officers of this region in that report."}</div>`; $("#hOut").scrollIntoView(); });
    $("#hList").querySelectorAll("[data-p]").forEach(x => x.onclick = async () => { const { zones, date } = await load(+x.dataset.p); reportPdf(zones, date, `GPS_Report_${date}${tag()}.pdf`); });
    $("#hList").querySelectorAll("[data-c]").forEach(x => x.onclick = async () => { const { zones, date } = await load(+x.dataset.c); downloadCsv(`GPS_Report_${date}${tag()}.csv`, reportCsvRows(zones)); });
  };
  ["#hFrom", "#hTo", "#hR"].forEach(s => $(s).onchange = () => { $("#hOut").innerHTML = ""; drawList(); });
  $("#hGo").onclick = async () => {
    const s = $("#hQ").value.trim(); if (!s) { $("#hOut").innerHTML = ""; return; }
    const a = $("#hFrom").value, b = $("#hTo").value; const dig = s.replace(/\D/g, "");
    let qq = sb.from("hb_report_rows").select("report_date,region_id,officer_name,sim_number,issue,reason,note,accepted").order("report_date", { ascending: false }).limit(500);
    qq = dig.length >= 5 ? qq.ilike("sim_number", `%${dig.slice(-6)}%`) : qq.ilike("officer_name", `%${s}%`);
    if (a) qq = qq.gte("report_date", a); if (b) qq = qq.lte("report_date", b); if (rid()) qq = qq.eq("region_id", rid());
    const rows = await q(qq);
    $("#hOut").innerHTML = `<div class="card"><h2>${rows.length} result${rows.length === 1 ? "" : "s"} for "${esc(s)}"</h2>${table(["Date", "Officer", "Region", "Number", "Result", "Reason"], rows.map(r => tr([fmtDate(r.report_date), esc(r.officer_name), esc(regionName(r.region_id)), esc(r.sim_number), r.issue ? pill(r.issue, "p-red") : pill("Tracked", "p-green"), r.reason ? esc(r.reason) + (r.note ? `<br><span class="muted">${esc(r.note)}</span>` : "") : ""])), "Nothing found.")}</div>`;
  };
  drawList();
};
