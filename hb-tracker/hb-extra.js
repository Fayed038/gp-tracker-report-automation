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
