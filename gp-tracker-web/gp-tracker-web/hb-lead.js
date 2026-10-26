/* HB-Tracker region lead pages (database rules limit every query to the lead's own region) */
const L = HB.pages;
const myRegion = () => HB.me.region_id;
const myOfficers = () => q(sb.from("hb_officers").select("*").eq("region_id", myRegion()).eq("status", "active").order("name"));
const send = (kind, payload) => q(sb.from("hb_requests").insert({ kind, region_id: myRegion(), payload }));

L.home = async el => {
  const [offs, sims, reqs, rep, anns] = await Promise.all([myOfficers(), q(sb.from("hb_sims").select("number,status").eq("region_id", myRegion())), q(sb.from("hb_requests").select("id").eq("status", "pending")), latestReport(), q(sb.from("hb_announcements").select("*").order("created_at", { ascending: false }).limit(3))]);
  let todo = 0; if (rep) { const { count } = await sb.from("hb_report_rows").select("id", { count: "exact", head: true }).eq("report_id", rep.id).not("issue", "is", null).is("justified_at", null); todo = count || 0; }
  el.innerHTML = head(`${esc(regionName(myRegion()))} region`, `Hello ${esc(HB.me.full_name || "")}, here is what needs you today`) +
    stats([["My officers", offs.length, "active"], ["To justify", todo, rep ? "report " + fmtDate(rep.report_date) : "no report yet", "#C2410C"], ["My spare SIMs", sims.filter(s => s.status === "spare").length, "ready for new joiners", "#15803D"], ["My requests", reqs.length, "waiting for HQ", "#1F5FD1"]]) +
    `<div class="tiles"><a class="tile" href="#/ljustify"><b>Justify today's list</b><span>${todo} officers need a reason</span></a><a class="tile" href="#/change?k=resignation"><b>Officer resigned</b><span>Report it today and say if the SIM came back</span></a><a class="tile" href="#/change?k=new_joiner"><b>New joiner</b><span>Add them and give a spare SIM</span></a><a class="tile" href="#/change?k=transfer"><b>Transfer or promotion</b><span>Move or update an officer</span></a></div>` +
    (!offs.length ? `<div class="note"><b>First time here?</b> Upload your officer list so HQ can verify it. <button class="sm primary" id="upl">Upload my list</button> <span class="muted">Excel columns: Name, Designation, Territory, Corporate SIM, Personal number</span><input type="file" id="uplF" accept=".xlsx,.xls" class="hidden"></div>` : "") +
    (anns.length ? `<div class="card"><h2>Announcements from HQ</h2>${anns.map(a => `<div style="border-left:3px solid var(--blue);padding-left:12px;margin-bottom:10px"><b>${esc(a.title)}</b><div>${esc(a.body)}</div><div class="muted" style="font-size:12px">${fmtTime(a.created_at)}</div></div>`).join("")}</div>` : "");
  const u = $("#upl"); if (u) { u.onclick = () => $("#uplF").click(); $("#uplF").onchange = async () => { try { const list = parseOfficerSheet(sheetObjects(await readExcel($("#uplF").files[0]))); if (!list.length) throw new Error("No rows with a Name column found."); if (!confirm(`Send ${list.length} officers to HQ for verification?`)) return; await send("submit_list", { officers: list }); toast("Sent to HQ"); reload(); } catch (e) { fail(e); } }; }
};

const REASONS = { "Couldn't locate": ["Phone off / no network", "Field visit (no signal)", "Leave", "Meeting at office", "Phone problem", "Other"], Absent: ["Leave", "Sick", "Field visit (no signal)", "Phone off", "Other"], "No raw record": ["Number not in Team Tracker", "SIM problem", "Leave", "Other"] };
L.ljustify = async el => {
  const rep = await latestReport();
  if (!rep) { el.innerHTML = head("Justify Today", "") + `<div class="card empty">No report yet.</div>`; return; }
  const rows = await q(sb.from("hb_report_rows").select("*").eq("report_id", rep.id).not("issue", "is", null).order("officer_name"));
  const dl = (await q(sb.from("hb_settings").select("value").eq("key", "justify_deadline")))[0]?.value || "18:00";
  el.innerHTML = head("Justify the list", `Report of ${fmtDate(rep.report_date)} · please answer by ${esc(dl)}`, rows.length ? `<button class="primary" id="jSave">Submit all</button>` : "") +
    `<div class="card">${table(["Officer", "Number", "Issue", "Reason", "Note"], rows.map(r => tr([esc(r.officer_name), esc(r.sim_number), pill(r.issue, "p-red"),
      `<select data-r="${r.id}"><option value="">Choose reason…</option>${(REASONS[r.issue] || REASONS.Absent).map(x => `<option ${r.reason === x ? "selected" : ""}>${x}</option>`).join("")}</select>`,
      `<input data-n="${r.id}" value="${esc(r.note)}" placeholder="Add a note">` + (r.justified_at ? `<div class="muted" style="font-size:12px">sent ${fmtTime(r.justified_at)}${r.accepted === true ? " · accepted" : r.accepted === false ? " · HQ not satisfied" : ""}</div>` : "")])), "Nothing to justify. All your officers were tracked.")}</div><div class="warn">Late answers count against your region's monthly score.</div>`;
  const b = $("#jSave"); if (b) b.onclick = async () => {
    const todo = rows.filter(r => el.querySelector(`[data-r="${r.id}"]`).value);
    if (!todo.length) return toast("Choose a reason first");
    b.disabled = true;
    try { for (const r of todo) await rpc("hb_justify", { p_row: r.id, p_reason: el.querySelector(`[data-r="${r.id}"]`).value, p_note: el.querySelector(`[data-n="${r.id}"]`).value || null }); toast(`${todo.length} sent to HQ`); refreshBell(); reload(); } catch (e) { fail(e); b.disabled = false; }
  };
};

L.lofficers = async el => {
  const [offs, sims] = await Promise.all([myOfficers(), q(sb.from("hb_requests").select("*").eq("status", "pending"))]);
  el.innerHTML = head("My Officers", `${esc(regionName(myRegion()))} only`, `<a href="#/change"><button class="primary">Report a change</button></a>`) + `<div class="filters"><input id="lq" placeholder="Search my officers"></div><div class="card" id="lo"></div>`;
  const draw = () => { const s = $("#lq").value.toLowerCase(); const l = offs.filter(o => !s || [o.name, o.sim_number, o.territory].join(" ").toLowerCase().includes(s));
    $("#lo").innerHTML = `<h2>${l.length} officers</h2>` + table(["Name", "Designation", "Territory", "Corporate SIM", "Personal", "Status"], l.map(o => { const pend = sims.find(r => r.payload.officer_id === o.id); return tr([esc(o.name), esc(o.designation), esc(o.territory), esc(o.sim_number || "—"), esc(o.personal_number), pend ? pill(REQ[pend.kind] + " pending", "p-amber") : o.verified ? pill("Active", "p-blue") : pill("Waiting for HQ check", "p-amber")]); }), "No officers yet. Upload your list from My Region."); };
  $("#lq").oninput = draw; draw();
};

L.spares = async el => {
  const sims = await q(sb.from("hb_sims").select("*").eq("region_id", myRegion()).in("status", ["spare", "spare_inactive", "suspended"]).order("number"));
  el.innerHTML = head("My Spare SIMs", "SIMs returned by officers, kept for new joiners") +
    `<div class="card">${table(["Number", "Note", "Since", "Status", ""], sims.map(s => tr([s.number, esc(s.note), fmtDate(s.updated_at), simPill(s), s.status === "spare" ? `<a href="#/change?k=new_joiner&sim=${s.number}"><button class="sm primary">Give to new joiner</button></a>` : ""])), "No spare SIMs.")}</div><div class="note">Need a SIM and have no spare? <a href="#/change?k=need_sim">Request one from HQ</a>.</div>`;
};

const KINDS = [["resignation", "Resignation"], ["new_joiner", "New joiner"], ["transfer", "Transfer"], ["promotion", "Promotion"], ["edit_officer", "Fix officer details"], ["need_sim", "Need a SIM from HQ"]];
L.change = async el => {
  const qs = new URLSearchParams((location.hash.split("?")[1] || ""));
  const k = qs.get("k") || S.chKind || "resignation"; S.chKind = k;
  const [offs, sims] = await Promise.all([myOfficers(), q(sb.from("hb_sims").select("number").eq("region_id", myRegion()).eq("status", "spare"))]);
  const offSel = `<label>Officer<select name="o" required><option value="">Choose from my officers…</option>${offs.map(o => `<option value="${o.id}">${esc(o.name)} · ${esc(o.sim_number || "")}</option>`).join("")}</select></label>`;
  const forms = {
    resignation: `<div class="fgrid">${offSel}<label>Last working day<input type="date" name="d" required></label></div><b>Did the officer give the corporate SIM back to you?</b><div class="choices" style="margin-top:8px"><label class="radio"><input type="radio" name="ret" value="1" required><span><b>Yes, I have the SIM</b><br><span class="muted">It becomes my spare SIM for the next new joiner.</span></span></label><label class="radio"><input type="radio" name="ret" value="0"><span><b>No, not returned</b><br><span class="muted">The SIM goes to the HQ pool. HQ decides what to do.</span></span></label></div>`,
    new_joiner: `<div class="fgrid"><label>Full name<input name="n" required></label><label>Designation<input name="g" required></label><label>Territory<input name="t"></label><label>Personal number<input name="p"></label><label>Joining date<input type="date" name="d"></label><label>Corporate SIM<select name="s"><option value="">I have no spare (HQ will give one)</option>${sims.map(s => `<option ${qs.get("sim") === s.number ? "selected" : ""}>${s.number}</option>`).join("")}</select></label></div>`,
    transfer: `<div class="fgrid">${offSel}<label>Moving to region<input name="to" required placeholder="Region name"></label></div><p class="muted">HQ moves the officer and their SIM. You will not see the other region.</p>`,
    promotion: `<div class="fgrid">${offSel}<label>New designation<input name="g" required></label></div><p class="muted">The number stays the same.</p>`,
    edit_officer: `<div class="fgrid">${offSel}<label>Correct name (if wrong)<input name="n"></label><label>Territory<input name="t"></label><label>Personal number<input name="p"></label></div>`,
    need_sim: `<label>Why do you need a SIM?<input name="w" required placeholder="e.g. new joiner [name] starting next week"></label>`
  };
  el.innerHTML = head("Report a Change", "HQ approves every change before it becomes final") +
    `<div class="choices">${KINDS.map(([x, l]) => `<button data-k="${x}" class="${x === k ? "on" : ""}">${l}</button>`).join("")}</div>` +
    `<form class="card" id="chF"><h2>${KINDS.find(x => x[0] === k)[1]}</h2>${forms[k]}<label style="margin-top:12px">Note for HQ (optional)<input name="note"></label><div class="acts" style="justify-content:flex-end;margin-top:14px"><button class="primary">Send to HQ for approval</button></div></form>`;
  el.querySelectorAll("[data-k]").forEach(b => b.onclick = () => { S.chKind = b.dataset.k; location.hash = "#/change"; reload(); });
  $("#chF").onsubmit = async e => {
    e.preventDefault(); const f = e.target; const o = f.o ? offs.find(x => x.id === +f.o.value) : null; let p;
    if (k === "resignation") p = { officer_id: o.id, officer_name: o.name, sim_number: o.sim_number, last_day: f.d.value, sim_returned: f.ret.value === "1" };
    if (k === "new_joiner") p = { name: f.n.value.trim(), designation: f.g.value.trim(), territory: f.t.value.trim(), personal_number: normNum(f.p.value) || f.p.value.trim(), joined_on: f.d.value, sim_number: f.s.value || null };
    if (k === "transfer") p = { officer_id: o.id, officer_name: o.name, to_region_name: f.to.value.trim() };
    if (k === "promotion") p = { officer_id: o.id, officer_name: o.name, old_designation: o.designation, designation: f.g.value.trim() };
    if (k === "edit_officer") p = { officer_id: o.id, officer_name: o.name, name: f.n.value.trim() || null, territory: f.t.value.trim() || null, personal_number: f.p.value.trim() || null };
    if (k === "need_sim") p = { reason: f.w.value.trim() };
    p.note = f.note.value.trim() || null;
    try { await send(k, p); toast("Sent to HQ for approval"); location.hash = "#/requests"; } catch (er) { fail(er); }
  };
};

L.requests = async el => {
  const reqs = await q(sb.from("hb_requests").select("*").order("created_at", { ascending: false }).limit(200));
  el.innerHTML = head("My Requests", "What you sent to HQ") + `<div class="card">${table(["Request", "Details", "Sent", "Status", "HQ note"], reqs.map(r => tr([esc(REQ[r.kind]), esc(reqSummary(r)), fmtTime(r.created_at), reqPill(r.status), esc(r.decision_note)])), "No requests yet.")}</div>`;
};

L.lmessages = async el => {
  const [msgs, anns] = await Promise.all([q(sb.from("hb_messages").select("*").eq("lead_id", HB.me.user_id).order("created_at")), q(sb.from("hb_announcements").select("*").order("created_at", { ascending: false }).limit(10))]);
  el.innerHTML = head("Message HQ", "Your private conversation with HQ") + `<div class="chat"><div id="thread" class="card thread"></div>${anns.length ? `<div class="card" style="flex:0 1 300px"><h2>Announcements</h2>${anns.map(a => `<div style="border-left:3px solid var(--blue);padding-left:12px;margin-bottom:10px"><b>${esc(a.title)}</b><div>${esc(a.body)}</div><div class="muted" style="font-size:12px">${fmtDate(a.created_at)}</div></div>`).join("")}</div>` : ""}</div>`;
  drawThread($("#thread"), HB.me.user_id, "HQ", "Only you and HQ can see these messages", msgs);
};
