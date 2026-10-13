const C = window.GP_CONFIG;
const sb = supabase.createClient(C.url, C.key);
const $ = id => document.getElementById(id);
const esc = v => (v === null || v === undefined ? "" : String(v)).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const OFF_STATUS = ["Active", "Left", "Vacant post", "HQ / Not tracked"];
const SIM_STATUS = ["Assigned", "Assigned (HQ, not tracked)", "Spare", "Spare (inactive SIM)", "Reserved - Chakaria", "Closed - Replace", "Check"];
let S = { officers: [], sims: [], zones: [], user: null, report: null };

/* ---------------- auth ---------------- */
async function boot() {
  const { data } = await sb.auth.getSession();
  if (data.session) return enter(data.session.user);
  $("login").classList.remove("hidden");
}
$("btnLogin").onclick = async () => {
  $("loginMsg").textContent = "";
  const { data, error } = await sb.auth.signInWithPassword({ email: $("email").value.trim(), password: $("pass").value });
  if (error) return ($("loginMsg").textContent = error.message);
  enter(data.user);
};
$("pass").onkeydown = e => { if (e.key === "Enter") $("btnLogin").click(); };
$("lnkSignup").onclick = async e => {
  e.preventDefault();
  const { error } = await sb.auth.signUp({ email: $("email").value.trim(), password: $("pass").value, options: { emailRedirectTo: location.origin } });
  $("loginMsg").textContent = error ? error.message : "Account created. Check your email to confirm, then sign in. (Only approved emails can see data.)";
};
$("lnkReset").onclick = async e => {
  e.preventDefault();
  const { error } = await sb.auth.resetPasswordForEmail($("email").value.trim(), { redirectTo: location.origin });
  $("loginMsg").textContent = error ? error.message : "Password reset email sent.";
};
sb.auth.onAuthStateChange(async (ev) => {
  if (ev === "PASSWORD_RECOVERY") {
    const p = prompt("Enter a new password:");
    if (p) { const { error } = await sb.auth.updateUser({ password: p }); alert(error ? error.message : "Password updated."); }
  }
});
$("btnLogout").onclick = async () => { await sb.auth.signOut(); location.reload(); };

async function enter(user) {
  S.user = user;
  $("login").classList.add("hidden");
  $("app").classList.remove("hidden");
  $("who").textContent = user.email;
  await loadAll();
  if (!S.zones.length) {
    $("page-report").innerHTML = `<div class="card"><h2>No access</h2><p>${esc(user.email)} is signed in but not on the approved list. Ask the admin to add this email.</p></div>`;
  }
}
async function loadAll() {
  const [o, s, z] = await Promise.all([
    sb.from("officers").select("*").order("id"),
    sb.from("sims").select("*").order("number"),
    sb.from("report_zones").select("*").order("sort"),
  ]);
  S.officers = o.data || []; S.sims = s.data || []; S.zones = z.data || [];
  fillFilters(); renderOfficers(); renderSims(); renderDash();
}

/* ---------------- tabs ---------------- */
document.querySelectorAll(".tab").forEach(b => b.onclick = () => {
  document.querySelectorAll(".tab").forEach(x => x.classList.toggle("active", x === b));
  document.querySelectorAll(".page").forEach(p => p.classList.toggle("hidden", p.id !== "page-" + b.dataset.page));
  if (b.dataset.page === "dash") renderDash();
});

/* ---------------- helpers ---------------- */
async function readTracker(files, acc) {
  for (const f of files) {
    const wb = XLSX.read(await f.arrayBuffer(), { cellDates: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
    acc = GPReport.loadTracker(rows, acc);
    acc.raw = (acc.raw || []).concat(rows.slice(1));
  }
  return acc;
}
async function log(action, details) {
  await sb.from("change_log").insert({ action, details, source: "web" });
}
function download(name, blob) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function exportXlsx(name, rows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Data");
  XLSX.writeFile(wb, name);
}
const ddmmyyyy = d => `${String(d.getDate()).padStart(2, "0")}-${String(d.getMonth() + 1).padStart(2, "0")}-${d.getFullYear()}`;

/* ---------------- report ---------------- */
$("btnBuild").onclick = async () => {
  const files = [...$("ttFiles").files];
  if (!files.length) return alert("Choose the Team Tracker export file(s) first.");
  $("btnBuild").disabled = true; $("btnBuild").textContent = "Working…";
  try {
    const acc = await readTracker(files);
    const offs = S.officers.filter(o => o.report_zone && o.status === "Active");
    const R = GPReport.buildReport(offs, S.zones, acc, C);
    S.report = R;
    renderReport(R);
    const st = R.stats;
    $("repStats").innerHTML = [["Officers", st.officers], ["Found in export", st.matched_rows], ["Couldn't locate", st.cl_tracker_set], ["No raw record", st.no_raw_record], ["Absent", st.absent_set], ["Inactive?", st.inactive_flagged]]
      .map(([k, v]) => `<div class="stat">${k}<b>${v}</b></div>`).join("");
    if (R.unmatchedCl.length) {
      $("repWarn").classList.remove("hidden");
      $("repWarn").textContent = `${R.unmatchedCl.length} "Couldn't locate" name(s) in the tracker did not match any officer in the report (spelling differs, or not a report officer):\n- ` + R.unmatchedCl.join("\n- ");
    } else $("repWarn").classList.add("hidden");
    $("btnPrint").disabled = false; $("btnAudit").disabled = false;
    document.title = `GP Report ${ddmmyyyy(R.date)}`;
    const y = R.date; const iso = `${y.getFullYear()}-${String(y.getMonth() + 1).padStart(2, "0")}-${String(y.getDate()).padStart(2, "0")}`;
    await sb.from("report_runs").insert({ report_date: iso, stats: { ...st, files: files.map(f => f.name), unmatched_cl: R.unmatchedCl } });
  } catch (e) { alert("Error: " + e.message); console.error(e); }
  $("btnBuild").disabled = false; $("btnBuild").textContent = "Build report";
};
function renderReport(R) {
  const d = ddmmyyyy(R.date);
  $("reportOut").innerHTML = R.zones.map(z => `
  <div class="rep-zone">
    <div class="rep-title">${esc(z.title)}</div>
    <div class="rep-date">Date: ${d}</div>
    <table class="rep">
      <colgroup><col style="width:3%"><col style="width:14%"><col style="width:10%"><col style="width:10%"><col style="width:8%">
      <col style="width:11.25%"><col style="width:11.25%"><col style="width:11.25%"><col style="width:11.25%"><col style="width:10%"></colgroup>
      <thead><tr><th>Sl.<br>No</th><th>Name of Officer</th><th>Designation</th><th>Base Location</th><th>Mobile Number</th>
      ${C.checkpoints.map(c => `<th>${c[0]}</th>`).join("")}<th>Remarks</th></tr></thead>
      <tbody>${z.rows.map(r => `<tr><td>${r.sl}</td><td>${esc(r.name)}</td><td>${esc(r.desig)}</td><td>${esc(r.base)}</td><td>${esc(r.mobile)}</td>
        ${r.cells.map(c => `<td${c.red ? ' class="red"' : ""}>${esc(c.txt)}</td>`).join("")}<td${r.rred ? ' class="red"' : ""}>${esc(r.remarks)}</td></tr>`).join("")}</tbody>
    </table>
    <div class="rep-sum no-print">Total officers: ${z.total} · Couldn't locate (all 4 checkpoints): ${z.allCl}</div>
  </div>`).join("");
}
$("btnPrint").onclick = () => window.print();
$("btnAudit").onclick = () => {
  const rows = S.report.audit;
  const cols = Object.keys(rows[0]);
  const csv = "﻿" + [cols.join(",")].concat(rows.map(r => cols.map(c => `"${String(r[c] ?? "").replace(/"/g, '""')}"`).join(","))).join("\r\n");
  download(`Automation_Audit_${ddmmyyyy(S.report.date)}.csv`, new Blob([csv], { type: "text/csv" }));
};

/* ---------------- TT check ---------------- */
$("btnCheck").onclick = async () => {
  const files = [...$("ttCheckFiles").files];
  if (!files.length) return alert("Choose a Team Tracker export first.");
  const acc = await readTracker(files);
  const ttNums = new Map(); // number -> TT name
  for (const r of acc.raw) {
    const n = GPReport.normMobile(r[1]);
    if (n && !ttNums.has(n)) ttNums.set(n, String(r[0] || "").split("\n")[0].trim());
  }
  const byNum = new Map(S.officers.filter(o => o.tracker_number).map(o => [o.tracker_number, o]));
  const simBy = new Map(S.sims.map(s => [s.number, s]));
  const notInDb = [], wrongPerson = [];
  for (const [n, name] of ttNums) {
    const o = byNum.get(n);
    if (!o) {
      const s = simBy.get(n);
      if (/no officer assigned/i.test(name)) continue;
      notInDb.push([GPReport.fmtMobile(n), name, s ? s.status : "Not a company SIM", s?.remarks || ""]);
    } else if (o.report_name && GPReport.nameKey(name) !== GPReport.nameKey(o.report_name) && GPReport.nameKey(name) !== GPReport.nameKey(o.name)) {
      wrongPerson.push([GPReport.fmtMobile(n), name, o.report_name || o.name, o.report_zone || ""]);
    }
  }
  const clNames = new Set(Object.keys(acc.clNames).map(GPReport.nameKey));
  const missing = S.officers.filter(o => o.report_zone && o.status === "Active" && !ttNums.has(GPReport.normMobile(o.tracker_number)) && !clNames.has(GPReport.nameKey(o.report_name)))
    .map(o => [GPReport.fmtMobile(o.tracker_number), o.report_name, o.report_designation, o.report_zone]);
  const tbl = (title, head, rows, note) => `<h3>${title} (${rows.length})</h3>${note ? `<p class="muted">${note}</p>` : ""}${rows.length ? `<div class="tablewrap"><table class="grid"><tr>${head.map(h => `<th>${h}</th>`).join("")}</tr>${rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</table></div>` : "<p>None 👍</p>"}`;
  $("checkOut").innerHTML =
    tbl("Report officers not found in Team Tracker", ["Number", "Name", "Designation", "Zone"], missing, "Need to be added in Team Tracker (or number is wrong there).") +
    tbl("Tracker number registered to a different name", ["Number", "Name in TT", "Name in report", "Zone"], wrongPerson, "Check spelling or whether the SIM changed hands.") +
    tbl("Numbers in Team Tracker that are not report officers", ["Number", "Name in TT", "SIM status", "Remarks"], notInDb, "Usually HQ/spare numbers. Leavers should be set to 'No officer assigned' in TT.");
};

/* ---------------- officers ---------------- */
function fillFilters() {
  $("offStatus").innerHTML = `<option value="">All status</option>` + OFF_STATUS.map(s => `<option>${s}</option>`).join("");
  $("offZone").innerHTML = `<option value="">All report zones</option><option value="-">Not in report</option>` + S.zones.map(z => `<option value="${z.code}">${esc(z.title)}</option>`).join("");
  $("simStatus").innerHTML = `<option value="">All status</option>` + SIM_STATUS.map(s => `<option>${s}</option>`).join("");
}
["offSearch", "offStatus", "offZone"].forEach(id => $(id).oninput = renderOfficers);
function filteredOfficers() {
  const q = $("offSearch").value.toLowerCase(), st = $("offStatus").value, z = $("offZone").value;
  return S.officers.filter(o => (!st || o.status === st) && (!z || (z === "-" ? !o.report_zone : o.report_zone === z)) &&
    (!q || [o.name, o.report_name, o.corporate_sim, o.personal_number, o.tracker_number, o.territory, o.region, o.tt_name].join(" ").toLowerCase().includes(q)));
}
function renderOfficers() {
  const rows = filteredOfficers();
  $("offTable").innerHTML = `<tr><th>Name</th><th>Desig.</th><th>Region</th><th>Territory</th><th>Corporate SIM</th><th>Personal</th><th>Status</th><th>Report</th><th>TT name</th><th>Remarks</th></tr>` +
    rows.map(o => `<tr data-id="${o.id}"><td>${esc(o.name)}</td><td>${esc(o.designation)}</td><td>${esc(o.region)}</td><td>${esc(o.territory)}</td><td>${esc(GPReport.fmtMobile(o.corporate_sim))}</td><td>${esc(o.personal_number)}</td>
    <td><span class="pill ${esc((o.status || "").split(" ")[0])}">${esc(o.status)}</span></td><td>${esc(o.report_zone ? o.report_zone + " #" + (o.report_order - 5) : "")}</td><td>${esc(o.tt_name)}</td><td>${esc(o.remarks)}</td></tr>`).join("");
  $("offTable").querySelectorAll("tr[data-id]").forEach(tr => tr.onclick = () => editOfficer(S.officers.find(o => o.id == tr.dataset.id)));
}
$("btnAddOff").onclick = () => editOfficer({ status: "Active" });
$("btnExportOff").onclick = () => exportXlsx("Officers.xlsx", filteredOfficers());

const OFF_FIELDS = [["name", "Name"], ["designation", "Designation"], ["region", "Region"], ["territory", "Territory"], ["zone", "Zone"], ["status", "Status", OFF_STATUS],
  ["corporate_sim", "Corporate SIM (10 digits, e.g. 1709661586)"], ["personal_number", "Personal number"], ["tracker_number", "Number used in Team Tracker"], ["tt_name", "Name in Team Tracker"],
  ["report_zone", "Report zone", null], ["report_order", "Report row order (number)"], ["report_name", "Name on report"], ["report_designation", "Designation on report"], ["report_base", "Base location on report"], ["remarks", "Remarks", "full"]];

function editOfficer(o) {
  const f = $("dlgForm");
  f.innerHTML = `<h2 class="full">${o.id ? "Edit officer" : "Add officer"}</h2>` + OFF_FIELDS.map(([k, label, opt]) => {
    const v = o[k] ?? "";
    if (k === "report_zone") return `<label>${label}<select name="${k}"><option value="">— not in report —</option>${S.zones.map(z => `<option value="${z.code}" ${v === z.code ? "selected" : ""}>${esc(z.title)}</option>`).join("")}</select></label>`;
    if (Array.isArray(opt)) return `<label>${label}<select name="${k}">${opt.map(s => `<option ${v === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>`;
    return `<label class="${opt === "full" ? "full" : ""}">${label}<input name="${k}" value="${esc(v)}"></label>`;
  }).join("") + `<div class="actions"><button value="cancel">Cancel</button><button value="save" class="primary">Save</button></div>`;
  $("dlg").showModal();
  $("dlg").onclose = async () => {
    if ($("dlg").returnValue !== "save") return;
    const rec = {};
    for (const [k] of OFF_FIELDS) { const v = f.elements[k].value.trim(); rec[k] = v === "" ? null : v; }
    for (const k of ["corporate_sim", "tracker_number"]) if (rec[k]) rec[k] = GPReport.normMobile(rec[k]);
    if (rec.report_order) rec.report_order = parseInt(rec.report_order, 10);
    if (rec.report_zone && !rec.report_order) {
      const max = Math.max(5, ...S.officers.filter(x => x.report_zone === rec.report_zone).map(x => x.report_order || 0));
      rec.report_order = max + 1;
    }
    if (rec.report_zone && !rec.report_name) rec.report_name = rec.name;
    if (rec.report_zone && !rec.tracker_number) rec.tracker_number = rec.corporate_sim;
    const changed = o.id ? Object.keys(rec).filter(k => (o[k] ?? null) != rec[k]).map(k => `${k}: ${o[k] ?? "—"} → ${rec[k] ?? "—"}`).join("; ") : "";
    const q = o.id ? sb.from("officers").update(rec).eq("id", o.id) : sb.from("officers").insert(rec);
    const { error } = await q;
    if (error) return alert(error.message);
    await log(o.id ? "officer_edit" : "officer_add", `${rec.name || o.name}: ${changed || "new record"}`);
    await loadAll();
  };
}

/* ---------------- sims ---------------- */
["simSearch", "simStatus"].forEach(id => $(id).oninput = renderSims);
function filteredSims() {
  const q = $("simSearch").value.toLowerCase(), st = $("simStatus").value;
  const offBy = new Map(S.officers.map(o => [o.id, o]));
  return S.sims.map(s => ({ ...s, officer: offBy.get(s.officer_id)?.name || "" }))
    .filter(s => (!st || s.status === st) && (!q || [s.number, s.remarks, s.officer, s.zone].join(" ").toLowerCase().includes(q)));
}
function renderSims() {
  $("simTable").innerHTML = `<tr><th>Number</th><th>Series</th><th>Status</th><th>Officer</th><th>Zone</th><th>Remarks</th></tr>` +
    filteredSims().map(s => `<tr data-n="${s.number}"><td>${GPReport.fmtMobile(s.number)}</td><td>${esc(s.series)}</td><td><span class="pill ${esc(s.status.split(" ")[0])}">${esc(s.status)}</span></td><td>${esc(s.officer)}</td><td>${esc(s.zone)}</td><td>${esc(s.remarks)}</td></tr>`).join("");
  $("simTable").querySelectorAll("tr[data-n]").forEach(tr => tr.onclick = () => editSim(S.sims.find(s => s.number === tr.dataset.n)));
}
$("btnExportSim").onclick = () => exportXlsx("SIMs.xlsx", filteredSims());
function editSim(s) {
  const f = $("dlgForm");
  const offs = S.officers.filter(o => o.status !== "Left").sort((a, b) => (a.name || "").localeCompare(b.name || ""));
  f.innerHTML = `<h2 class="full">SIM ${GPReport.fmtMobile(s.number)}</h2>
   <label>Status<select name="status">${SIM_STATUS.map(x => `<option ${x === s.status ? "selected" : ""}>${x}</option>`).join("")}</select></label>
   <label>Zone<input name="zone" value="${esc(s.zone)}"></label>
   <label class="full">Officer<select name="officer_id"><option value="">— none —</option>${offs.map(o => `<option value="${o.id}" ${o.id === s.officer_id ? "selected" : ""}>${esc(o.name)} · ${esc(o.designation)} · ${esc(o.territory)}</option>`).join("")}</select></label>
   <label class="full">Remarks<input name="remarks" value="${esc(s.remarks)}"></label>
   <div class="actions"><button value="cancel">Cancel</button><button value="save" class="primary">Save</button></div>`;
  $("dlg").showModal();
  $("dlg").onclose = async () => {
    if ($("dlg").returnValue !== "save") return;
    const rec = { status: f.elements.status.value, zone: f.elements.zone.value || null, officer_id: f.elements.officer_id.value ? +f.elements.officer_id.value : null, remarks: f.elements.remarks.value || null };
    const { error } = await sb.from("sims").update(rec).eq("number", s.number);
    if (error) return alert(error.message);
    await log("sim_edit", `${s.number}: ${s.status} → ${rec.status}; officer ${s.officer_id ?? "—"} → ${rec.officer_id ?? "—"}; ${rec.remarks || ""}`);
    await loadAll();
  };
}

/* ---------------- dashboard ---------------- */
async function renderDash() {
  const cnt = (arr, f) => arr.filter(f).length;
  const tiles = [
    ["Active officers", cnt(S.officers, o => o.status === "Active")],
    ["On daily report", cnt(S.officers, o => o.report_zone && o.status === "Active")],
    ["Left", cnt(S.officers, o => o.status === "Left")],
    ["Vacant posts", cnt(S.officers, o => o.status === "Vacant post")],
    ["Total SIMs", S.sims.length],
    ["Spare (active)", cnt(S.sims, s => s.status === "Spare")],
    ["Spare (inactive)", cnt(S.sims, s => s.status === "Spare (inactive SIM)")],
    ["Reserved Chakaria", cnt(S.sims, s => s.status === "Reserved - Chakaria")],
    ["Need replacement", cnt(S.sims, s => s.status === "Closed - Replace")],
    ["To check", cnt(S.sims, s => s.status === "Check")],
  ];
  const { data: runs } = await sb.from("report_runs").select("*").order("created_at", { ascending: false }).limit(10);
  $("dashOut").innerHTML = `<div class="stats">${tiles.map(([k, v]) => `<div class="stat">${k}<b>${v}</b></div>`).join("")}</div>
   <h3>Recent reports</h3><table class="grid"><tr><th>Date</th><th>By</th><th>Found</th><th>Couldn't locate</th><th>No raw record</th><th>Absent</th></tr>
   ${(runs || []).map(r => `<tr><td>${r.report_date}</td><td>${esc(r.created_by)}</td><td>${r.stats?.matched_rows ?? ""}</td><td>${r.stats?.cl_tracker_set ?? ""}</td><td>${r.stats?.no_raw_record ?? ""}</td><td>${r.stats?.absent_set ?? ""}</td></tr>`).join("")}</table>`;
  const { data: logs } = await sb.from("change_log").select("*").order("at", { ascending: false }).limit(30);
  $("logOut").innerHTML = `<table class="grid"><tr><th>When</th><th>Who</th><th>Action</th><th>Details</th></tr>${(logs || []).map(l => `<tr><td>${new Date(l.at).toLocaleString()}</td><td>${esc(l.user_email)}</td><td>${esc(l.action)}</td><td>${esc(l.details)}</td></tr>`).join("")}</table>`;
}

boot();
