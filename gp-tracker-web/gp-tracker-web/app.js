const C = window.GP_CONFIG;
async function relayFetch(input, init = {}) {
  const u = new URL(typeof input === "string" ? input : input.url);
  const h = {};
  new Headers(init.headers || {}).forEach((v, k) => (h[k] = v));
  const res = await fetch(C.relay, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ path: u.pathname + u.search, method: init.method || "GET", headers: h, body: init.body ?? null }),
    signal: init.signal,
  });
  const r = await res.json();
  const noBody = r.status === 204 || r.status === 205 || r.status === 304;
  return new Response(noBody ? null : r.body, { status: r.status, headers: r.headers || {} });
}
const useRelay = C.relay && C.relay.startsWith("https://");
const sb = supabase.createClient(C.url, C.key, useRelay ? { global: { fetch: relayFetch } } : {});
const $ = id => document.getElementById(id);
const esc = v => (v === null || v === undefined ? "" : String(v)).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const OFF_STATUS = ["Active", "Left", "Vacant post", "HQ / Not tracked", "Suspended"];
const SIM_STATUS = ["Assigned", "Assigned (HQ, not tracked)", "Spare", "Spare (inactive SIM)", "Reserved - Chakaria", "Closed - Replace", "Check", "Suspended"];
let S = { officers: [], sims: [], zones: [], user: null, report: null };

/* ---------------- auth ---------------- */
async function boot() {
  const { data } = await sb.auth.getSession();
  if (data.session) return enter(data.session.user);
  $("login").classList.remove("hidden");
}
$("btnLogin").onclick = async () => {
  $("loginMsg").textContent = "Signing in…";
  const { data, error } = await sb.auth.signInWithPassword({ email: $("email").value.trim(), password: $("pass").value });
  if (error) return ($("loginMsg").textContent = error.message);
  enter(data.user);
};
$("pass").onkeydown = e => { if (e.key === "Enter") $("btnLogin").click(); };
$("lnkSignup").onclick = async e => {
  e.preventDefault();
  if (!$("email").value.trim() || $("pass").value.length < 6) return ($("loginMsg").textContent = "Enter your email and a password (at least 6 characters) first, then click Create account.");
  const { error } = await sb.auth.signUp({ email: $("email").value.trim(), password: $("pass").value, options: { emailRedirectTo: location.origin + location.pathname } });
  $("loginMsg").textContent = error ? error.message : "Account created. Check your email to confirm, then sign in. (Only approved emails can see data.)";
};
$("lnkReset").onclick = async e => {
  e.preventDefault();
  const { error } = await sb.auth.resetPasswordForEmail($("email").value.trim(), { redirectTo: location.origin + location.pathname });
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
  const { data: me } = await sb.from("allowed_users").select("*").eq("email", user.email).maybeSingle();
  S.me = me;
  if (!me) {
    document.querySelector("header nav").innerHTML = "";
    document.querySelectorAll(".page").forEach(p => p.classList.add("hidden"));
    $("page-report").classList.remove("hidden");
    $("page-report").innerHTML = `<div class="card"><h2>No access yet</h2><p>${esc(user.email)} is signed in but not approved. Ask the admin (Fayed) to add this email.</p></div>`;
    $("who").textContent = user.email; return;
  }
  $("who").textContent = `${user.email} · ${me.role === "admin" ? "Admin" : "Lead · " + (me.region || "")}`;
  if (me.role !== "admin") {
    ["report", "ttcheck", "dash"].forEach(p => document.querySelector(`.tab[data-page=${p}]`).remove());
    document.querySelector(".tab[data-page=requests]").textContent = "My requests";
    document.querySelector(".tab[data-page=changes]").click();
  }
  await loadAll();
  renderRequests();
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
  if (b.dataset.page === "requests") renderRequests();
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
    const iss = ttIssues(acc);
    const clNote = R.unmatchedCl.length ? `<h3>“Couldn't locate” names that match no report officer (${R.unmatchedCl.length})</h3><p class="muted">Spelling differs, or the person is not in the report.</p><p>${R.unmatchedCl.map(esc).join("<br>")}</p>` : "";
    const n = iss.count + R.unmatchedCl.length;
    $("repWarn").classList.remove("hidden");
    $("repWarn").innerHTML = (n ? `<b>⚠ ${n} thing(s) to check</b> — the report is still correct for everyone else.` : "<b>✓ No problems found</b> between Team Tracker and our officer list.") + `<details ${n ? "open" : ""}><summary>Details</summary>${clNote}${iss.html}</details>`;
    ["btnPrint", "btnAudit", "btnCsv"].forEach(id => $(id).disabled = false);
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
$("btnPrint").onclick = async () => {
  const b = $("btnPrint"); b.disabled = true; b.textContent = "Making PDF…";
  const box = $("reportOut"); box.classList.add("pdfmode");
  try {
    await html2pdf().set({
      margin: [8, 6, 8, 6], filename: `GP_Report_${ddmmyyyy(S.report.date)}.pdf`,
      image: { type: "jpeg", quality: 0.85 }, html2canvas: { scale: 1.8, backgroundColor: "#ffffff" },
      jsPDF: { unit: "mm", format: "a4", orientation: "landscape" },
      pagebreak: { mode: ["css", "avoid-all"], before: ".rep-zone + .rep-zone", avoid: "tr" },
    }).from($("reportOut")).save();
  } catch (e) { box.classList.remove("pdfmode"); alert("PDF failed: " + e.message + "\nUse Ctrl+P → Save as PDF instead."); }
  box.classList.remove("pdfmode"); b.disabled = false; b.textContent = "Download PDF";
};
$("btnCsv").onclick = () => {
  const R = S.report, d = ddmmyyyy(R.date), q = v => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const head = ["Sl. No", "Name of Officer", "Designation", "Base Location", "Mobile Number", ...C.checkpoints.map(c => c[0]), "Remarks"];
  const lines = [];
  for (const z of R.zones) {
    lines.push(q(z.title), q("Date: " + d), head.map(q).join(","));
    z.rows.forEach(r => lines.push([r.sl, r.name, r.desig, r.base, r.mobile, ...r.cells.map(c => c.txt), r.remarks].map(q).join(",")));
    lines.push([q(""), q("TOTAL Officers"), z.total].join(","), [q(""), q("Couldn't Locate (all 4 checkpoints)"), z.allCl].join(","), "");
  }
  download(`GP_Report_${d}.csv`, new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv" }));
};
$("btnAudit").onclick = () => {
  const rows = S.report.audit;
  const cols = Object.keys(rows[0]);
  const csv = "﻿" + [cols.join(",")].concat(rows.map(r => cols.map(c => `"${String(r[c] ?? "").replace(/"/g, '""')}"`).join(","))).join("\r\n");
  download(`Automation_Audit_${ddmmyyyy(S.report.date)}.csv`, new Blob([csv], { type: "text/csv" }));
};

/* ---------------- TT check ---------------- */
function ttIssues(acc) {
  const ttNums = new Map(); // number -> TT name
  for (const r of acc.raw) {
    const n = GPReport.normMobile(r[1]);
    if (n && !ttNums.has(n)) ttNums.set(n, String(r[0] || "").split("\n")[0].trim());
  }
  const byNum = new Map(S.officers.filter(o => o.tracker_number && o.status !== "Left").map(o => [o.tracker_number, o]));
  const simBy = new Map(S.sims.map(s => [s.number, s]));
  const notInDb = [], wrongPerson = [], noaButAssigned = [];
  for (const [n, name] of ttNums) {
    const o = byNum.get(n);
    const noa = /no officer assigned/i.test(name);
    if (noa) {
      if (o && o.status === "Active" && o.report_zone) noaButAssigned.push([fmN(n), o.report_name || o.name, o.designation, o.report_zone]);
      continue;
    }
    if (!o) {
      const s = simBy.get(n);
      if (s && s.status.startsWith("Assigned (HQ")) continue; // HQ numbers are fine in TT
      notInDb.push([fmN(n), name, s ? s.status : "Not in our SIM list", s?.remarks || ""]);
    } else if (o.status === "Active" && o.report_name && ![o.report_name, o.name, o.tt_name].filter(Boolean).some(x => GPReport.nameKey(x) === GPReport.nameKey(name))) {
      wrongPerson.push([fmN(n), name, o.report_name || o.name, o.report_zone || ""]);
    }
  }
  const clNames = new Set(Object.keys(acc.clNames).map(GPReport.nameKey));
  const missing = S.officers.filter(o => o.report_zone && o.status === "Active" && !ttNums.has(GPReport.normMobile(o.tracker_number)) && ![o.report_name, o.tt_name].filter(Boolean).some(x => clNames.has(GPReport.nameKey(x))))
    .map(o => [fmN(o.tracker_number), o.report_name, o.report_designation, o.report_zone]);
  const tbl = (title, head, rows, note) => `<h3>${title} (${rows.length})</h3>${note ? `<p class="muted">${note}</p>` : ""}${rows.length ? `<div class="tablewrap"><table class="grid"><tr>${head.map(h => `<th>${h}</th>`).join("")}</tr>${rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</table></div>` : "<p>None 👍</p>"}`;
  const html =
    tbl("Report officers not found in Team Tracker", ["Number", "Name", "Designation", "Zone"], missing, "Add them in Team Tracker, or the number there is wrong.") +
    tbl("Team Tracker says “No officer assigned” but our list has an active officer", ["Number", "Officer in our list", "Designation", "Zone"], noaButAssigned, "Either register the officer in Team Tracker, or mark them Left here (Changes → Officer left).") +
    tbl("Tracker number registered to a different name", ["Number", "Name in TT", "Name in report", "Zone"], wrongPerson, "Check the spelling, or whether the SIM changed hands.") +
    tbl("Numbers in Team Tracker that are not report officers", ["Number", "Name in TT", "Our SIM status", "Remarks"], notInDb, "Spare/left numbers should be “No officer assigned” in TT. Numbers not in our SIM list need checking.");
  return { html, count: missing.length + noaButAssigned.length + wrongPerson.length + notInDb.length };
}
const fmN = n => GPReport.fmtMobile(n);
$("btnCheck").onclick = async () => {
  const files = [...$("ttCheckFiles").files];
  if (!files.length) return alert("Choose a Team Tracker export first.");
  const acc = await readTracker(files);
  $("checkOut").innerHTML = ttIssues(acc).html;
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
    rows.map(o => `<tr data-id="${o.id}"><td>${isHQ(o) ? "🔒 " : isLeaf(o) ? "🛡 " : ""}${esc(o.name)}</td><td>${esc(o.designation)}</td><td>${esc(o.region)}</td><td>${esc(o.territory)}</td><td>${esc(GPReport.fmtMobile(o.corporate_sim))}</td><td>${esc(o.personal_number)}</td>
    <td><span class="pill ${esc((o.status || "").split(" ")[0])}">${esc(o.status)}</span></td><td>${esc(o.report_zone ? o.report_zone + " #" + (o.report_order - 5) : "")}</td><td>${esc(o.tt_name)}</td><td>${esc(o.remarks)}</td></tr>`).join("");
  $("offTable").querySelectorAll("tr[data-id]").forEach(tr => tr.onclick = () => editOfficer(S.officers.find(o => o.id == tr.dataset.id)));
}
$("btnAddOff").onclick = () => editOfficer({ status: "Active" });
$("btnExportOff").onclick = () => exportXlsx("Officers.xlsx", filteredOfficers());

const OFF_FIELDS = [["name", "Name"], ["designation", "Designation"], ["region", "Region"], ["territory", "Territory"], ["zone", "Zone"], ["status", "Status", OFF_STATUS],
  ["corporate_sim", "Corporate SIM (10 digits, e.g. 1709661586)"], ["personal_number", "Personal number"], ["tracker_number", "Number used in Team Tracker"], ["tt_name", "Name in Team Tracker"],
  ["report_zone", "Report zone", null], ["report_order", "Report row order (number)"], ["report_name", "Name on report"], ["report_designation", "Designation on report"], ["report_base", "Base location on report"], ["remarks", "Remarks", "full"]];

function editOfficer(o) {
  if (!isAdmin()) return alert("To change anything, use the Changes tab. Your request goes to the admin for approval.");
  if (isHQ(o)) return alert(`🔒 ${o.name} is an HQ number (${fmN(o.corporate_sim)}).\nHQ numbers are locked and cannot be changed from the website.`);
  if (isLeaf(o) && o.id && !confirm(`🛡 ${o.name} is a Leaf officer. Number/status changes are only allowed for admins and are double-checked.\nContinue?`)) return;
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
  if (!isAdmin()) return alert("To change a SIM, use the Changes tab. Your request goes to the admin for approval.");
  if (s.status === "Assigned (HQ, not tracked)") return alert(`🔒 ${fmN(s.number)} is an HQ number and is locked.`);
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


/* ---------------- guided changes (admin applies directly, leads send for approval) ---------------- */
const today = () => ddmmyyyy(new Date());
const fm = n => GPReport.fmtMobile(n);
const isAdmin = () => S.me && S.me.role === "admin";
function offOptions(filter = o => o.status === "Active") {
  return S.officers.filter(filter).sort((a, b) => (a.name || "").localeCompare(b.name || ""))
    .map(o => `<option value="${o.id}">${esc(o.name)} · ${esc(o.designation)} · ${esc(o.territory || o.region)} · ${esc(fm(o.corporate_sim || o.tracker_number))}</option>`).join("");
}
function simChooser(name) {
  const free = S.sims.filter(s => ["Spare", "Reserved - Chakaria", "Spare (inactive SIM)"].includes(s.status));
  return `<label>SIM to give<select name="${name}"><option value="">— type a number below instead —</option>${free.map(s => `<option value="${s.number}">${fm(s.number)} · ${esc(s.status)}${s.zone ? " · " + esc(s.zone) : ""}</option>`).join("")}</select></label>
  <label>…or new SIM number<input name="${name}_new" placeholder="01XXX-XXXXXX"></label>`;
}
const zoneSelect = (name, v) => `<label>Show in daily report<select name="${name}"><option value="">— not in report —</option>${S.zones.map(z => `<option value="${z.code}" ${z.code === v ? "selected" : ""}>${esc(z.title)}</option>`).join("")}</select></label>`;
const inp = (name, label, v = "") => `<label>${label}<input name="${name}" value="${esc(v)}"></label>`;
function pickedSim(f, name) { return GPReport.normMobile(f.elements[name + "_new"].value) || f.elements[name].value || null; }
const offById = id => S.officers.find(o => o.id == id);
async function upsertSim(number, fields) {
  const exists = S.sims.find(s => s.number === number);
  const series = number.startsWith("13") ? "013 Series" : number.startsWith("17") ? "017 Series" : "Other";
  const q = exists ? sb.from("sims").update(fields).eq("number", number) : sb.from("sims").insert({ number, series, ...fields });
  const { error } = await q; if (error) throw error;
}
async function freeSim(number, returned, who) {
  if (!number) return "";
  await upsertSim(number, returned
    ? { status: "Spare", officer_id: null, remarks: `Freed from ${who} (${today()}). Verify SIM in hand before giving out` }
    : { status: "Closed - Replace", officer_id: null, remarks: `${who} left with SIM, not returned (${today()}). Needs replacement` });
  return returned ? `SIM ${fm(number)} is now Spare.` : `SIM ${fm(number)} marked "Closed - Replace" (SIM not returned).`;
}
function nextOrder(zone) { return Math.max(5, ...S.officers.filter(o => o.report_zone === zone && o.status === "Active").map(o => o.report_order || 0)) + 1; }
const must = (c, m) => { if (!c) throw new Error(m); };
const chk = r => { if (r && r.error) throw r.error; return r; };

/* Each action: form(), collect(f) -> {data, summary}, apply(data) -> {msg, tt} */
const ACTIONS = {
  join: {
    title: "New officer joined", btn: "Add officer",
    form: () => `<div class="fgrid">${inp("name", "Full name *")}${inp("designation", "Designation * (e.g. TSO)")}${isAdmin() ? inp("region", "Region") : `<label>Region<input value="${esc(S.me.region)}" disabled></label>`}${inp("territory", "Territory / base *")}${inp("zone", "Zone (e.g. Zone 1)")}${inp("personal", "Personal number")}${simChooser("sim")}${zoneSelect("rz")}</div>`,
    collect(f) {
      const v = k => (f.elements[k] && f.elements[k].value.trim()) || null;
      must(v("name") && v("designation") && v("territory"), "Name, designation and territory are required.");
      const d = { name: v("name"), designation: v("designation"), region: isAdmin() ? v("region") : S.me.region, territory: v("territory"), zone: v("zone"), personal: v("personal"), sim: pickedSim(f, "sim"), rz: v("rz") };
      return { data: d, summary: `New officer: ${d.name} (${d.designation}, ${d.territory}) · SIM ${d.sim ? fm(d.sim) : "none yet"}${d.rz ? " · report " + d.rz : ""}` };
    },
    async apply(d) {
      if (d.sim) must(!S.officers.some(o => o.corporate_sim === d.sim && o.status !== "Left"), `SIM ${fm(d.sim)} already belongs to another officer.`);
      const rec = { name: d.name, designation: d.designation, region: d.region, territory: d.territory, zone: d.zone, personal_number: d.personal, status: "Active",
        corporate_sim: d.sim, tracker_number: d.sim, report_zone: d.rz, report_order: d.rz ? nextOrder(d.rz) : null,
        report_name: d.rz ? d.name : null, report_designation: d.rz ? d.designation : null, report_base: d.rz ? d.territory : null,
        tt_name: d.sim ? `${d.name}_${d.designation}_${d.territory}` : null, remarks: d.sim ? `Joined ${today()}` : `Joined ${today()} — no corporate SIM yet` };
      const { data } = chk(await sb.from("officers").insert(rec).select().single());
      if (d.sim) await upsertSim(d.sim, { status: "Assigned", officer_id: data.id, zone: d.zone, remarks: null });
      return { msg: `${d.name} added${d.sim ? " with SIM " + fm(d.sim) : ""}.`, tt: d.sim ? `Register ${fm(d.sim)} as: ${rec.tt_name}` : "" };
    },
  },
  replace: {
    title: "New officer on the same number", btn: "Save replacement",
    form: () => `<p class="muted">The old officer is marked Left. The new person takes the same SIM and the same row in the daily report.</p>
      <div class="fgrid"><label>Officer who left<select name="old"><option value="">— choose —</option>${offOptions(o => o.status === "Active" && o.corporate_sim)}</select></label>
      ${inp("name", "New officer's full name *")}${inp("designation", "Designation (blank = same)")}${inp("territory", "Territory (blank = same)")}${inp("personal", "New officer's personal number")}</div>`,
    collect(f) {
      const old = offById(f.elements.old.value); must(old, "Choose the officer who left.");
      const name = f.elements.name.value.trim(); must(name, "Enter the new officer's name.");
      const d = { oldId: old.id, oldName: old.name, sim: old.corporate_sim, name, designation: f.elements.designation.value.trim() || old.designation, territory: f.elements.territory.value.trim() || old.territory, personal: f.elements.personal.value.trim() || null };
      return { data: d, summary: `${old.name} left → ${name} (${d.designation}) takes ${fm(d.sim)}` };
    },
    async apply(d) {
      const old = offById(d.oldId); must(old && old.status === "Active" && old.corporate_sim === d.sim, `${d.oldName} no longer holds ${fm(d.sim)} — nothing changed.`);
      chk(await sb.from("officers").update({ status: "Left", corporate_sim: null, report_zone: null, report_order: null, remarks: `Left ${today()}; SIM ${fm(d.sim)} handed to ${d.name}` }).eq("id", old.id));
      const { data } = chk(await sb.from("officers").insert({ name: d.name, designation: d.designation, region: old.region, territory: d.territory, zone: old.zone, personal_number: d.personal, status: "Active",
        corporate_sim: d.sim, tracker_number: d.sim, report_zone: old.report_zone, report_order: old.report_order, report_name: old.report_zone ? d.name : null,
        report_designation: old.report_zone ? d.designation : null, report_base: old.report_zone ? d.territory : null, tt_name: `${d.name}_${d.designation}_${d.territory}`, remarks: `Took over ${fm(d.sim)} from ${old.name} on ${today()}` }).select().single());
      await upsertSim(d.sim, { status: "Assigned", officer_id: data.id, remarks: null });
      return { msg: `${old.name} marked Left. ${d.name} now has ${fm(d.sim)}.`, tt: `On ${fm(d.sim)} rename the entry to: ${d.name}_${d.designation}_${d.territory}` };
    },
  },
  leave: {
    title: "Officer left", btn: "Mark as left",
    form: () => `<div class="fgrid"><label>Officer who left<select name="old"><option value="">— choose —</option>${offOptions()}</select></label>
      <label>Did they return the SIM?<select name="ret"><option value="yes">Yes — SIM is back (make it Spare)</option><option value="no">No — kept the SIM (needs replacement)</option></select></label>
      <label>Keep the post as Vacant?<select name="vac"><option value="no">No</option><option value="yes">Yes — record a vacant post</option></select></label></div>`,
    collect(f) {
      const old = offById(f.elements.old.value); must(old, "Choose the officer.");
      const d = { oldId: old.id, oldName: old.name, sim: old.corporate_sim, returned: f.elements.ret.value === "yes", vacant: f.elements.vac.value === "yes" };
      return { data: d, summary: `${old.name} (${old.designation}) left · SIM ${d.sim ? fm(d.sim) : "none"} ${d.returned ? "returned" : "NOT returned"}${d.vacant ? " · keep vacant post" : ""}` };
    },
    async apply(d) {
      const old = offById(d.oldId); must(old && old.status === "Active", `${d.oldName} is no longer active — nothing changed.`);
      chk(await sb.from("officers").update({ status: "Left", corporate_sim: null, report_zone: null, report_order: null, remarks: `Left ${today()}${d.sim ? "; SIM " + fm(d.sim) : ""}` }).eq("id", old.id));
      if (d.vacant) chk(await sb.from("officers").insert({ name: null, designation: old.designation, region: old.region, territory: old.territory, zone: old.zone, status: "Vacant post", remarks: `Vacant: ${old.name} left ${today()}` }));
      const simMsg = await freeSim(d.sim, d.returned, old.name);
      return { msg: `${old.name} marked Left and removed from the daily report. ${simMsg}`, tt: d.sim ? `On ${fm(d.sim)} change the name to "No officer assigned".` : "" };
    },
  },
  renumber: {
    title: "Officer's number changed", btn: "Change number",
    form: () => `<div class="fgrid"><label>Officer<select name="off"><option value="">— choose —</option>${offOptions()}</select></label>
      ${simChooser("sim")}<label>What happens to the old SIM?<select name="ret"><option value="yes">Returned — make it Spare</option><option value="no">Not returned / lost — needs replacement</option></select></label></div>`,
    collect(f) {
      const o = offById(f.elements.off.value); must(o, "Choose the officer.");
      const sim = pickedSim(f, "sim"); must(sim, "Choose or type the new SIM number.");
      const d = { id: o.id, name: o.name, oldSim: o.corporate_sim, sim, returned: f.elements.ret.value === "yes" };
      return { data: d, summary: `${o.name}: ${d.oldSim ? fm(d.oldSim) : "no SIM"} → ${fm(sim)} · old SIM ${d.returned ? "returned" : "NOT returned"}` };
    },
    async apply(d) {
      const o = offById(d.id); must(o && o.corporate_sim === d.oldSim, `${d.name}'s number changed since the request — nothing changed.`);
      must(!S.officers.some(x => x.corporate_sim === d.sim && x.status !== "Left" && x.id !== o.id), `${fm(d.sim)} already belongs to another officer.`);
      chk(await sb.from("officers").update({ corporate_sim: d.sim, tracker_number: d.sim, remarks: `Number changed ${today()}${d.oldSim ? " from " + fm(d.oldSim) : ""}` }).eq("id", o.id));
      await upsertSim(d.sim, { status: "Assigned", officer_id: o.id, zone: o.zone, remarks: null });
      const simMsg = d.oldSim ? await freeSim(d.oldSim, d.returned, o.name) : "";
      return { msg: `${o.name} now uses ${fm(d.sim)}. ${simMsg}`, tt: `Register ${fm(d.sim)} as: ${o.tt_name || o.name + "_" + o.designation + "_" + (o.territory || "")}${d.oldSim ? `\nOn ${fm(d.oldSim)} change the name to "No officer assigned".` : ""}` };
    },
  },
  simstatus: {
    title: "Change a SIM's status", btn: "Save",
    form: () => `<div class="fgrid">${inp("num", "SIM number *")}<label>New status<select name="st">${SIM_STATUS.filter(s => !s.startsWith("Assigned")).map(s => `<option>${s}</option>`).join("")}</select></label>${inp("rem", "Note (why)")}</div>
      <p class="muted">To give a SIM to a person, use "New officer joined" or "Officer's number changed" instead.</p>`,
    collect(f) {
      const n = GPReport.normMobile(f.elements.num.value); must(n, "Enter a valid number.");
      const d = { number: n, status: f.elements.st.value, note: f.elements.rem.value.trim() };
      return { data: d, summary: `SIM ${fm(n)} → ${d.status}${d.note ? " (" + d.note + ")" : ""}` };
    },
    async apply(d) {
      const holder = S.officers.find(o => o.corporate_sim === d.number && o.status === "Active");
      if (holder) chk(await sb.from("officers").update({ corporate_sim: null, remarks: `SIM ${fm(d.number)} taken back ${today()}` }).eq("id", holder.id));
      await upsertSim(d.number, { status: d.status, officer_id: null, remarks: (d.note || d.status) + ` (${today()})` });
      return { msg: `${fm(d.number)} is now "${d.status}".${holder ? " Taken from " + holder.name + "." : ""}`, tt: holder ? `On ${fm(d.number)} change the name to "No officer assigned".` : "" };
    },
  },
};

document.querySelectorAll(".choice").forEach(b => b.onclick = () => {
  document.querySelectorAll(".choice").forEach(x => x.classList.toggle("sel", x === b));
  showAction(b.dataset.act);
});
function showAction(act) {
  const A = ACTIONS[act];
  $("actForm").innerHTML = `<div class="card"><h2>${A.title}</h2><form id="af">${A.form()}<div class="toolbar"><button class="primary" type="submit">${isAdmin() ? A.btn : "Send for approval"}</button></div></form><div id="afOut"></div></div>`;
  const f = $("af");
  f.onsubmit = async e => {
    e.preventDefault();
    const b = f.querySelector("button[type=submit]"); b.disabled = true;
    try {
      const { data, summary } = A.collect(f);
      if (isAdmin()) {
        b.textContent = "Saving…";
        const r = await A.apply(data);
        await log(act, summary);
        await loadAll();
        $("afOut").innerHTML = `<div class="todo"><b>✓ Saved.</b> ${esc(r.msg)}${r.tt ? `\n\n<b>Now do this in Team Tracker:</b>\n${esc(r.tt)}` : ""}</div>`;
      } else {
        b.textContent = "Sending…";
        chk(await sb.from("change_requests").insert({ region: S.me.region, action: act, summary, payload: data }));
        await log("request_" + act, summary);
        $("afOut").innerHTML = `<div class="todo"><b>✓ Sent for approval.</b> ${esc(summary)}\nNothing changes until the admin approves it. You can see the status in <b>My requests</b>.</div>`;
        renderRequests();
      }
    } catch (err) { alert("Not saved: " + (err.message || err)); b.disabled = false; b.textContent = "Try again"; }
  };
}

/* ---------------- approvals ---------------- */
async function renderRequests() {
  const { data, error } = await sb.from("change_requests").select("*").order("created_at", { ascending: false }).limit(200);
  if (error) { $("reqOut").textContent = error.message; return; }
  const pending = (data || []).filter(r => r.status === "pending");
  const badge = $("reqBadge"); if (badge) badge.textContent = isAdmin() && pending.length ? ` (${pending.length})` : "";
  const row = r => `<tr><td>${new Date(r.created_at).toLocaleString()}</td><td>${esc(r.region || "")}</td><td>${esc(r.requested_by)}</td><td>${esc(ACTIONS[r.action]?.title || r.action)}</td><td>${esc(r.summary)}</td>
    <td><span class="pill ${r.status}">${r.status}</span>${r.note ? `<br><small>${esc(r.note)}</small>` : ""}</td>
    <td>${isAdmin() && r.status === "pending" ? `<button class="primary" data-ok="${r.id}">Approve</button> <button data-no="${r.id}">Reject</button>` : esc(r.reviewed_by || "")}</td></tr>`;
  $("reqOut").innerHTML = `<p>${isAdmin() ? `${pending.length} waiting for your approval.` : "Your region's requests. The admin approves or rejects each one."}</p>
    <div class="tablewrap"><table class="grid"><tr><th>When</th><th>Region</th><th>By</th><th>Type</th><th>Change</th><th>Status</th><th></th></tr>${(data || []).map(row).join("")}</table></div>`;
  $("reqOut").querySelectorAll("[data-ok]").forEach(b => b.onclick = () => decide(data.find(r => r.id == b.dataset.ok), true));
  $("reqOut").querySelectorAll("[data-no]").forEach(b => b.onclick = () => decide(data.find(r => r.id == b.dataset.no), false));
}
async function decide(r, ok) {
  try {
    if (ok) {
      if (!confirm(`Approve and apply?\n\n${r.summary}\n(from ${r.requested_by}, ${r.region})`)) return;
      await loadAll();
      const res = await ACTIONS[r.action].apply(r.payload);
      chk(await sb.from("change_requests").update({ status: "approved", reviewed_by: S.user.email, reviewed_at: new Date().toISOString(), note: res.tt ? "Team Tracker: " + res.tt : null }).eq("id", r.id));
      await log("approved_" + r.action, `${r.summary} (requested by ${r.requested_by})`);
      await loadAll();
      alert("✓ Applied. " + res.msg + (res.tt ? "\n\nTeam Tracker:\n" + res.tt : ""));
    } else {
      const note = prompt("Reason for rejecting (the lead will see this):", "");
      if (note === null) return;
      chk(await sb.from("change_requests").update({ status: "rejected", reviewed_by: S.user.email, reviewed_at: new Date().toISOString(), note }).eq("id", r.id));
      await log("rejected_" + r.action, `${r.summary} (requested by ${r.requested_by}) — ${note}`);
    }
  } catch (e) { alert("Not applied: " + (e.message || e)); }
  renderRequests();
}

/* ---------------- master Excel (same sheets as Officer_Master) ---------------- */
const isHQ = o => o && o.status === "HQ / Not tracked";
const isLeaf = o => o && (o.report_zone === "LEAF" || o.region === "Leaf");
function seriesRank(n) { return n.startsWith("17") ? 0 : n.startsWith("13") ? 1 : 2; }
function sheet(rows, widths) { const ws = XLSX.utils.json_to_sheet(rows); if (widths) ws["!cols"] = widths.map(w => ({ wch: w })); return ws; }
function exportMaster() {
  const offById = new Map(S.officers.map(o => [o.id, o]));
  const holder = s => offById.get(s.officer_id) || S.officers.find(o => o.corporate_sim === s.number && o.status !== "Left");
  const wb = XLSX.utils.book_new();
  const corp = S.officers.filter(o => o.corporate_sim && o.status !== "Left")
    .sort((a, b) => (a.region || "").localeCompare(b.region || "") || (a.designation || "").localeCompare(b.designation || "") || (a.name || "").localeCompare(b.name || ""));
  XLSX.utils.book_append_sheet(wb, sheet(corp.map((o, i) => ({ "SL": i + 1, "Name": o.name, "Designation": o.designation, "Region": o.region, "Territory": o.territory, "Zone": o.zone,
    "Corporate SIM": fmN(o.corporate_sim), "Personal Number": o.personal_number || "", "Status": o.status, "Name in Team Tracker": o.tt_name || "", "Remarks": o.remarks || "" })), [5, 28, 18, 14, 22, 10, 15, 15, 16, 40, 40]), "All corporate");
  const seq = S.sims.slice().sort((a, b) => seriesRank(a.number) - seriesRank(b.number) || a.number.localeCompare(b.number));
  XLSX.utils.book_append_sheet(wb, sheet(seq.map((s, i) => { const o = holder(s); return { "SL": i + 1, "Number": fmN(s.number), "Series": s.series, "Status": s.status,
    "Holder": o?.name || "", "Designation": o?.designation || "", "Region": o?.region || "", "Zone": s.zone || o?.zone || "", "Remarks": s.remarks || "" }; }), [5, 15, 11, 24, 28, 18, 14, 10, 50]), "Sequential Numbers");
  const rw = [];
  const regions = [...new Set(corp.map(o => o.region || "(no region)"))].sort();
  for (const r of regions) {
    const list = corp.filter(o => (o.region || "(no region)") === r);
    rw.push({ "SL": "", "Name": `${r} (${list.length})` });
    list.forEach((o, i) => rw.push({ "SL": i + 1, "Name": o.name, "Designation": o.designation, "Territory": o.territory, "Corporate SIM": fmN(o.corporate_sim), "Personal Number": o.personal_number || "", "Status": o.status }));
    rw.push({});
  }
  XLSX.utils.book_append_sheet(wb, sheet(rw, [5, 30, 18, 24, 15, 15, 16]), "Region wise Corporate");
  const spareRows = st => S.sims.filter(s => st.includes(s.status)).sort((a, b) => seriesRank(a.number) - seriesRank(b.number) || a.number.localeCompare(b.number))
    .map((s, i) => ({ "SL": i + 1, "Number": fmN(s.number), "Series": s.series, "Status": s.status, "Zone": s.zone || "", "Remarks": s.remarks || "" }));
  XLSX.utils.book_append_sheet(wb, sheet(spareRows(["Spare", "Reserved - Chakaria", "Check"]), [5, 15, 11, 20, 10, 70]), "Active Spare SIM");
  XLSX.utils.book_append_sheet(wb, sheet(spareRows(["Spare (inactive SIM)"]), [5, 15, 11, 20, 10, 50]), "Inactive Spare SIM");
  XLSX.utils.book_append_sheet(wb, sheet(spareRows(["Suspended"]), [5, 15, 11, 20, 10, 70]), "Suspended SIM");
  XLSX.utils.book_append_sheet(wb, sheet(spareRows(["Closed - Replace"]), [5, 15, 11, 20, 10, 70]), "SIM Replacement Required");
  XLSX.utils.book_append_sheet(wb, sheet(S.officers.filter(o => o.status === "Active" && !o.corporate_sim).map((o, i) => ({ "SL": i + 1, "Name": o.name, "Designation": o.designation, "Region": o.region, "Territory": o.territory, "Personal Number": o.personal_number || "", "Remarks": o.remarks || "" })), [5, 28, 18, 14, 22, 15, 50]), "No SIM List");
  XLSX.utils.book_append_sheet(wb, sheet(S.officers.filter(o => o.status === "HQ / Not tracked").map((o, i) => ({ "SL": i + 1, "Name": o.name, "Designation": o.designation, "Region": o.region, "Corporate SIM": fmN(o.corporate_sim), "Locked": "🔒 HQ — cannot be changed" })), [5, 34, 26, 16, 15, 26]), "HQ Numbers (locked)");
  const cnt = f => S.sims.filter(f).length, ocnt = f => S.officers.filter(f).length;
  const dash = [["GP Tracker — Officer & SIM summary", ""], ["Exported", new Date().toLocaleString()], [],
    ["Active officers", ocnt(o => o.status === "Active")], ["On daily report", ocnt(o => o.status === "Active" && o.report_zone)], ["HQ / not tracked", ocnt(isHQ)], ["Left", ocnt(o => o.status === "Left")], ["Vacant posts", ocnt(o => o.status === "Vacant post")], ["Active officers without SIM", ocnt(o => o.status === "Active" && !o.corporate_sim)], [],
    ["Total SIMs", S.sims.length], ["017 series", cnt(s => s.number.startsWith("17"))], ["013 series", cnt(s => s.number.startsWith("13"))], [],
    ...SIM_STATUS.map(st => [st, cnt(s => s.status === st)]), [],
    ["Officers per region", ""], ...[...new Set(S.officers.filter(o => o.status === "Active").map(o => o.region || "(none)"))].sort().map(r => [r, ocnt(o => o.status === "Active" && (o.region || "(none)") === r)])];
  const ds = XLSX.utils.aoa_to_sheet(dash); ds["!cols"] = [{ wch: 34 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(wb, ds, "Dashboard");
  XLSX.writeFile(wb, `Officer_Master_${ddmmyyyy(new Date())}.xlsx`);
  log("export_master", "Downloaded master Excel");
}

$("btnMaster").onclick = exportMaster;
boot();
