/* HB-Tracker core: auth, layout, routing, helpers */
const C = window.HB_CONFIG;
const sb = supabase.createClient(C.url, C.key, { auth: { persistSession: true, autoRefreshToken: true } });
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const S = { me: null, regions: [], route: "", pages: {} };
window.HB = S;

function toast(t, ms = 3000) { const e = $("#toast"); e.textContent = t; e.classList.remove("hidden"); clearTimeout(e._t); e._t = setTimeout(() => e.classList.add("hidden"), ms); }
function fail(e) { console.error(e); toast("Error: " + (e?.message || e), 6000); }
async function q(p) { const { data, error } = await p; if (error) throw error; return data; }
const rpc = (fn, args) => q(sb.rpc(fn, args));
const fmtDate = d => d ? new Date(d).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "";
const fmtTime = d => d ? new Date(d).toLocaleString("en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : "";
const today = () => new Date().toISOString().slice(0, 10);
const regionName = id => (S.regions.find(r => r.id === id) || {}).name || "";
const normNum = n => { const d = String(n ?? "").replace(/\D/g, ""); if (d.length < 10) return ""; return "0" + d.slice(-10); };
function pill(t, cls) { return `<span class="pill ${cls || ""}">${esc(t)}</span>`; }
const SIM_LABEL = { active: ["Active", "p-blue"], spare: ["Active spare", "p-green"], spare_inactive: ["Inactive spare", ""], suspended: ["Suspended", ""], hq_pool: ["HQ pool", "p-amber"], closed: ["Closed", "p-red"] };
const simPill = s => s.locked ? pill("HQ locked", "p-navy") : pill(...(SIM_LABEL[s.status] || [s.status]));
const REQ = { new_joiner: "New joiner", resignation: "Resignation", transfer: "Transfer", promotion: "Promotion", need_sim: "SIM request", edit_officer: "Officer details", submit_list: "Officer list" };
const reqPill = s => pill(s[0].toUpperCase() + s.slice(1), { pending: "p-amber", approved: "p-green", rejected: "p-red" }[s]);
function table(cols, rows, empty = "Nothing here yet.") {
  if (!rows.length) return `<div class="empty">${empty}</div>`;
  return `<div class="tw"><table class="t"><thead><tr>${cols.map(c => `<th>${c}</th>`).join("")}</tr></thead><tbody>${rows.join("")}</tbody></table></div>`;
}
const tr = (cells, attrs = "") => `<tr ${attrs}>${cells.map(c => `<td>${c ?? ""}</td>`).join("")}</tr>`;
function head(t, sub, acts = "") { return `<div class="head"><div><h1>${t}</h1>${sub ? `<p>${sub}</p>` : ""}</div><div class="acts">${acts}</div></div>`; }
function stats(items) { return `<div class="stats">${items.map(([l, v, s, c]) => `<div class="stat"><div class="l">${l}</div><div class="v" style="color:${c || "inherit"}">${v}</div><div class="s">${s || ""}</div></div>`).join("")}</div>`; }
function regionOptions(sel, { all, fieldOnly } = {}) {
  return (all ? `<option value="">${all}</option>` : "") + S.regions.filter(r => !fieldOnly || r.kind === "field").map(r => `<option value="${r.id}" ${+sel === r.id ? "selected" : ""}>${esc(r.name)}</option>`).join("");
}
function dialog(html, onSubmit) {
  const d = $("#dlg"), f = $("#dlgBody");
  f.innerHTML = html + `<p class="msg" id="dlgMsg"></p><div class="dacts"><button value="cancel" formnovalidate>Cancel</button><button class="primary" value="ok" id="dlgOk">OK</button></div>`;
  f.onsubmit = async ev => {
    if (ev.submitter?.value !== "ok") return;
    ev.preventDefault();
    try { $("#dlgOk").disabled = true; await onSubmit(f); d.close(); } catch (e) { $("#dlgMsg").textContent = e.message || e; } finally { const b = $("#dlgOk"); if (b) b.disabled = false; }
  };
  d.showModal();
  return f;
}
function readExcel(file) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => { try { res(XLSX.read(r.result, { type: "array", cellDates: true })); } catch (e) { rej(e); } }; r.onerror = rej; r.readAsArrayBuffer(file); });
}
function sheetObjects(wb) { const ws = wb.Sheets[wb.SheetNames[0]]; return XLSX.utils.sheet_to_json(ws, { defval: "" }); }
function pick(row, ...names) { for (const k of Object.keys(row)) { const n = k.toLowerCase().replace(/[^a-z]/g, ""); if (names.some(x => n.includes(x))) return String(row[k]).trim(); } return ""; }
function parseOfficerSheet(rows) {
  return rows.map(r => ({ name: pick(r, "name", "officer"), designation: pick(r, "designation", "desig"), region: pick(r, "region", "division", "zone"), territory: pick(r, "territory", "area", "base"), sim_number: normNum(pick(r, "corporate", "sim", "officialnumber", "mobile")), personal_number: normNum(pick(r, "personal")) })).filter(o => o.name);
}
function downloadXlsx(name, sheets) {
  const wb = XLSX.utils.book_new();
  for (const [n, rows] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ "": "empty" }]), n.slice(0, 31));
  XLSX.writeFile(wb, name);
}
function downloadCsv(name, rows) {
  const ws = XLSX.utils.json_to_sheet(rows); const csv = XLSX.utils.sheet_to_csv(ws);
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv" })); a.download = name; a.click();
}

/* ---------- auth ---------- */
let mode = "signin";
function showLogin(msg = "") { $("#app").classList.add("hidden"); $("#login").classList.remove("hidden"); $("#loginMsg").textContent = msg; document.title = "Sign in"; }
$("#lnkFirst").onclick = e => { e.preventDefault(); mode = mode === "first" ? "signin" : "first"; $("#btnLogin").textContent = mode === "first" ? "Set password" : "Sign in"; $("#lnkFirst").textContent = mode === "first" ? "Back to sign in" : "First time? Set your password"; $("#loginMsg").textContent = mode === "first" ? "Use the email HQ invited. Password: at least 8 characters." : ""; };
$("#lnkReset").onclick = async e => {
  e.preventDefault(); const em = $("#email").value.trim(); if (!em) { $("#loginMsg").textContent = "Type your email first."; return; }
  await sb.auth.resetPasswordForEmail(em, { redirectTo: location.origin + location.pathname });
  $("#loginMsg").textContent = "If this email has access, a reset link has been sent.";
};
$("#loginForm").onsubmit = async e => {
  e.preventDefault(); const em = $("#email").value.trim(), pw = $("#pass").value; $("#loginMsg").textContent = ""; $("#btnLogin").disabled = true;
  try {
    if (mode === "first") {
      const { data, error } = await sb.auth.signUp({ email: em, password: pw });
      if (error) throw new Error("invalid sign-in");
      if (!data.session) { $("#loginMsg").textContent = "Check your email to confirm, then sign in."; return; }
    } else {
      const { error } = await sb.auth.signInWithPassword({ email: em, password: pw });
      if (error) throw new Error("invalid sign-in");
    }
    await start(true);
  } catch (err) { $("#loginMsg").textContent = "Invalid sign-in."; } finally { $("#btnLogin").disabled = false; }
};
sb.auth.onAuthStateChange(async (ev) => {
  if (ev === "PASSWORD_RECOVERY") {
    const pw = prompt("Set a new password (at least 8 characters):");
    if (pw && pw.length >= 8) { const { error } = await sb.auth.updateUser({ password: pw }); toast(error ? "Could not set password" : "Password changed"); }
  }
});
async function signOut(msg) { await sb.auth.signOut(); S.me = null; showLogin(msg || ""); }

let idleT;
function bumpIdle() { clearTimeout(idleT); idleT = setTimeout(() => signOut("Signed out after " + C.autoSignOutMinutes + " minutes without activity."), C.autoSignOutMinutes * 60000); }
["click", "keydown", "touchstart"].forEach(ev => document.addEventListener(ev, () => S.me && bumpIdle(), { passive: true }));

async function start(fresh) {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) return showLogin();
  let me = await q(sb.from("hb_profiles").select("*").eq("user_id", session.user.id).maybeSingle()).catch(() => null);
  if (!me && await rpc("hb_claim_invite").catch(() => false)) me = await q(sb.from("hb_profiles").select("*").eq("user_id", session.user.id).maybeSingle()).catch(() => null);
  if (!me || me.status !== "active") { await sb.auth.signOut(); return showLogin("Invalid sign-in."); }
  S.me = me; S.admin = me.role === "admin";
  S.regions = await q(sb.from("hb_regions").select("*").order("name"));
  if (fresh) rpc("hb_touch").catch(() => {});
  $("#login").classList.add("hidden"); $("#app").classList.remove("hidden");
  document.title = "HB-Tracker";
  bumpIdle(); renderNav(); route(); refreshBell(); clearInterval(S.poll); S.poll = setInterval(() => { refreshBell(); checkStillActive(); }, 60000);
}
async function checkStillActive() {
  const me = await q(sb.from("hb_profiles").select("status").eq("user_id", S.me.user_id).maybeSingle()).catch(() => null);
  if (!me || me.status !== "active") signOut("Your access has been removed.");
}

/* ---------- nav & routing ---------- */
const ICON = { dash: '<rect x="3" y="3" width="7" height="9"/><rect x="14" y="3" width="7" height="5"/><rect x="14" y="12" width="7" height="9"/><rect x="3" y="16" width="7" height="5"/>', file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/>', match: '<path d="M7 7h10"/><path d="m13 3 4 4-4 4"/><path d="M17 17H7"/><path d="m11 13-4 4 4 4"/>', chat: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>', check: '<path d="M20 6 9 17l-5-5"/>', users: '<circle cx="9" cy="7" r="4"/><path d="M2 21v-1a6 6 0 0 1 12 0v1"/><path d="M16 3.1a4 4 0 0 1 0 7.8"/>', sim: '<rect x="5" y="2" width="14" height="20" rx="2"/><path d="M9 18h6"/>', box: '<path d="M21 8 12 3 3 8v8l9 5 9-5z"/><path d="M3 8l9 5 9-5"/><path d="M12 13v8"/>', edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>', list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>', msg: '<path d="M4 4h16v12H8l-4 4z"/><path d="M8 9h8M8 12h5"/>', key: '<circle cx="8" cy="15" r="4"/><path d="m11 12 9-9M17 6l3 3"/>' };
const ic = n => `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">${ICON[n]}</svg>`;
const NAV_ADMIN = [["dashboard", "Dashboard", "dash"], ["crossmatch", "Cross-match", "match"], ["report", "Daily Report", "file"], ["justify", "Justifications", "chat", "just"], ["approvals", "Approvals", "check", "req"], ["officers", "Officers", "users"], ["sims", "SIM Register", "sim"], ["pool", "HQ Spare Pool", "box"], ["history", "Report History", "list"], ["users", "Users", "key"], ["messages", "Messages", "msg", "msg"]];
const NAV_LEAD = [["home", "My Region", "dash"], ["ljustify", "Justify Today", "chat", "just"], ["lofficers", "My Officers", "users"], ["spares", "My Spare SIMs", "sim"], ["change", "Report a Change", "edit"], ["requests", "My Requests", "list"], ["lmessages", "Message HQ", "msg", "msg"]];
function renderNav() {
  const items = S.admin ? NAV_ADMIN : NAV_LEAD;
  $("#nav").innerHTML = `<div class="brand"><img src="assets/leaf.png" alt="GLTC"><div><b>HB-Tracker</b><span>${S.admin ? "HQ · all regions" : esc(regionName(S.me.region_id)) + " region"}</span></div></div>` +
    items.map(([k, l, i, b]) => `<a href="#/${k}" data-k="${k}">${ic(i)}${l}${b ? `<span class="badge hidden" data-b="${b}"></span>` : ""}</a>`).join("") +
    `<div class="me"><div class="av">${esc((S.me.full_name || S.me.email)[0].toUpperCase())}</div><div><b>${esc(S.me.full_name || S.me.email)}</b>${S.admin ? "Admin" : "Region lead"}<br><a href="#" id="btnPw" style="color:#9FB0C9;font-size:12px">Change password</a></div><button id="btnOut">Sign out</button></div>`;
  $("#btnOut").onclick = () => signOut();
  $("#btnPw").onclick = e => { e.preventDefault(); dialog(`<h3>Set a new password</h3><label>New password (at least 8 characters)<input name="p" type="password" minlength="8" required autocomplete="new-password"></label><label>Type it again<input name="p2" type="password" minlength="8" required autocomplete="new-password"></label>`, async f => { if (f.p.value !== f.p2.value) throw new Error("The two passwords are different"); const { error } = await sb.auth.updateUser({ password: f.p.value }); if (error) throw error; toast("Password changed. Use it next time you sign in."); }); };
  $("#scope").innerHTML = S.admin ? "Admin view · all regions" : `Lead view · you can only see ${esc(regionName(S.me.region_id))}`;
}
$("#btnMenu").onclick = () => $("#nav").classList.toggle("open");
window.addEventListener("hashchange", route);
async function route() {
  if (!S.me) return;
  const def = S.admin ? "dashboard" : "home";
  let k = (location.hash.match(/^#\/(\w+)/) || [])[1] || def;
  const allowed = (S.admin ? NAV_ADMIN : NAV_LEAD).map(x => x[0]);
  if (!allowed.includes(k)) k = def;
  S.route = k;
  document.querySelectorAll("#nav a").forEach(a => a.classList.toggle("on", a.dataset.k === k));
  $("#nav").classList.remove("open");
  const el = $("#page"); el.innerHTML = `<div class="empty">Loading…</div>`;
  try { await S.pages[k](el); } catch (e) { el.innerHTML = `<div class="warn">Could not load: ${esc(e.message || e)}</div>`; console.error(e); }
}
const reload = () => route();

/* ---------- notifications ---------- */
async function latestReport() {
  if (S.admin) return (await q(sb.from("hb_reports").select("*").order("report_date", { ascending: false }).limit(1)))[0] || null;
  const r = (await q(sb.from("hb_report_rows").select("report_id,report_date").order("report_id", { ascending: false }).limit(1)))[0];
  return r ? { id: r.report_id, report_date: r.report_date, created_at: r.report_date } : null;
}
async function refreshBell() {
  try {
    const items = []; let just = 0, req = 0, msg = 0;
    const rep = await latestReport();
    if (rep) {
      let qq = sb.from("hb_report_rows").select("id", { count: "exact", head: true }).eq("report_id", rep.id).not("issue", "is", null).is("justified_at", null);
      qq = qq.not("sent_at", "is", null);
      const { count } = await qq; just = count || 0;
    }
    if (S.admin) {
      const reqs = await q(sb.from("hb_requests").select("*").eq("status", "pending").order("created_at", { ascending: false }).limit(20));
      req = reqs.length;
      reqs.forEach(r => items.push({ t: `${REQ[r.kind]} · ${regionName(r.region_id)}`, d: reqSummary(r), when: r.created_at, req: r }));
      const m = await q(sb.from("hb_messages").select("lead_id,body,created_at,sender").is("read_at", null).neq("sender", S.me.user_id).order("created_at", { ascending: false }));
      msg = m.length; if (msg) items.push({ t: `${msg} unread message${msg > 1 ? "s" : ""}`, d: m[0].body.slice(0, 80), when: m[0].created_at, href: "#/messages" });
      if (just) items.push({ t: `${just} justifications waiting`, d: "Latest report " + fmtDate(rep.report_date), href: "#/justify" });
    } else {
      if (just) items.push({ t: "Today's list is ready", d: `${just} officers to justify`, href: "#/ljustify", when: rep.created_at });
      const done = await q(sb.from("hb_requests").select("*").neq("status", "pending").gte("decided_at", new Date(Date.now() - 3 * 864e5).toISOString()).order("decided_at", { ascending: false }));
      done.forEach(r => items.push({ t: `HQ ${r.status} your request`, d: `${REQ[r.kind]}: ${reqSummary(r)}`, when: r.decided_at, href: "#/requests" }));
      const m = await q(sb.from("hb_messages").select("body,created_at").eq("lead_id", S.me.user_id).neq("sender", S.me.user_id).is("read_at", null));
      msg = m.length; if (msg) items.push({ t: "New message from HQ", d: m[0].body.slice(0, 80), when: m[0].created_at, href: "#/lmessages" });
    }
    const n = S.admin ? req + msg + (just ? 1 : 0) : (just ? 1 : 0) + msg + items.filter(i => i.href === "#/requests").length;
    const dot = $("#bellCount"); dot.textContent = n; dot.classList.toggle("hidden", !n);
    document.querySelectorAll("[data-b]").forEach(b => { const v = { just, req, msg }[b.dataset.b]; b.textContent = v; b.classList.toggle("hidden", !v); });
    S.bellItems = items; if (!$("#bellPanel").classList.contains("hidden")) drawBell();
  } catch (e) { console.warn(e); }
}
function drawBell() {
  const p = $("#bellPanel");
  p.innerHTML = `<div class="hd"><b>Notifications</b></div>` + ((S.bellItems || []).map((it, i) => `<div class="bell-item"><b>${esc(it.t)}</b><div class="muted">${esc(it.d)}</div><div class="muted" style="font-size:12px">${fmtTime(it.when)}</div>${it.req ? `<div class="acts"><button class="sm" data-rej="${i}">Reject</button><button class="sm primary" data-ok="${i}">Approve</button><a href="#/approvals" style="font-size:13px;margin-left:auto">Details</a></div>` : it.href ? `<a href="${it.href}" style="font-size:13px">Open</a>` : ""}</div>`).join("") || `<div class="empty">You're all caught up.</div>`);
  p.querySelectorAll("[data-ok],[data-rej]").forEach(b => b.onclick = async () => {
    const it = S.bellItems[b.dataset.ok ?? b.dataset.rej]; const ok = b.dataset.ok !== undefined;
    if (ok && (needsSim(it.req) || it.req.kind === "transfer")) { location.hash = "#/approvals"; p.classList.add("hidden"); return; }
    try { await rpc("hb_decide", { p_id: it.req.id, p_approve: ok, p_note: null, p_sim: null }); toast(ok ? "Approved" : "Rejected"); refreshBell(); if (S.route === "approvals") reload(); } catch (e) { fail(e); }
  });
  p.querySelectorAll("a").forEach(a => a.onclick = () => p.classList.add("hidden"));
}
$("#btnBell").onclick = e => { e.stopPropagation(); const p = $("#bellPanel"); p.classList.toggle("hidden"); if (!p.classList.contains("hidden")) drawBell(); };
document.addEventListener("click", e => { if (!e.target.closest(".bell-wrap")) $("#bellPanel").classList.add("hidden"); });
const needsSim = r => r.kind === "need_sim" || (r.kind === "new_joiner" && !r.payload.sim_number);
function reqSummary(r) {
  const p = r.payload || {};
  switch (r.kind) {
    case "new_joiner": return `${p.name}, ${p.designation || ""}${p.sim_number ? " · SIM " + p.sim_number : " · needs SIM"}`;
    case "resignation": return `${p.officer_name} · last day ${p.last_day || "?"} · SIM ${p.sim_returned ? "returned" : "NOT returned"}`;
    case "transfer": return `${p.officer_name} → ${regionName(+p.to_region_id) || p.to_region_name || ""}`;
    case "promotion": return `${p.officer_name}: ${p.old_designation || ""} → ${p.designation}`;
    case "need_sim": return p.reason || "Needs a SIM for a new joiner";
    case "edit_officer": return `${p.officer_name}: details update`;
    case "submit_list": return p.officers ? `${p.officers.length} officers` : `List update: ${(p.add||[]).length} new, ${(p.update||[]).length} changed, ${(p.left||[]).length} left, ${(p.missing||[]).length} missing`;
  }
  return "";
}

/* ---------- report rendering (shared) ---------- */
function renderReportHtml(zones, date, wm) {
  const d = date instanceof Date ? date : new Date(date);
  const ds = d.toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" });
  return zones.map(z => `<div class="rep-zone"><div class="rep-title">${esc(z.title)} — GPS Location Report</div><div class="rep-date">${ds}</div>
  <table class="rep"><colgroup><col style="width:3%"><col style="width:12%"><col style="width:6%"><col style="width:8%"><col style="width:9%"><col><col><col><col><col style="width:12%"></colgroup>
  <thead><tr><th>SL</th><th>Name</th><th>Desig.</th><th>Territory</th><th>Number</th>${C.checkpoints.map(c => `<th>${c[0]}</th>`).join("")}<th>Remarks / Justification</th></tr></thead>
  <tbody>${z.rows.map(r => `<tr><td>${r.sl}</td><td>${esc(r.name)}</td><td>${esc(r.desig)}</td><td>${esc(r.base)}</td><td>${esc(r.mobile)}</td>${r.cells.map(c => `<td class="${c.red ? "red" : ""}">${esc(c.txt)}</td>`).join("")}<td class="${r.rred ? "red" : ""}">${esc(r.remarks)}${r.reason ? `<br><span style="color:#0F172A">↳ ${esc(r.reason)}${r.note ? ": " + esc(r.note) : ""}</span>` : ""}</td></tr>`).join("")}</tbody></table>
  ${wm ? `<div class="wm">${esc(wm)}</div>` : ""}</div>`).join("");
}
function reportPdf(zones, date, name) {
  const wm = `Confidential · downloaded by ${S.me.full_name || S.me.email} · ${new Date().toLocaleString("en-GB")}`;
  const box = document.createElement("div"); box.className = "pdfmode";
  box.innerHTML = zones.map((z, i) => `<div style="${i ? "page-break-before:always" : ""}">${renderReportHtml([z], date, wm)}</div>`).join("");
  return html2pdf().set({ margin: [8, 6, 8, 6], filename: name, image: { type: "jpeg", quality: .95 }, html2canvas: { scale: 2 }, jsPDF: { unit: "mm", format: "a4", orientation: "landscape" }, pagebreak: { mode: ["css"] } }).from(box).save();
}
function reportCsvRows(zones) {
  const out = [];
  zones.forEach(z => z.rows.forEach(r => { const o = { Region: z.title, SL: r.sl, Name: r.name, Designation: r.desig, Territory: r.base, Number: r.mobile }; C.checkpoints.forEach((c, i) => o[c[0]] = r.cells[i].txt); o.Remarks = r.remarks; o.Justification = (r.reason || "") + (r.note ? ": " + r.note : ""); out.push(o); }));
  return out;
}
// rebuild zones from saved rows
function zonesFromRows(rows) {
  const by = {};
  rows.forEach(r => { (by[r.region_id] = by[r.region_id] || []).push(r); });
  return Object.entries(by).map(([rid, list]) => ({ title: regionName(+rid) || "Region", rows: list.sort((a, b) => a.checkpoints.sl - b.checkpoints.sl).map(r => ({ ...r.checkpoints, reason: r.reason, note: r.note })) })).sort((a, b) => a.title.localeCompare(b.title));
}

window.addEventListener("load", () => start(false));
