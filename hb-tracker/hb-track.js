/* Track History: one officer's Team Tracker export -> summary + PDF */
const TK = {};
TK.parse = wb => {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  const hi = raw.findIndex(r => /employee/i.test(String(r[0])) && /mobile/i.test(String(r[1])));
  if (hi < 0) throw new Error("This is not a Team Tracker track history file");
  const rows = raw.slice(hi + 1).filter(r => String(r[1]).trim() && String(r[2]).trim());
  return rows.map(r => ({ emp: String(r[0]).split("\n")[0].trim(), mob: normNum(r[1]), st: TK.dt(r[2]), en: TK.dt(r[3]), dur: String(r[4]).trim(), status: String(r[5]).trim(),
    moved: TK.km(r[6]), movedTxt: String(r[6]).trim(), addr: String(r[7]).replace(/[^\x00-ɏ\s.,/'’()-]/g, "").replace(/\s+,/g, ",").replace(/\s{2,}/g, " ").trim(), ll: String(r[8]).replace(/\s+/g, " ").trim(), base: TK.km(r[11]), baseTxt: String(r[11]).trim(), poi: String(r[12] || "").trim() }));
};
TK.km = s => { const m = String(s).match(/([\d.]+)\s*(km|m)\b/i); return m ? +m[1] * (m[2].toLowerCase() === "km" ? 1 : .001) : 0; };
TK.dt = v => { if (v instanceof Date) return v; const m = String(v).match(/(\d{1,2})-(\w{3})-(\d{4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i); if (!m) return null; let h = +m[4] % 12; if (/pm/i.test(m[6])) h += 12; return new Date(`${m[2]} ${m[1]}, ${m[3]} ${h}:${m[5]}`); };
TK.t = (d, multi) => !d ? "" : (multi ? d.toLocaleDateString("en-GB", { day: "2-digit", month: "short" }) + " " : "") + d.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" });
TK.mins = (a, b) => a && b ? Math.round((b - a) / 60000) : 0;
TK.hm = m => m >= 60 ? `${Math.floor(m / 60)} hr ${m % 60 ? (m % 60) + " min" : ""}`.trim() : `${m} min`;
TK.analyse = rows => {
  const first = rows[0], last = rows[rows.length - 1];
  const multi = first.st && last.en && first.st.toDateString() !== last.en.toDateString();
  const stops = rows.filter(r => /idl|stop/i.test(r.status));
  const dist = rows.reduce((a, r) => a + r.moved, 0);
  const work = r => { const h = r.st && r.st.getHours(); return h >= 8 && h < 18; };
  const flags = [];
  stops.forEach(r => { const m = TK.mins(r.st, r.en); if (m >= 120 && work(r)) flags.push(`Stopped ${TK.hm(m)} at ${r.addr.split(",").slice(0, 3).join(",")} (${TK.t(r.st, multi)}–${TK.t(r.en, multi)})`); });
  for (let i = 1; i < rows.length; i++) { const g = TK.mins(rows[i - 1].en, rows[i].st); if (g >= 60) flags.push(`No GPS record for ${TK.hm(g)} (${TK.t(rows[i - 1].en, multi)}–${TK.t(rows[i].st, multi)})`); }
  const bases = rows.map(r => r.base).filter(x => x);
  if (bases.length && Math.min(...bases) > 50) flags.push(`Base location looks wrong: every record is ${Math.round(Math.min(...bases))}+ km from base. Fix the base in Team Tracker.`);
  const areas = {}; rows.forEach(r => { const p = r.addr.split(",").map(x => x.trim()).filter(x => x && !/^\d|ward|^dhaka( \d+)?$|bangladesh|^[A-Z0-9]{4}\+/i.test(x)); const k = p[p.length - 2] || p[0] || ""; if (k) areas[k] = (areas[k] || 0) + 1; });
  return { first, last, multi, stops, dist, flags, areas: Object.entries(areas).sort((a, b) => b[1] - a[1]).slice(0, 6).map(x => x[0]) };
};
TK.html = (rows, a, off) => {
  const [name, desig, reg] = (rows[0].emp.split("_").concat(["", ""]));
  const ds = a.multi ? `${a.first.st.toLocaleDateString("en-GB", { day: "2-digit", month: "short" })} – ${a.last.en.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}` : a.first.st.toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" });
  return `<div class="rep-zone"><div class="rep-title">${esc(off ? off.name : name)} — Location Track History</div><div class="rep-date">${ds}</div>
  <table class="rep" style="margin:6px 0"><tbody><tr><td><b>Designation</b><br>${esc(off ? off.designation || desig : desig)}</td><td><b>Region / Territory</b><br>${esc(off ? regionName(off.region_id) + " · " + (off.territory || "") : reg)}</td><td><b>Mobile</b><br>${esc(rows[0].mob)}</td><td><b>First – last record</b><br>${TK.t(a.first.st, a.multi)} – ${TK.t(a.last.en, a.multi)}</td><td><b>Records</b><br>${rows.length} (${a.stops.length} stops)</td><td><b>Distance moved</b><br>${a.dist.toFixed(1)} km</td></tr>
  <tr><td colspan="6"><b>Areas visited:</b> ${esc(a.areas.join(", ") || "—")}${a.flags.length ? `<br><b class="red">Check:</b> ${a.flags.map(esc).join("; ")}` : ""}</td></tr></tbody></table>
  <table class="rep"><colgroup><col style="width:3%"><col style="width:9%"><col style="width:9%"><col style="width:8%"><col style="width:8%"><col style="width:6%"><col><col style="width:11%"><col style="width:6%"></colgroup>
  <thead><tr><th>#</th><th>Start</th><th>End</th><th>Duration</th><th>Status</th><th>Moved</th><th>Address</th><th>Lat, Lon</th><th>From base</th></tr></thead>
  <tbody>${rows.map((r, i) => `<tr><td>${i + 1}</td><td>${TK.t(r.st, a.multi)}</td><td>${TK.t(r.en, a.multi)}</td><td>${r.dur === "0" ? "" : esc(r.dur)}</td><td class="${/idl|stop/i.test(r.status) ? "red" : ""}">${esc(r.status)}</td><td>${esc(r.movedTxt)}</td><td>${esc(r.addr)}${r.poi ? ` <i>(${esc(r.poi)})</i>` : ""}</td><td>${esc(r.ll)}</td><td>${esc(r.baseTxt)}</td></tr>`).join("")}</tbody></table>
  <div class="wm">Source: Team Tracker export · HB-Tracker</div></div>`;
};
P.track = async el => {
  el.innerHTML = head("Track History", "Upload one officer's Track History export from Team Tracker to get a clean report and PDF") +
    `<div class="card"><label>Track History file (.xlsx)<input type="file" id="tkF" accept=".xlsx,.xls"></label></div><div id="tkOut"></div>`;
  $("#tkF").onchange = async e => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const rows = TK.parse(await readExcel(f)); if (!rows.length) throw new Error("No records in this file");
      const off = (await q(sb.from("hb_officers").select("name,designation,territory,region_id,sim_number").eq("sim_number", rows[0].mob)))[0];
      const a = TK.analyse(rows), html = TK.html(rows, a, off);
      const nm = (off ? off.name : rows[0].emp.split("_")[0]).replace(/[^\w]+/g, "_");
      const fn = `Track_${nm}_${a.first.st.toISOString().slice(0, 10)}${a.multi ? "_to_" + a.last.en.toISOString().slice(0, 10) : ""}.pdf`;
      $("#tkOut").innerHTML = stats([["Records", rows.length], ["Stops", a.stops.length, "", "#C2410C"], ["Distance", a.dist.toFixed(1) + " km"], ["To check", a.flags.length, a.flags.length ? "see report" : "nothing unusual", a.flags.length ? "#C00000" : "#15803D"]]) +
        (off ? "" : `<div class="warn">Number ${esc(rows[0].mob)} is not in the officer list, so the name comes from the file.</div>`) +
        `<div class="acts" style="margin:10px 0"><button class="primary" id="tkPdf">Download PDF</button></div><div class="card" id="tkRep">${html}</div>`;
      $("#tkPdf").onclick = () => { const box = document.createElement("div"); box.className = "pdfmode"; box.innerHTML = html; html2pdf().set({ margin: [8, 6, 8, 6], filename: fn, image: { type: "jpeg", quality: .95 }, html2canvas: { scale: 2 }, jsPDF: { unit: "mm", format: "a4", orientation: "landscape" }, pagebreak: { mode: ["css", "legacy"], avoid: "tr" } }).from(box).save(); };
    } catch (err) { fail(err); }
  };
};
