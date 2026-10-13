// Port of gp_auto_watcher v6.1 report logic (same rules, same labels).
(function (root) {
  const CL = "Couldn't locate";
  const PREFIXES = new Set(["md", "mohammad", "mohammed", "muhammad", "mohd", "mst", "mosammat", "mossammat", "most"]);
  const MON = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

  function normMobile(m) {
    if (m === null || m === undefined) return null;
    const d = String(m).replace(/\D/g, "");
    return d.length < 9 ? null : d.slice(-10);
  }
  function normName(raw) {
    if (!raw) return "";
    const s = String(raw).split("\n")[0].trim();
    return s.split("_")[0].trim().toLowerCase().replace(/\s+/g, " ");
  }
  function nameKey(raw) {
    const words = normName(raw).replace(/\./g, " ").split(/[^\p{L}\p{N}_]+/u).filter(Boolean);
    const core = words.filter(w => !PREFIXES.has(w));
    return (core.length ? core : words).slice().sort().filter((w, i, a) => a.indexOf(w) === i).join(" ");
  }
  function parseDt(v) {
    if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds());
    const s = String(v).trim();
    const m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/);
    if (!m) return null;
    let h = +m[4];
    if (m[7]) { const pm = m[7].toLowerCase() === "pm"; if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
    const mo = MON[m[2].toLowerCase()];
    if (mo === undefined) return null;
    return new Date(+m[3], mo, +m[1], h, +m[5], +(m[6] || 0));
  }
  function durMin(d) {
    if (d === null || d === undefined) return 0;
    const s = String(d);
    const h = s.match(/(\d+)\s*hr/), mi = s.match(/(\d+)\s*min/);
    if (!h && !mi) { const n = parseInt(s.trim(), 10); return isNaN(n) ? 0 : n; }
    return (h ? +h[1] * 60 : 0) + (mi ? +mi[1] : 0);
  }
  const dayKey = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

  // rows: array of arrays (row 0 = header), as from XLSX.utils.sheet_to_json(ws,{header:1,raw:true})
  function loadTracker(rows, acc) {
    acc = acc || { byMobile: {}, clNames: {}, dates: new Set() };
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const [name, mobile, start, end, duration, status, , address] = r;
      if (typeof status === "string" && status.toLowerCase().includes("couldn")) {
        const n = normName(name);
        if (n && n !== "no officer assigned") acc.clNames[n] = String(name).split("\n")[0].trim();
        continue;
      }
      const mk = normMobile(mobile);
      if (!mk || !start || !end || !address) continue;
      const s = parseDt(start), e = parseDt(end);
      if (!s || !e) continue;
      acc.dates.add(dayKey(s));
      (acc.byMobile[mk] = acc.byMobile[mk] || []).push({ start: s, end: e, address: String(address), status, dur: durMin(duration), ttName: name });
    }
    return acc;
  }

  // officers: [{report_zone, report_order, report_name, report_designation, report_base, tracker_number}]
  // zones: [{code,title,sort}]
  function buildReport(officers, zones, acc, cfg) {
    const dates = [...acc.dates].sort();
    if (!dates.length) throw new Error("No valid GPS rows found in the file(s).");
    const [Y, M, D] = dates[0].split("-").map(Number);
    const tol = cfg.toleranceMinutes * 60000;
    const clKeys = Object.keys(acc.clNames).map(n => [n, nameKey(n)]);
    const matchedCl = new Set();
    const stats = { matched_rows: 0, cells_filled: 0, cl_tracker_set: 0, no_raw_record: 0, absent_set: 0, inactive_flagged: 0, officers: 0 };
    const audit = [];
    const out = [];
    for (const z of zones.slice().sort((a, b) => a.sort - b.sort)) {
      const list = officers.filter(o => o.report_zone === z.code).sort((a, b) => a.report_order - b.report_order);
      if (!list.length) continue;
      const rows = [];
      let allCl = 0;
      list.forEach((o, idx) => {
        stats.officers++;
        const mk = normMobile(o.tracker_number);
        const recs = mk ? acc.byMobile[mk] : null;
        if (recs) stats.matched_rows++;
        const matches = cfg.checkpoints.map(([label, h, mi]) => {
          const cp = new Date(Y, M - 1, D, h, mi).getTime();
          let rec = null, reason = "";
          if (recs) {
            rec = recs.find(r => r.start.getTime() <= cp && cp <= r.end.getTime()) || null;
            if (rec) reason = "Interval match";
            else {
              let best = null, gap = null;
              for (const r of recs) {
                const g = Math.min(Math.abs(r.start - cp), Math.abs(r.end - cp));
                if (gap === null || g < gap) { gap = g; best = r; }
              }
              if (best && gap <= tol) { rec = best; reason = `Nearest ping (${Math.floor(gap / 60000)} min gap)`; }
            }
          }
          return { label, rec, reason };
        });
        const found = matches.some(m => m.rec);
        let ost = null;
        if (!found) {
          const k = nameKey(o.report_name);
          const hit = k && clKeys.find(([, ck]) => ck === k);
          if (hit) { matchedCl.add(hit[0]); ost = [CL, "Couldn't locate (tracker explicit)", "Tracker explicitly reported CL"]; }
          else if (!recs) ost = ["No raw record", "No raw record in export", "Officer mobile not in export"];
          else ost = ["Absent", "GPS record exists but no ping near checkpoint", `No ping within ±${cfg.toleranceMinutes} min`];
        }
        let inactive = false;
        const cells = matches.map(m => {
          let txt, red = false, raw, res, why = m.reason;
          if (m.rec) {
            txt = m.rec.address; stats.cells_filled++;
            if (m.rec.status && String(m.rec.status).toLowerCase().includes("idl") && m.rec.dur >= cfg.idleMinutes) inactive = true;
            raw = m.rec.status; res = "Address filled";
          } else if (ost) { txt = ost[0]; red = true; raw = ost[1]; res = ost[0]; why = ost[2]; }
          else { txt = CL; red = true; raw = "GPS record exists but no ping near checkpoint"; res = CL; why = `No ping within ±${cfg.toleranceMinutes} min`; }
          audit.push({ Sheet: z.title, Officer: o.report_name, Mobile: fmtMobile(o.tracker_number), Checkpoint: m.label, "Raw Status": raw || "", "Auto Result": res, Reason: why });
          return { txt, red };
        });
        if (cells.every(c => c.txt === CL)) allCl++;
        let remarks = "", rred = false;
        if (ost) { remarks = ost[0]; rred = true; stats[{ [CL]: "cl_tracker_set", "No raw record": "no_raw_record", "Absent": "absent_set" }[ost[0]]]++; }
        else if (inactive) { remarks = "Inactive? (long idle - please review)"; stats.inactive_flagged++; }
        rows.push({ sl: idx + 1, name: o.report_name, desig: o.report_designation || "", base: o.report_base || "", mobile: fmtMobile(o.tracker_number), cells, remarks, rred });
      });
      out.push({ code: z.code, title: z.title, rows, total: rows.length, allCl });
    }
    const unmatchedCl = Object.keys(acc.clNames).filter(n => !matchedCl.has(n)).map(n => acc.clNames[n]);
    return { date: new Date(Y, M - 1, D), zones: out, stats, audit, unmatchedCl };
  }

  function fmtMobile(n) {
    const d = normMobile(n);
    return d ? `0${d.slice(0, 4)}-${d.slice(4)}` : (n || "");
  }

  root.GPReport = { normMobile, normName, nameKey, parseDt, durMin, loadTracker, buildReport, fmtMobile, CL };
})(typeof window !== "undefined" ? window : globalThis);
