  /* ---------- 10. Export a report (PDF or Excel) ---------- */
  var XC = { accent: "#8A1C3F", accentSoft: "#F4E3E9", ink: "#18201E", muted: "#5B6763", line: "#D7DDDA", sunk: "#F3F5F4",
    ok: "#2F7D50", okSoft: "#DDF0E4", part: "#B7790A", partSoft: "#FBEFD3", miss: "#BF3B2A", missSoft: "#F9E0DB", none: "#9AA4A0", noneSoft: "#ECEFEE", over: "#BF3B2A" };
  var PROG_ORDER = ["PG Diploma in Teaching", "M.Ed. Educational Leadership", "M.Ed. Educational Technologies & AI", "PhD in Education"];
  var STATE_COL = { complete: XC.ok, in_progress: XC.part, not_due: XC.none, overdue: XC.miss, not_started: XC.none };
  var ITEM_WORD = { ok: "Meets", part: "Partly", miss: "Missing", none: "Not uploaded", acc: "Accepted" };
  S.exp = { scope: "course", target: "", fmt: "pdf", detail: true, comments: true, resolved: false, assess: true };

  function fileSort(a, b) { return PROG_ORDER.indexOf(a.programShort) - PROG_ORDER.indexOf(b.programShort) || (a.course + a.section).localeCompare(b.course + b.section); }
  function secShown(c) { return c.section && !/,/.test(c.section) && c.section !== c.instructor ? c.section : ""; }
  function fileName(c) { return c.course + (secShown(c) ? " · " + secShown(c) : ""); }
  function termShort(t) { return t === "Full semester" ? "Full" : t || "—"; }
  function courseLabel(c) { return fileName(c) + (c.instructor ? " — " + c.instructor : ""); }
  function expTargets(scope) {
    var cs = semCourses().slice().sort(fileSort);
    if (scope === "course") return cs.map(function (c) { return { v: c.id, l: courseLabel(c) }; });
    if (scope === "instructor") return Array.from(new Set(cs.map(function (c) { return c.instructor; }).filter(Boolean))).sort().map(function (n) { return { v: n, l: n + " (" + cs.filter(function (c) { return c.instructor === n; }).length + ")" }; });
    if (scope === "program") return Array.from(new Set(cs.map(function (c) { return c.programShort; }))).sort(function (a, b) { return PROG_ORDER.indexOf(a) - PROG_ORDER.indexOf(b); }).map(function (p) { return { v: p, l: p }; });
    return [{ v: "all", l: "All programs, " + semLabel() }];
  }
  function expFiles() {
    var e = S.exp, cs = semCourses().slice().sort(fileSort);
    if (e.scope === "course") return cs.filter(function (c) { return c.id === e.target; });
    if (e.scope === "instructor") return cs.filter(function (c) { return c.instructor === e.target; });
    if (e.scope === "program") return cs.filter(function (c) { return c.programShort === e.target; });
    return cs;
  }
  function openExport(courseId) {
    if (courseId) { S.exp.scope = "course"; S.exp.target = courseId; }
    renderExport();
  }
  function renderExport() {
    var e = S.exp, tg = expTargets(e.scope);
    if (!tg.some(function (t) { return t.v === e.target; })) e.target = tg[0] ? tg[0].v : "";
    var n = expFiles().length;
    var seg = function (name, val, label) { return '<button type="button" class="btn small expseg" data-expset="' + name + '" data-val="' + val + '" aria-pressed="' + (e[name] === val) + '">' + label + "</button>"; };
    var chk = function (name, label) { return '<label class="note" style="display:flex;gap:8px;align-items:center;color:var(--ink)"><input type="checkbox" data-expchk="' + name + '"' + (e[name] ? " checked" : "") + "> " + label + "</label>"; };
    var body = '<div style="display:grid;gap:16px">' +
      '<div style="display:grid;gap:6px"><b>Report for</b><div class="decide">' + seg("scope", "course", "One course file") + seg("scope", "instructor", "An instructor") + seg("scope", "program", "A program") + seg("scope", "all", "Whole department") + "</div></div>" +
      (e.scope !== "all" ? '<label style="display:grid;gap:6px"><b>' + ({ course: "Course file", instructor: "Instructor", program: "Program" }[e.scope]) + '</b><select id="expTarget" class="field">' + tg.map(function (t) { return '<option value="' + esc(t.v) + '"' + (t.v === e.target ? " selected" : "") + ">" + esc(t.l) + "</option>"; }).join("") + "</select></label>" : "") +
      '<div style="display:grid;gap:6px"><b>Format</b><div class="decide">' + seg("fmt", "pdf", "PDF (formal report)") + seg("fmt", "xlsx", "Excel (workbook)") + "</div></div>" +
      '<div style="display:grid;gap:8px"><b>Include</b>' +
      (e.scope !== "course" ? chk("detail", "A detailed page for each course file (checklist and findings)") : "") +
      chk("assess", "Assessment evidence (sample counts only, no student names)") + chk("comments", "Review comments") + (e.comments ? chk("resolved", "Resolved comments as well") : "") + "</div>" +
      '<div class="checks" style="padding:12px 14px;font-size:.86rem">' + (n ? "<span><b>" + n + "</b> course file" + (n === 1 ? "" : "s") + " from " + esc(semLabel()) + ". " + (e.fmt === "pdf" ? (e.scope === "course" ? "The PDF has the course file details, a chart, the checklist and findings, and a sign-off block." : "The PDF has a summary with charts, an overview table" + (e.detail ? ", a page per course file" : "") + " and a sign-off block.") : "The workbook has a Summary sheet with charts, then Course files, Checklist" + (e.assess ? ", Assessments" : "") + (e.comments ? ", Comments" : "") + " and Findings sheets.") + "</span>" : "No course files match.") + "</div>" +
      '<div><button class="btn primary" type="button" id="expGo"' + (n ? "" : " disabled") + ">Create " + (e.fmt === "pdf" ? "PDF" : "Excel") + " report</button></div></div>";
    sheetModal("Export a report", semLabel(), body);
    S._panelIs = "exp";
    var sel = $("expTarget"); if (sel) sel.style.cssText = "border:1px solid var(--line);background:var(--surface);border-radius:8px;padding:8px 10px;width:100%";
  }

  /* --- report model --- */
  function itemState(sc, k) {
    if (sc.state === "not_started") return { s: "none", note: "Nothing uploaded yet." };
    var it = (sc.items || []).filter(function (i) { return i.k === k; })[0] || { s: "miss", note: "" };
    return { s: it.accepted && it.acceptedFrom !== "ok" ? "acc" : it.s, note: it.note || "" };
  }
  function buildModel() {
    var e = S.exp, files = expFiles(), today = new Date();
    var me = S.config.reviewerName || "Mousa Alhatti", meEmail = S.config.reviewerEmail || "mousa.alhatti@adu.ac.ae";
    var m = { scope: e.scope, files: files, opts: e, sem: semLabel(), generated: fmtDate(today.toISOString().slice(0, 10)) + ", " + today.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }), reviewer: me, reviewerEmail: meEmail,
      lastScan: semLastScan() ? fmtDT(semLastScan().at) : "Not recorded" };
    var cnt = { complete: 0, in_progress: 0, not_due: 0, overdue: 0 }, approved = 0, score = 0, open = 0;
    files.forEach(function (c) { cnt[dstate(c)] = (cnt[dstate(c)] || 0) + 1; if ((c.review || {}).status === "approved") approved++; score += scan(c).score; open += openComments(c.id).length; });
    m.cnt = cnt; m.approved = approved; m.avg = files.length ? Math.round(score / files.length) : 0; m.open = open;
    m.items = CF.SECTIONS.map(function (s) { var r = { label: s.k + ". " + s.short, ok: 0, part: 0, miss: 0, none: 0 }; files.forEach(function (c) { var st = itemState(scan(c), s.k).s; r[st === "acc" ? "ok" : st]++; }); return r; });
    var c0 = files[0];
    if (e.scope === "course" && c0) {
      var sc0 = scan(c0), rv = (c0.review || {}).status || "pending";
      m.kicker = "Course file review report"; m.title = fileName(c0);
      m.subtitle = (c0.instructor || "Instructor not set") + " · " + c0.programShort;
      m.meta = [["Program", c0.programShort], ["Course", c0.course], ["Section / folder", c0.section || "—"], ["Instructor", c0.instructor || "—"], ["Email", c0.email || "—"], ["Term", c0.term || "—"], ["Due date", dueFor(c0) ? fmtDate(dueFor(c0)) : "—"], ["Class size", c0.enrolled ? c0.enrolled + " students" : "Not recorded"], ["Semester", m.sem], ["Reviewer decision", { pending: "Not reviewed yet", approved: "Approved", revise: "Needs revision" }[rv] + ((c0.review || {}).at ? " (" + fmtDate(c0.review.at.slice(0, 10)) + ")" : "")]];
      var met = (sc0.items || []).filter(function (i) { return i.s === "ok"; }).length;
      m.kpis = [["Completion", sc0.score + "%", STATE_COL[dstate(c0)]], ["Items met", (sc0.state === "not_started" ? 0 : met) + " of 7", XC.ink], ["Status", stateLabel(dstate(c0)), STATE_COL[dstate(c0)]], ["Open comments", String(openComments(c0.id).length), openComments(c0.id).length ? XC.miss : XC.ink]];
    } else {
      var insts = Array.from(new Set(files.map(function (c) { return c.instructor; }).filter(Boolean)));
      var progs = Array.from(new Set(files.map(function (c) { return c.programShort; })));
      m.kicker = e.scope === "instructor" ? "Instructor course file report" : e.scope === "program" ? "Program course file report" : "Department course file report";
      m.title = e.scope === "instructor" ? e.target : e.scope === "program" ? e.target : "Department of Education";
      m.subtitle = files.length + " course file" + (files.length === 1 ? "" : "s") + " · " + m.sem;
      m.meta = (e.scope === "instructor" ? [["Instructor", e.target], ["Email", (files.filter(function (c) { return c.email; })[0] || {}).email || "—"], ["Programs", progs.join(", ")]] : e.scope === "program" ? [["Program", e.target], ["Instructors", String(insts.length)]] : [["Department", "Department of Education"], ["Programs", String(progs.length)], ["Instructors", String(insts.length)]])
        .concat([["Course files", String(files.length)], ["Semester", m.sem]]);
      m.kpis = [["Course files", String(files.length), XC.ink], ["Complete", String(cnt.complete), XC.ok], ["In progress", String(cnt.in_progress), XC.part], ["Not due yet", String(cnt.not_due), XC.muted], ["Overdue", String(cnt.overdue), cnt.overdue ? XC.miss : XC.ink], ["Approved", String(approved), XC.accent], ["Avg. completion", m.avg + "%", XC.ink]];
      m.progs = PROG_ORDER.filter(function (p) { return progs.indexOf(p) >= 0; }).map(function (p) { var r = { label: p, complete: 0, in_progress: 0, not_due: 0, overdue: 0 }; files.forEach(function (c) { if (c.programShort === p) r[dstate(c)]++; }); return r; });
    }
    m.meta.push(["Reviewer", me + (meEmail ? " (" + meEmail + ")" : "")], ["Last SharePoint check", m.lastScan], ["Generated", m.generated]);
    m.comments = function (c) {
      var all = S.comments.filter(function (x) { return x.courseId === c.id; }).sort(function (a, b) { return (a.at || "").localeCompare(b.at || ""); });
      var roots = all.filter(function (x) { return !x.parent && (e.resolved || x.kind !== "issue" || x.status !== "resolved"); }), out = [];
      roots.forEach(function (r) { out.push(r); all.filter(function (x) { return x.parent === r.id; }).forEach(function (x) { out.push(x); }); });
      return out;
    };
    return m;
  }

  /* --- charts (drawn on a canvas, saved as PNG) --- */
  function canvasPNG(w, h, draw) {
    var k = 3, cv = document.createElement("canvas"); cv.width = w * k; cv.height = h * k;
    var g = cv.getContext("2d"); g.scale(k, k); g.fillStyle = "#FFFFFF"; g.fillRect(0, 0, w, h); g.textBaseline = "middle"; draw(g);
    return new Promise(function (res) { cv.toBlob(function (b) { b.arrayBuffer().then(function (ab) { res({ bytes: new Uint8Array(ab), w: w, h: h }); }); }, "image/png"); });
  }
  function fnt(g, size, weight, color) { g.font = (weight || 400) + " " + size + "px Arial, Helvetica, sans-serif"; g.fillStyle = color || XC.ink; }
  function chartTitle(g, t, sub) { fnt(g, 14, 700); g.fillText(t, 0, 12); if (sub) { fnt(g, 11, 400, XC.muted); g.fillText(sub, 0, 30); } }
  function chartDonut(m) {
    var segs = [["Complete", m.cnt.complete, XC.ok], ["In progress", m.cnt.in_progress, XC.part], ["Not due yet", m.cnt.not_due, XC.none], ["Overdue", m.cnt.overdue, XC.miss]], tot = m.files.length || 1;
    return canvasPNG(520, 200, function (g) {
      chartTitle(g, "Status of course files", "Share of course files by status");
      var cx = 90, cy = 120, R = 62, r = 40, a = -Math.PI / 2;
      segs.forEach(function (s) { if (!s[1]) return; var b = a + (s[1] / tot) * Math.PI * 2; g.beginPath(); g.moveTo(cx + Math.cos(a) * r, cy + Math.sin(a) * r); g.arc(cx, cy, R, a, b); g.arc(cx, cy, r, b, a, true); g.closePath(); g.fillStyle = s[2]; g.fill(); g.strokeStyle = "#FFF"; g.lineWidth = 2; g.stroke(); a = b; });
      g.textAlign = "center"; fnt(g, 20, 700); g.fillText(Math.round(m.cnt.complete / tot * 100) + "%", cx, cy - 4); fnt(g, 10, 400, XC.muted); g.fillText("complete", cx, cy + 14); g.textAlign = "left";
      segs.forEach(function (s, i) { var y = 72 + i * 30; g.fillStyle = s[2]; g.fillRect(200, y - 7, 14, 14); fnt(g, 13, 600); g.fillText(s[0], 224, y); fnt(g, 13, 400, XC.muted); g.textAlign = "right"; g.fillText(s[1] + "  (" + Math.round(s[1] / tot * 100) + "%)", 500, y); g.textAlign = "left"; });
    });
  }
  function chartItems(m) {
    var rows = m.items, tot = m.files.length || 1, H = 52 + rows.length * 26 + 24;
    return canvasPNG(520, H, function (g) {
      chartTitle(g, "Checklist items across the course files", "How many course files meet each of the seven items");
      var x0 = 120, W = 330;
      rows.forEach(function (r, i) {
        var y = 56 + i * 26; fnt(g, 11, 600); g.fillText(r.label, 0, y + 9);
        var x = x0; [["ok", XC.ok], ["part", XC.part], ["miss", XC.miss], ["none", XC.noneSoft]].forEach(function (p) { var w = r[p[0]] / tot * W; if (w > 0) { g.fillStyle = p[1]; g.fillRect(x, y, w, 18); x += w; } });
        fnt(g, 11, 700, XC.ink); g.fillText(r.ok + " of " + tot, x0 + W + 10, y + 9);
      });
      var ly = H - 10, lx = 0; [["Meets", XC.ok], ["Partly", XC.part], ["Missing", XC.miss], ["Not uploaded", XC.noneSoft]].forEach(function (p) { g.fillStyle = p[1]; g.fillRect(lx, ly - 5, 10, 10); fnt(g, 10, 400, XC.muted); g.fillText(p[0], lx + 14, ly); lx += g.measureText(p[0]).width + 34; });
    });
  }
  function chartScores(m) {
    var fs = m.files.slice().sort(function (a, b) { return scan(b).score - scan(a).score; }), H = 52 + fs.length * 20 + 6;
    return canvasPNG(520, H, function (g) {
      chartTitle(g, "Completion by course file", "Percentage of the model course file in place");
      var x0 = 200, W = 262;
      fs.forEach(function (c, i) {
        var y = 52 + i * 20, sc = scan(c).score; fnt(g, 10.5, 600); var code = c.course.split(/\s+-\s+/)[0], lab = code + (secShown(c) ? " " + secShown(c) : "") + (m.scope !== "instructor" && c.instructor ? " · " + c.instructor.split(" ").slice(-1)[0] : "");
        if (g.measureText(lab).width > x0 - 10) { while (g.measureText(lab + "…").width > x0 - 10 && lab.length > 4) lab = lab.slice(0, -1); lab += "…"; } g.fillText(lab, 0, y + 7);
        g.fillStyle = XC.sunk; g.fillRect(x0, y, W, 14); g.fillStyle = STATE_COL[dstate(c)]; g.fillRect(x0, y, W * sc / 100, 14);
        fnt(g, 10.5, 700); g.fillText(sc + "%", x0 + W + 8, y + 7);
      });
    });
  }
  function chartPrograms(m) {
    var rows = m.progs, mx = Math.max.apply(null, rows.map(function (r) { return r.complete + r.in_progress + r.not_due + r.overdue; }).concat([1])), H = 52 + rows.length * 34 + 24;
    return canvasPNG(520, H, function (g) {
      chartTitle(g, "Course files by program", "Count of course files in each status");
      var x0 = 225, W = 250;
      rows.forEach(function (r, i) {
        var y = 54 + i * 34; fnt(g, 11, 600); g.fillText(r.label, 0, y + 10);
        var x = x0; [["complete", XC.ok], ["in_progress", XC.part], ["not_due", XC.none], ["overdue", XC.miss]].forEach(function (p) { var w = r[p[0]] / mx * W; if (w > 0) { g.fillStyle = p[1]; g.fillRect(x, y, w, 20); if (w > 16) { fnt(g, 10, 700, "#FFFFFF"); g.textAlign = "center"; g.fillText(String(r[p[0]]), x + w / 2, y + 10); g.textAlign = "left"; } x += w; } });
        fnt(g, 11, 700); g.fillText(String(r.complete + r.in_progress + r.not_due + r.overdue), x + 8, y + 10);
      });
      var ly = H - 10, lx = 0; [["Complete", XC.ok], ["In progress", XC.part], ["Not due yet", XC.none], ["Overdue", XC.miss]].forEach(function (p) { g.fillStyle = p[1]; g.fillRect(lx, ly - 5, 10, 10); fnt(g, 10, 400, XC.muted); g.fillText(p[0], lx + 14, ly); lx += g.measureText(p[0]).width + 34; });
    });
  }
  function chartCourse(m) {
    var c = m.files[0], sc = scan(c);
    return canvasPNG(520, 170, function (g) {
      chartTitle(g, "Course file at a glance", "Completion and the seven checklist items");
      var cx = 70, cy = 105, R = 50; g.lineWidth = 12; g.strokeStyle = XC.sunk; g.beginPath(); g.arc(cx, cy, R, 0, Math.PI * 2); g.stroke();
      g.strokeStyle = STATE_COL[dstate(c)]; g.lineCap = "round"; g.beginPath(); g.arc(cx, cy, R, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * sc.score / 100); if (sc.score) g.stroke(); g.lineCap = "butt";
      g.textAlign = "center"; fnt(g, 20, 700); g.fillText(sc.score + "%", cx, cy - 2); fnt(g, 10, 400, XC.muted); g.fillText(stateLabel(dstate(c)).toLowerCase(), cx, cy + 16); g.textAlign = "left";
      CF.SECTIONS.forEach(function (s, i) {
        var st = itemState(sc, s.k).s, col = { ok: XC.ok, acc: XC.ok, part: XC.part, miss: XC.miss, none: XC.noneSoft }[st], x = 160 + (i % 4) * 90, y = 62 + Math.floor(i / 4) * 52;
        g.fillStyle = col; g.fillRect(x, y, 80, 6); fnt(g, 11, 700); g.fillText(s.k + ". " + s.short, x, y + 20); fnt(g, 10, 400, st === "none" ? XC.muted : col === XC.noneSoft ? XC.muted : col); g.fillText(ITEM_WORD[st], x, y + 35);
      });
    });
  }
  async function buildCharts(m) {
    var ch = [];
    if (m.scope === "course") ch.push(await chartCourse(m));
    else { ch.push(await chartDonut(m)); if (m.progs && m.progs.length > 1) ch.push(await chartPrograms(m)); ch.push(await chartItems(m)); ch.push(await chartScores(m)); }
    return ch;
  }
  function logoPNG(px) {
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="' + px + '" height="' + px + '"><path d="M14 6h22l10 10v24H14z" fill="#FFFFFF" stroke="#8A1C3F" stroke-width="3"/><path d="M9 11h22l10 10v24H9z" fill="#FFFFFF" stroke="#8A1C3F" stroke-width="3"/><path d="M4 18h18v12h22c3 0 5 2 5 5v20c0 3-2 5-5 5H9c-3 0-5-2-5-5z" fill="#8A1C3F"/><path d="M14 30h28v22H14z" fill="#FFFFFF"/><path d="M19 41l6 6 12-12" fill="none" stroke="#2F7D50" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    return new Promise(function (res, rej) {
      var img = new Image(); img.onload = function () { var cv = document.createElement("canvas"); cv.width = px; cv.height = px; var g = cv.getContext("2d"); g.drawImage(img, 0, 0, px, px); cv.toBlob(function (b) { b.arrayBuffer().then(function (ab) { res(new Uint8Array(ab)); }); }, "image/png"); };
      img.onerror = rej; img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
    });
  }

  /* --- PDF --- */
  function loadScript(src, test) {
    if (test()) return Promise.resolve(test());
    return new Promise(function (res, rej) { var s = document.createElement("script"); s.src = src; s.onload = function () { test() ? res(test()) : rej(new Error("Library did not load.")); }; s.onerror = function () { rej(new Error("Could not load the PDF library. Check your connection and try again.")); }; document.head.appendChild(s); });
  }
  var CP1252 = "€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ";
  function safe(s) {
    s = String(s == null ? "" : s).replace(/✓|✔/g, "Yes").replace(/✕|✗|×/g, "No").replace(/→/g, "->").replace(/≥/g, ">=").replace(/≤/g, "<=").replace(/[   ]/g, " ").replace(/[\r\t]/g, " ");
    var o = ""; for (var i = 0; i < s.length; i++) { var ch = s[i], cc = s.charCodeAt(i); o += (cc < 256 && cc !== 127 && (cc >= 32 || cc === 10)) || CP1252.indexOf(ch) >= 0 ? ch : "?"; }
    return o;
  }
  function hexRGB(P, h) { var n = parseInt(h.slice(1), 16); return P.rgb((n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255); }

  async function buildPDF(m, charts) {
    var P = await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js", function () { return window.PDFLib; });
    var doc = await P.PDFDocument.create();
    doc.setTitle(safe(m.kicker + " — " + m.title)); doc.setAuthor(safe(m.reviewer || "Folio")); doc.setSubject(safe("Department of Education · " + m.sem)); doc.setCreator("Folio"); doc.setProducer("Folio");
    var F = { r: await doc.embedFont(P.StandardFonts.Helvetica), b: await doc.embedFont(P.StandardFonts.HelveticaBold), i: await doc.embedFont(P.StandardFonts.HelveticaOblique) };
    var logo = await doc.embedPng(await logoPNG(192));
    var sig = null; try { var sd = S.config.signature; if (sd && /^data:image\/png;base64,/.test(sd)) { var bin = atob(sd.split(",")[1]), u8 = new Uint8Array(bin.length); for (var bi = 0; bi < bin.length; bi++) u8[bi] = bin.charCodeAt(bi); sig = await doc.embedPng(u8); } } catch (se) { sig = null; }
    var imgs = []; for (var q = 0; q < charts.length; q++) imgs.push({ img: await doc.embedPng(charts[q].bytes), w: charts[q].w, h: charts[q].h });
    var W = 595.28, H = 841.89, M = 46, CW = W - 2 * M, TOP = H - 74, BOT = 58;
    var C = {}; Object.keys(XC).forEach(function (k) { C[k] = hexRGB(P, XC[k]); }); C.white = P.rgb(1, 1, 1);
    var pg, y;
    function txt(s, x, yy, size, font, color, opt) { pg.drawText(safe(s), Object.assign({ x: x, y: yy, size: size, font: font || F.r, color: color || C.ink }, opt || {})); }
    function tw(s, font, size) { return (font || F.r).widthOfTextAtSize(safe(s), size); }
    function wrap(s, font, size, width) {
      var out = []; safe(s).split("\n").forEach(function (para) {
        var words = para.split(/ +/), line = "";
        words.forEach(function (w) {
          while (tw(w, font, size) > width && w.length > 1) { var cut = w.length; while (cut > 1 && tw(w.slice(0, cut), font, size) > width) cut--; if (line) { out.push(line); line = ""; } out.push(w.slice(0, cut)); w = w.slice(cut); }
          var t = line ? line + " " + w : w; if (tw(t, font, size) > width && line) { out.push(line); line = w; } else line = t;
        });
        out.push(line);
      });
      return out;
    }
    function newPage() {
      pg = doc.addPage([W, H]);
      pg.drawImage(logo, { x: M, y: H - 52, width: 22, height: 22 });
      txt("Folio", M + 29, H - 39, 12, F.b, C.accent); txt("Course file review", M + 29 + tw("Folio", F.b, 12) + 6, H - 39, 9, F.r, C.muted);
      var rt = "Department of Education  ·  " + m.sem; txt(rt, W - M - tw(rt, F.r, 8.5), H - 33, 8.5, F.r, C.muted);
      var rt2 = "College of Arts, Education and Social Sciences"; txt(rt2, W - M - tw(rt2, F.r, 8.5), H - 45, 8.5, F.r, C.muted);
      pg.drawRectangle({ x: M, y: H - 60, width: CW, height: 1.2, color: C.accent });
      y = TOP;
    }
    function ensure(h) { if (y - h < BOT) newPage(); }
    function gap(h) { y -= h; }
    function heading(s, sub) {
      ensure(40 + (sub ? 12 : 0)); gap(8);
      pg.drawRectangle({ x: M, y: y - 13, width: 3, height: 15, color: C.accent });
      txt(s, M + 10, y - 10, 12.5, F.b, C.ink); gap(18);
      if (sub) { wrap(sub, F.r, 8.5, CW).forEach(function (l) { txt(l, M, y - 8, 8.5, F.r, C.muted); gap(11); }); }
      gap(6);
    }
    function para(s, size, font, color, indent) {
      size = size || 9; var ls = wrap(s, font || F.r, size, CW - (indent || 0));
      ls.forEach(function (l) { ensure(size + 4); txt(l, M + (indent || 0), y - size, size, font, color); gap(size + 3.5); });
    }
    function pill(word, s, x, yy, w) {
      var col = { ok: C.ok, acc: C.ok, part: C.part, miss: C.miss, none: C.muted, complete: C.ok, in_progress: C.part, not_due: C.muted, overdue: C.miss, approved: C.ok, revise: C.miss, pending: C.muted }[s] || C.muted;
      var bg = { ok: C.okSoft, acc: C.okSoft, part: C.partSoft, miss: C.missSoft, none: C.noneSoft, complete: C.okSoft, in_progress: C.partSoft, not_due: C.noneSoft, overdue: C.missSoft, approved: C.okSoft, revise: C.missSoft, pending: C.noneSoft }[s] || C.noneSoft;
      var tws = tw(word, F.b, 7.5), pw = Math.min(w, tws + 10);
      pg.drawRectangle({ x: x, y: yy - 3, width: pw, height: 11, color: bg }); txt(word, x + 5, yy, 7.5, F.b, col);
    }
    // table: cols [{h, w, align}], rows: arrays of cells; a cell is text or {t, pill:s} or {dots:[s...]} or {t, color, bold}
    function table(cols, rows, opt) {
      opt = opt || {}; var fs = opt.size || 8.3, lh = fs + 2.6, pad = 4;
      function head() {
        var hh = 0; var hl = cols.map(function (c) { var l = wrap(c.h, F.b, 7.6, c.w - 2 * pad); hh = Math.max(hh, l.length); return l; });
        var rh = hh * 9.5 + 8; ensure(rh + lh + 8);
        pg.drawRectangle({ x: M, y: y - rh, width: CW, height: rh, color: C.accent });
        var x = M; cols.forEach(function (c, i) { hl[i].forEach(function (l, j) { var lx = c.align === "center" ? x + (c.w - tw(l, F.b, 7.6)) / 2 : x + pad; txt(l, lx, y - 11 - j * 9.5, 7.6, F.b, C.white); }); x += c.w; });
        gap(rh);
      }
      head();
      rows.forEach(function (r, ri) {
        var lines = r.map(function (cell, i) { if (cell && typeof cell === "object" && (cell.pill || cell.dots)) return [""]; var ob = cell && typeof cell === "object", t = ob ? cell.t : cell; return wrap(t == null ? "" : String(t), ob && cell.bold ? F.b : F.r, fs, cols[i].w - 2 * pad); });
        var n = Math.max.apply(null, lines.map(function (l) { return l.length; })), rh = n * lh + 7;
        if (y - rh < BOT) { newPage(); head(); }
        if (ri % 2) pg.drawRectangle({ x: M, y: y - rh, width: CW, height: rh, color: C.sunk });
        var x = M;
        r.forEach(function (cell, i) {
          var c = cols[i];
          if (cell && cell.pill) pill(cell.t, cell.pill, x + pad, y - 4 - fs, c.w - 2 * pad);
          else if (cell && cell.dots) { var dx = x + (c.w - (cell.dots.length * 11 - 3)) / 2; cell.dots.forEach(function (s) { var col = { ok: C.ok, acc: C.ok, part: C.part, miss: C.miss, none: C.line }[s]; pg.drawRectangle({ x: dx, y: y - 5 - fs, width: 8, height: 8, color: col }); dx += 11; }); }
          else lines[i].forEach(function (l, j) { var ob = cell && typeof cell === "object", f = ob && cell.bold ? F.b : F.r, col = ob && cell.color ? hexRGB(P, cell.color) : C.ink; var lx = c.align === "center" ? x + (c.w - tw(l, f, fs)) / 2 : c.align === "right" ? x + c.w - pad - tw(l, f, fs) : x + pad; txt(l, lx, y - 3.5 - fs - j * lh, fs, f, col); });
          x += c.w;
        });
        pg.drawRectangle({ x: M, y: y - rh, width: CW, height: 0.5, color: C.line });
        gap(rh);
      });
      gap(opt.after == null ? 10 : opt.after);
    }
    function chartImg(ci) { var w = Math.min(CW, ci.w * 0.86), h = ci.h * w / ci.w; ensure(h + 8); pg.drawImage(ci.img, { x: M, y: y - h, width: w, height: h }); gap(h + 12); }

    // ---- cover / summary ----
    newPage();
    txt(m.kicker.toUpperCase(), M, y - 9, 8.5, F.b, C.accent); gap(16);
    wrap(m.title, F.b, 21, CW).forEach(function (l) { txt(l, M, y - 20, 21, F.b, C.ink); gap(25); });
    txt(m.subtitle, M, y - 10, 10.5, F.r, C.muted); gap(24);
    // meta grid (two columns)
    var half = (CW - 16) / 2, rowsM = Math.ceil(m.meta.length / 2);
    for (var i = 0; i < rowsM; i++) {
      var pair = [m.meta[i], m.meta[i + rowsM]], hmax = 0;
      pair.forEach(function (kv, j) { if (!kv) return; var ls = wrap(kv[1], F.r, 9, half - 98); hmax = Math.max(hmax, ls.length); });
      var rh = hmax * 11.5 + 7;
      pair.forEach(function (kv, j) { if (!kv) return; var x = M + j * (half + 16); txt(kv[0], x, y - 11, 8, F.b, C.muted); wrap(kv[1], F.r, 9, half - 98).forEach(function (l, k) { txt(l, x + 96, y - 11 - k * 11.5, 9, F.r, C.ink); }); pg.drawRectangle({ x: x, y: y - rh, width: half, height: 0.5, color: C.line }); });
      gap(rh);
    }
    gap(14);
    // KPI tiles
    var kn = m.kpis.length, kw = (CW - (kn - 1) * 6) / kn;
    m.kpis.forEach(function (k, j) {
      var x = M + j * (kw + 6); pg.drawRectangle({ x: x, y: y - 50, width: kw, height: 50, color: C.sunk });
      pg.drawRectangle({ x: x, y: y - 50, width: 2.5, height: 50, color: hexRGB(P, k[2] === XC.ink ? XC.line : k[2]) });
      var vs = String(k[1]), sz = 17; while (tw(vs, F.b, sz) > kw - 14 && sz > 9) sz -= 1;
      txt(vs, x + 9, y - 24, sz, F.b, hexRGB(P, k[2])); wrap(k[0], F.r, 7.8, kw - 14).forEach(function (l, li) { txt(l, x + 9, y - 37 - li * 9, 7.8, F.r, C.muted); });
    });
    gap(66);
    if (imgs.length) { heading(m.scope === "course" ? "Summary" : "Analysis"); imgs.forEach(chartImg); }

    // overview table
    if (m.scope !== "course") {
      ensure(150);
      heading("Course files overview", "Columns 1–7 follow the model course file: 1 Checklist, 2 Syllabus, 3 Materials, 4 Assessment, 5 Attendance, 6 CRC, 7 SET. Green meets the model, amber is partly there, red is missing, grey is not uploaded yet.");
      var showInst = m.scope !== "instructor";
      var cols = [{ h: "Course file", w: showInst ? 100 : 140 }].concat(showInst ? [{ h: "Instructor", w: 80 }] : []).concat([{ h: "Term", w: 40 }, { h: "Due", w: 58 }, { h: "Status", w: 60 }, { h: "Done", w: 30, align: "right" }, { h: "Items 1–7", w: 80, align: "center" }, { h: "Review", w: 40 }]);
      var tot = cols.reduce(function (a, c) { return a + c.w; }, 0); cols[cols.length - 1].w += CW - tot;
      table(cols, m.files.map(function (c) {
        var sc = scan(c), st = dstate(c), rv = (c.review || {}).status || "pending";
        return [{ t: fileName(c), bold: true }].concat(showInst ? [c.instructor || "—"] : []).concat([termShort(c.term), dueFor(c) ? fmtDate(dueFor(c)) : "—", { t: stateLabel(st), pill: st }, { t: sc.score + "%", bold: true }, { dots: CF.SECTIONS.map(function (s) { return itemState(sc, s.k).s; }) }, { approved: "Approved", revise: "Revise", pending: "—" }[rv]]);
      }));
    }

    // per course file detail
    var detail = m.scope === "course" || m.opts.detail;
    if (detail) m.files.forEach(function (c, fi) {
      var sc = scan(c), st = dstate(c), nd = need(c);
      if (m.scope !== "course") { newPage(); txt("COURSE FILE " + (fi + 1) + " OF " + m.files.length, M, y - 9, 8, F.b, C.accent); gap(14); wrap(fileName(c), F.b, 16, CW).forEach(function (l) { txt(l, M, y - 15, 16, F.b, C.ink); gap(20); }); para([c.instructor || "Instructor not set", c.programShort, c.term || "Term not set", dueFor(c) ? "Due " + fmtDate(dueFor(c)) : "", stateLabel(st) + " · " + sc.score + "%", { pending: "Not reviewed yet", approved: "Approved by the reviewer", revise: "Needs revision" }[(c.review || {}).status || "pending"]].filter(Boolean).join("   ·   "), 8.8, F.r, C.muted); gap(4); }
      heading("Checklist against the model course file");
      table([{ h: "#", w: 18, align: "center" }, { h: "Item", w: 128 }, { h: "Result", w: 66 }, { h: "Finding", w: CW - 212 }], CF.SECTIONS.map(function (s) {
        var it = itemState(sc, s.k); return [String(s.k), { t: s.label, bold: true }, { t: ITEM_WORD[it.s], pill: it.s }, it.s === "none" && !isOverdue(c) ? "Nothing uploaded yet. Not due until " + (dueFor(c) ? fmtDate(dueFor(c)) : "the end of the term") + "." : it.note];
      }));
      heading("Department rules");
      para("Graded samples: " + nd + " High, " + nd + " Medium and " + nd + " Low per assessment " + (c.enrolled ? "(class of " + c.enrolled + " students)." : "(class size not recorded, minimum of 1 applied)."), 9);
      var sy = sc.syllabus;
      para("Syllabus: must state the section and campus. " + (sy ? (sy.read ? (sy.file || "Current syllabus") + " — campus " + (sy.campus ? "found" : "not found") + ", section " + (sy.section ? "found" : "not found") + ", office hours " + (sy.office ? "found" : "not found") + ", faculty name " + (sy.faculty ? "found" : "not found") + "." + (c.accepted && c.accepted[2] ? " Accepted by the reviewer." : "") : "The current syllabus could not be read.") : "Not checked yet."), 9);
      gap(4);
      if (m.opts.assess && (sc.assessments || []).length) {
        heading("Assessment evidence", "Counts of graded samples in each level folder. Student work is counted, never listed. Samples needed per level: " + nd + ".");
        var yn = function (b) { return b ? { t: "Yes", color: XC.ok, bold: true } : { t: "No", color: XC.miss, bold: true }; };
        var cnt = function (n) { return { t: String(n || 0), color: (n || 0) >= nd ? XC.ok : XC.miss, bold: true }; };
        table([{ h: "Assessment", w: CW - 290 }, { h: "Instrument", w: 58, align: "center" }, { h: "Key / rubric", w: 58, align: "center" }, { h: "High", w: 58, align: "center" }, { h: "Medium", w: 58, align: "center" }, { h: "Low", w: 58, align: "center" }],
          sc.assessments.map(function (a) { return [{ t: a.name, bold: true }, yn(a.ins), yn(a.key), cnt(a.h), cnt(a.m), cnt(a.l)]; }));
      }
      if ((sc.issues || []).length) { heading("Found by the automatic check"); sc.issues.forEach(function (x) { para("•  " + x, 9, F.r, C.ink, 4); }); gap(6); }
      if (m.opts.comments) {
        var cms = m.comments(c);
        heading("Review comments");
        if (!cms.length) para("No comments recorded.", 9, F.i, C.muted);
        else table([{ h: "Date", w: 56 }, { h: "Folder", w: 72 }, { h: "From", w: 78 }, { h: "Status", w: 66 }, { h: "Comment", w: CW - 272 }], cms.map(function (x) {
          return [fmtDate((x.at || "").slice(0, 10)), x.parent ? "Reply" : x.sec ? x.sec + ". " + CF.SECTIONS[x.sec - 1].short : "Whole file", personName(x.by) + (x.role === "lead" ? " (reviewer)" : " (instructor)"), x.kind === "issue" && !x.parent ? { t: { open: "Open", fixed: "Says fixed", resolved: "Resolved" }[x.status] || "", pill: { open: "miss", fixed: "part", resolved: "ok" }[x.status] || "none" } : "", x.text || ""];
        }));
      }
    });

    // sign-off
    ensure(110); gap(16);
    pg.drawRectangle({ x: M, y: y, width: CW, height: 0.8, color: C.line }); gap(18);
    txt("REVIEW SIGN-OFF", M, y - 8, 8.5, F.b, C.accent); gap(22);
    var sw = (CW - 24) / 3;
    if (sig) { ensure(80); gap(26); }
    [["Reviewer" + (m.reviewerEmail ? " · " + m.reviewerEmail : ""), m.reviewer || ""], ["Signature", ""], ["Date", sig ? m.generated.split(",")[0] : ""]].forEach(function (s, j) {
      var x = M + j * (sw + 12); txt(s[1], x, y - 22, 10, F.r, C.ink);
      if (j === 1 && sig) { var sh = 46, swd = sig.width * sh / sig.height; if (swd > sw - 8) { swd = sw - 8; sh = sig.height * swd / sig.width; } pg.drawImage(sig, { x: x + 4, y: y - 30, width: swd, height: sh }); } pg.drawRectangle({ x: x, y: y - 28, width: sw, height: 0.8, color: C.muted }); txt(s[0], x, y - 40, 8, F.r, C.muted);
    });
    gap(50);
    para("This report reflects the course file folders on SharePoint as last checked by Folio (" + m.lastScan + ") and the reviewer's decisions and comments recorded in Folio. It contains no student names.", 8, F.i, C.muted);

    // footers
    var pages = doc.getPages();
    pages.forEach(function (p, i) {
      pg = p;
      pg.drawRectangle({ x: M, y: 40, width: CW, height: 0.5, color: C.line });
      var fl = "Folio · " + m.kicker + " · Generated " + m.generated; txt(fl, M, 28, 7.5, F.r, C.muted);
      var pn = "Page " + (i + 1) + " of " + pages.length; txt(pn, W - M - tw(pn, F.r, 7.5), 28, 7.5, F.r, C.muted);
    });
    return new Blob([await doc.save()], { type: "application/pdf" });
  }

  /* --- Excel (styled workbook written directly, charts as pictures) --- */
  var CRC_T = (function () { var t = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zip(files) {
    var enc = new TextEncoder(), parts = [], central = [], off = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = typeof f.data === "string" ? enc.encode(f.data) : f.data, crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30)); h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true); h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var c = new DataView(new ArrayBuffer(46)); c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true); c.setUint16(12, 0, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + data.length;
    });
    var csize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var e = new DataView(new ArrayBuffer(22)); e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, csize, true); e.setUint32(16, off, true);
    return new Blob(parts.concat(central, [new Uint8Array(e.buffer)]), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  function xe(s) { return String(s == null ? "" : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function colL(i) { var s = ""; i++; while (i) { var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; }
  // Style ids (see stylesXML)
  var XS = { title: 1, sub: 2, head: 3, body: 4, pct: 5, ok: 6, part: 7, miss: 8, none: 9, label: 10, kpiN: 11, kpiL: 12, center: 13, section: 14, val: 15, bold: 16 };
  function stylesXML() {
    var font = function (sz, col, b, i) { return "<font>" + (b ? "<b/>" : "") + (i ? "<i/>" : "") + '<sz val="' + sz + '"/><color rgb="FF' + col.slice(1) + '"/><name val="Calibri"/><family val="2"/></font>'; };
    var fill = function (col) { return '<fill><patternFill patternType="solid"><fgColor rgb="FF' + col.slice(1) + '"/><bgColor indexed="64"/></patternFill></fill>'; };
    var fonts = [font(11, XC.ink), font(18, XC.accent, 1), font(10, XC.muted), font(10, "#FFFFFF", 1), font(10, XC.ink), font(10, XC.ok, 1), font(10, XC.part, 1), font(10, XC.miss, 1), font(10, XC.muted), font(10, XC.ink, 1), font(20, XC.accent, 1), font(13, XC.ink, 1)];
    var fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>', fill(XC.accent), fill(XC.okSoft), fill(XC.partSoft), fill(XC.missSoft), fill(XC.noneSoft), fill(XC.sunk)];
    var bc = 'rgb="FF' + XC.line.slice(1) + '"';
    var borders = ["<border><left/><right/><top/><bottom/><diagonal/></border>", '<border><left style="thin"><color ' + bc + '/></left><right style="thin"><color ' + bc + '/></right><top style="thin"><color ' + bc + '/></top><bottom style="thin"><color ' + bc + "/></bottom><diagonal/></border>", '<border><left/><right/><top/><bottom style="medium"><color rgb="FF' + XC.accent.slice(1) + '"/></bottom><diagonal/></border>'];
    var xf = function (f, fi, b, al, num) { return '<xf numFmtId="' + (num || 0) + '" fontId="' + f + '" fillId="' + fi + '" borderId="' + b + '" xfId="0"' + (f ? ' applyFont="1"' : "") + (fi ? ' applyFill="1"' : "") + (b ? ' applyBorder="1"' : "") + (num ? ' applyNumberFormat="1"' : "") + (al ? ' applyAlignment="1"><alignment ' + al + "/></xf>" : "/>"); };
    var top = 'vertical="top" wrapText="1"', cen = 'horizontal="center" vertical="center" wrapText="1"';
    var xfs = [xf(0, 0, 0), xf(1, 0, 0, 'vertical="center"'), xf(2, 0, 0, 'vertical="center"'), xf(3, 2, 1, 'vertical="center" wrapText="1"'), xf(4, 0, 1, top), xf(9, 0, 1, 'horizontal="center" vertical="top"', 9), xf(5, 3, 1, cen), xf(6, 4, 1, cen), xf(7, 5, 1, cen), xf(8, 6, 1, cen), xf(9, 0, 0, 'vertical="top"'), xf(10, 7, 0, 'horizontal="center" vertical="center"'), xf(2, 7, 0, 'horizontal="center" vertical="center" wrapText="1"'), xf(4, 0, 1, 'horizontal="center" vertical="top" wrapText="1"'), xf(11, 0, 2, 'vertical="center"'), xf(4, 0, 0, 'vertical="top" wrapText="1"'), xf(9, 0, 1, top)];
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="' + fonts.length + '">' + fonts.join("") + '</fonts><fills count="' + fills.length + '">' + fills.join("") + '</fills><borders count="' + borders.length + '">' + borders.join("") + '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join("") + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  }
  function XSheet(name) { this.name = name; this.rows = []; this.merges = []; this.cols = []; this.images = []; this.freeze = 0; this.filter = null; }
  XSheet.prototype.row = function (cells, ht) { this.rows.push({ cells: cells, ht: ht }); return this.rows.length; };
  XSheet.prototype.xml = function (imgRel) {
    var h = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>';
    h += '<sheetViews><sheetView workbookViewId="0" showGridLines="0"' + (this.freeze ? "" : ' tabSelected="0"') + ">" + (this.freeze ? '<pane ySplit="' + this.freeze + '" topLeftCell="A' + (this.freeze + 1) + '" activePane="bottomLeft" state="frozen"/>' : "") + "</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>";
    if (this.cols.length) h += "<cols>" + this.cols.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join("") + "</cols>";
    h += "<sheetData>";
    this.rows.forEach(function (r, ri) {
      h += '<row r="' + (ri + 1) + '"' + (r.ht ? ' ht="' + r.ht + '" customHeight="1"' : "") + ">";
      (r.cells || []).forEach(function (c, ci) {
        if (c == null) return; if (typeof c !== "object") c = { v: c };
        var ref = colL(ci) + (ri + 1), s = c.s ? ' s="' + c.s + '"' : "";
        if (c.v === "" || c.v == null) h += '<c r="' + ref + '"' + s + "/>";
        else if (typeof c.v === "number") h += '<c r="' + ref + '"' + s + "><v>" + c.v + "</v></c>";
        else h += '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + xe(c.v) + "</t></is></c>";
      });
      h += "</row>";
    });
    h += "</sheetData>";
    if (this.filter) h += '<autoFilter ref="' + this.filter + '"/>';
    if (this.merges.length) h += '<mergeCells count="' + this.merges.length + '">' + this.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join("") + "</mergeCells>";
    h += '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>';
    if (imgRel) h += '<drawing r:id="' + imgRel + '"/>';
    return h + "</worksheet>";
  };
  function drawingXML(imgs) {
    var h = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    imgs.forEach(function (im, i) {
      h += '<xdr:oneCellAnchor><xdr:from><xdr:col>' + im.col + "</xdr:col><xdr:colOff>" + (im.colOff || 0) + "</xdr:colOff><xdr:row>" + im.row + "</xdr:row><xdr:rowOff>" + (im.rowOff || 0) + '</xdr:rowOff></xdr:from><xdr:ext cx="' + Math.round(im.w * 9525) + '" cy="' + Math.round(im.h * 9525) + '"/>' +
        '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 2) + '" name="' + xe(im.name) + '" descr="' + xe(im.name) + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId' + (i + 1) + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + Math.round(im.w * 9525) + '" cy="' + Math.round(im.h * 9525) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
    });
    return h + "</xdr:wsDr>";
  }
  function writeXLSX(sheets, meta) {
    var files = [], media = 0, ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Default Extension="png" ContentType="image/png"/>', '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'];
    var wbRels = [];
    sheets.forEach(function (sh, i) {
      var n = i + 1, rel = null;
      if (sh.images.length) {
        rel = "rId1";
        files.push({ name: "xl/worksheets/_rels/sheet" + n + ".xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + n + '.xml"/></Relationships>' });
        files.push({ name: "xl/drawings/drawing" + n + ".xml", data: drawingXML(sh.images) });
        files.push({ name: "xl/drawings/_rels/drawing" + n + ".xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sh.images.map(function (im, j) { media++; files.push({ name: "xl/media/image" + media + ".png", data: im.bytes }); return '<Relationship Id="rId' + (j + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image' + media + '.png"/>'; }).join("") + "</Relationships>" });
        ct.push('<Override PartName="/xl/drawings/drawing' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
      }
      files.push({ name: "xl/worksheets/sheet" + n + ".xml", data: sh.xml(rel) });
      ct.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      wbRels.push('<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>');
    });
    wbRels.push('<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');
    var defNames = sheets.map(function (sh, i) { return sh.filter ? '<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">\'' + xe(sh.name) + "'!$" + sh.filter.replace(/(\D+)(\d+):(\D+)(\d+)/, "$1$$$2:$$$3$$$4") + "</definedName>" : ""; }).join("");
    files.unshift(
      { name: "[Content_Types].xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + ct.join("") + "</Types>" },
      { name: "_rels/.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>' },
      { name: "docProps/core.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + xe(meta.title) + "</dc:title><dc:creator>" + xe(meta.author || "Folio") + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString().slice(0, 19) + "Z</dcterms:created></cp:coreProperties>" },
      { name: "xl/workbook.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' + sheets.map(function (sh, i) { return '<sheet name="' + xe(sh.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join("") + "</sheets>" + (defNames ? "<definedNames>" + defNames + "</definedNames>" : "") + "</workbook>" },
      { name: "xl/_rels/workbook.xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels.join("") + "</Relationships>" },
      { name: "xl/styles.xml", data: stylesXML() });
    return zip(files);
  }
  function stCell(word, s) { return { v: word, s: { ok: XS.ok, acc: XS.ok, part: XS.part, miss: XS.miss, none: XS.none, complete: XS.ok, in_progress: XS.part, not_due: XS.none, overdue: XS.miss }[s] || XS.center }; }
  async function buildXLSX(m, charts) {
    var logo = await logoPNG(96), sheets = [];
    // Summary
    var su = new XSheet("Summary"); su.cols = [7, 24, 14, 14, 14, 14, 14, 14, 14]; su.images.push({ name: "Folio logo", bytes: logo, col: 0, row: 0, colOff: 30000, rowOff: 50000, w: 40, h: 40 });
    su.row([null, { v: "Folio · " + m.kicker, s: XS.title }], 34); su.merges.push("B1:I1");
    su.row([null, { v: "Abu Dhabi University · College of Arts, Education and Social Sciences · Department of Education", s: XS.sub }], 18); su.merges.push("B2:I2");
    su.row([]);
    su.row([null, { v: m.title, s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }], 24);
    su.row([null, { v: m.subtitle, s: XS.sub }]);
    su.row([]);
    m.meta.forEach(function (kv) { su.row([null, { v: kv[0], s: XS.label }, { v: kv[1], s: XS.val }]); su.merges.push("C" + su.rows.length + ":I" + su.rows.length); });
    su.row([]);
    var kr = [null], kl = [null]; m.kpis.forEach(function (k, i) { kr.push({ v: k[1], s: XS.kpiN }); kl.push({ v: k[0], s: XS.kpiL }); });
    su.row(kr, 34); su.row(kl, 30); su.row([]);
    su.row([null, { v: m.scope === "course" ? "At a glance" : "Analysis", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }, { v: "", s: XS.section }], 24);
    var r0 = su.rows.length + 1;
    charts.forEach(function (c, i) { var w = 640, h = Math.round(c.h * w / c.w); su.images.push({ name: "Chart " + (i + 1), bytes: c.bytes, col: 1, row: r0 - 1, rowOff: 60000, w: w, h: h }); var rowsNeeded = Math.ceil((h + 20) / 20); for (var k = 0; k < rowsNeeded; k++) su.row([]); r0 += rowsNeeded; });
    sheets.push(su);
    // Course files
    var cf = new XSheet("Course files"); cf.cols = [26, 30, 22, 24, 28, 12, 13, 9, 9, 13, 8, 12, 12, 12, 12, 12, 12, 12, 15, 10]; cf.freeze = 1;
    var hdr = ["Program", "Course file", "Section / folder", "Instructor", "Email", "Term", "Due date", "Students", "Samples per level", "Status", "Done", "1 Checklist", "2 Syllabus", "3 Materials", "4 Assessment", "5 Attendance", "6 CRC", "7 SET", "Reviewer decision", "Open comments"];
    cf.row(hdr.map(function (h) { return { v: h, s: XS.head }; }), 30);
    m.files.forEach(function (c) {
      var sc = scan(c), st = dstate(c), rv = (c.review || {}).status || "pending";
      cf.row([{ v: c.programShort, s: XS.body }, { v: c.course, s: XS.bold }, { v: c.section || "", s: XS.body }, { v: c.instructor || "", s: XS.body }, { v: c.email || "", s: XS.body }, { v: c.term || "", s: XS.center }, { v: dueFor(c) ? fmtDate(dueFor(c)) : "", s: XS.center }, { v: +c.enrolled || "", s: XS.center }, { v: need(c), s: XS.center }, stCell(stateLabel(st), st), { v: sc.score / 100, s: XS.pct }]
        .concat(CF.SECTIONS.map(function (s) { var it = itemState(sc, s.k); return stCell(ITEM_WORD[it.s], it.s); }))
        .concat([stCell({ pending: "Not reviewed", approved: "Approved", revise: "Needs revision" }[rv], { pending: "none", approved: "ok", revise: "miss" }[rv]), { v: openComments(c.id).length, s: XS.center }]));
    });
    cf.filter = "A1:" + colL(hdr.length - 1) + (m.files.length + 1);
    sheets.push(cf);
    // Checklist
    var ck = new XSheet("Checklist"); ck.cols = [30, 24, 34, 14, 80]; ck.freeze = 1;
    ck.row(["Course file", "Instructor", "Item", "Result", "Finding"].map(function (h) { return { v: h, s: XS.head }; }), 22);
    m.files.forEach(function (c) { var sc = scan(c); CF.SECTIONS.forEach(function (s) { var it = itemState(sc, s.k); ck.row([{ v: fileName(c), s: XS.bold }, { v: c.instructor || "", s: XS.body }, { v: s.k + ". " + s.label, s: XS.body }, stCell(ITEM_WORD[it.s], it.s), { v: it.note, s: XS.body }]); }); });
    ck.filter = "A1:E" + (m.files.length * 7 + 1);
    sheets.push(ck);
    if (m.opts.assess) {
      var as = new XSheet("Assessments"); as.cols = [30, 24, 34, 12, 12, 12, 12, 12, 12]; as.freeze = 1;
      as.row(["Course file", "Instructor", "Assessment", "Instrument", "Key / rubric", "High", "Medium", "Low", "Needed per level"].map(function (h) { return { v: h, s: XS.head }; }), 22);
      m.files.forEach(function (c) { var nd = need(c); (scan(c).assessments || []).forEach(function (a) { var yn = function (b) { return stCell(b ? "Yes" : "No", b ? "ok" : "miss"); }, n = function (x) { return { v: x || 0, s: (x || 0) >= nd ? XS.ok : XS.miss }; }; as.row([{ v: fileName(c), s: XS.bold }, { v: c.instructor || "", s: XS.body }, { v: a.name, s: XS.body }, yn(a.ins), yn(a.key), n(a.h), n(a.m), n(a.l), { v: nd, s: XS.center }]); }); });
      if (as.rows.length === 1) as.row([{ v: "No assessment folders found yet.", s: XS.body }]);
      sheets.push(as);
    }
    if (m.opts.comments) {
      var co = new XSheet("Comments"); co.cols = [30, 24, 13, 18, 24, 14, 80]; co.freeze = 1;
      co.row(["Course file", "Instructor", "Date", "Folder", "From", "Status", "Comment"].map(function (h) { return { v: h, s: XS.head }; }), 22);
      m.files.forEach(function (c) { m.comments(c).forEach(function (x) { var stw = x.kind === "issue" && !x.parent ? ({ open: "Open", fixed: "Says fixed", resolved: "Resolved" }[x.status] || "") : ""; co.row([{ v: fileName(c), s: XS.bold }, { v: c.instructor || "", s: XS.body }, { v: fmtDate((x.at || "").slice(0, 10)), s: XS.center }, { v: x.parent ? "Reply" : x.sec ? x.sec + ". " + CF.SECTIONS[x.sec - 1].short : "Whole file", s: XS.body }, { v: personName(x.by) + (x.role === "lead" ? " (reviewer)" : " (instructor)"), s: XS.body }, stw ? stCell(stw, { Open: "miss", "Says fixed": "part", Resolved: "ok" }[stw]) : { v: "", s: XS.center }, { v: x.text || "", s: XS.body }]); }); });
      if (co.rows.length === 1) co.row([{ v: "No comments recorded.", s: XS.body }]);
      sheets.push(co);
    }
    var fd = new XSheet("Findings"); fd.cols = [30, 24, 100]; fd.freeze = 1;
    fd.row(["Course file", "Instructor", "Found by the automatic check"].map(function (h) { return { v: h, s: XS.head }; }), 22);
    m.files.forEach(function (c) { (scan(c).issues || []).forEach(function (x) { fd.row([{ v: fileName(c), s: XS.bold }, { v: c.instructor || "", s: XS.body }, { v: x, s: XS.body }]); }); });
    if (fd.rows.length === 1) fd.row([{ v: "Nothing found.", s: XS.body }]);
    sheets.push(fd);
    return writeXLSX(sheets, { title: m.kicker + " — " + m.title, author: m.reviewer });
  }

  async function runExport() {
    var dl = S.downloads; if (!dl) { toast("Downloads are not available in this view."); return; }
    var btn = $("expGo"); if (btn) { btn.disabled = true; btn.textContent = "Preparing…"; }
    try {
      var m = buildModel(); if (!m.files.length) return;
      var charts = await buildCharts(m);
      var blob = S.exp.fmt === "pdf" ? await buildPDF(m, charts) : await buildXLSX(m, charts);
      var name = ("Folio - " + (m.scope === "course" ? m.title + (m.files[0].instructor && m.title.indexOf(m.files[0].instructor) < 0 ? " - " + m.files[0].instructor : "") : m.title) + " - " + m.sem).replace(/[\\/:*?"<>|·]+/g, "-").replace(/\s+/g, " ").replace(/–/g, "-").trim() + (S.exp.fmt === "pdf" ? ".pdf" : ".xlsx");
      S._lastExport = { name: name, size: blob.size, blob: blob };
      await dl.save({ filename: name, data: blob });
      toast("Report saved");
    } catch (e) { if (!(e && e.code === "declined")) toast((e && e.message) ? e.message : "Could not create the report."); if (window.console) console.error(e); }
    finally { var b2 = $("expGo"); if (b2) { b2.disabled = false; b2.textContent = "Create " + (S.exp.fmt === "pdf" ? "PDF" : "Excel") + " report"; } }
  }

