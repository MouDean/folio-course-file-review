/* ===== Course File Monitor app ===== */
(function () {
  var M365 = "Microsoft 365";
  var DRIVE = "b!puZM-TFnlkCsdeKHodaislnsRN_Yzj5Hl19_Us3vS4dhhzkd7J0LQ55jyq94WXx0";
  var SITE = "https://studentsaduac.sharepoint.com";
  var LIB = "/sites/CAS_CourseFiles/Shared Documents";
  var DEPT = "Department of Education";
  var TERMS_LONG = ["Term A", "Term B", "Full semester"], SHORT = "Short semester";
  // Winter and Summer are short six-week semesters: no Term A / Term B, one due date for every course file.
  function isShort(k) { return /\/(Winter|Summer)$/.test(k || semKey()); }
  function terms() { return isShort() ? [SHORT] : TERMS_LONG; }
  function eTerm(c) { return isShort(c.semester || semKey()) ? SHORT : c.term; }

  var S = {
    db: null, user: null, mcp: null, isLead: false, me: null, myEmail: "",
    courses: [], comments: [], people: [], config: {}, cfgLoaded: false,
    view: (function () { try { var v = localStorage.getItem("folio-sem"); if (v && v.indexOf("/") > 0) { var p = v.split("/"); return { year: p[0], semester: p[1] }; } } catch (e) {} return null; })(),
    scope: "all", open: null, filters: { prog: "", term: "", state: "", inst: "", text: "" },
    names: {}, scanning: false, kpiFilter: "",
    sort: { key: "course", dir: 1 }, order: [],
    collapsed: (function () { try { return JSON.parse(localStorage.getItem("folio-collapsed") || "{}") || {}; } catch (e) { return {}; } })()
  };

  var $ = function (id) { return document.getElementById(id); };
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function toast(msg) { var t = $("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2600); }
  function fmtDate(iso) { if (!iso) return ""; var d = new Date(iso); if (isNaN(d)) return ""; return d.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" }); }
  function fmtDT(iso) { if (!iso) return ""; var d = new Date(iso); return d.toLocaleDateString("en-GB", { day: "numeric", month: "short" }) + ", " + d.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }); }
  function shortYear(y) { var m = /(20\d\d)\D+20(\d\d)/.exec(y || "2026 - 2027"); return m ? m[1] + "\u2013" + m[2] : (y || ""); }
  function nowISO() { return new Date().toISOString(); }
  // The semester being viewed: the viewer's own choice (remembered in this browser), else the default from Settings.
  var FALL26 = "2026 - 2027/Fall", SEMS = ["Fall", "Winter", "Spring", "Summer"];
  function curYear() { return (S.view && S.view.year) || S.config.year || "2026 - 2027"; }
  function curSem() { return (S.view && S.view.semester) || S.config.semester || "Fall"; }
  function semKey() { return curYear() + "/" + curSem(); }
  function semLabelOf(k) { var p = k.split("/"); return p[1] + " " + shortYear(p[0]); }
  function semLabel() { return semLabelOf(semKey()); }
  function semDue() { var b = S.config.dueBySem || {}, k = semKey(); return b[k] || (k === FALL26 ? S.config.due : null) || {}; }
  function semLastScan() { var b = S.config.lastScanBySem || {}, k = semKey(); return b[k] || (k === FALL26 ? S.config.lastScan : null); }
  // Row ids: Fall 2026-27 keeps its original ids; every other semester gets its own, so semesters never overwrite each other.
  function rid(rel) { return semKey() === FALL26 ? CF.rowId(rel) : CF.rowId(semKey() + "/" + rel); }
  function semOptions() {
    var years = {}; years[S.config.year || "2026 - 2027"] = 1; years[curYear()] = 1;
    S.courses.forEach(function (c) { if (c.semester) years[c.semester.split("/")[0]] = 1; });
    var ys = Object.keys(years).sort(), last = /(20\d\d)\D+(20\d\d)/.exec(ys[ys.length - 1]);
    if (last) ys.push((+last[1] + 1) + " - " + (+last[2] + 1));
    var n = {}; S.courses.forEach(function (c) { if (!c.gone && !c.removed) { var k = c.semester || FALL26; n[k] = (n[k] || 0) + 1; } });
    var out = []; ys.forEach(function (y) { SEMS.forEach(function (s) { var k = y + "/" + s; out.push({ k: k, n: n[k] || 0 }); }); });
    return out;
  }
  function switchSem(k) {
    var p = k.split("/"); S.view = { year: p[0], semester: p[1] };
    try { localStorage.setItem("folio-sem", k); } catch (e) {}
    CF.setYear(p[0]); if (S.open) closeSheet(); S.filters.prog = ""; S.filters.inst = "";
    if (!$("settingsPanel").hidden) fillSettings();
    renderAll(); toast("Showing " + semLabelOf(k));
  }
  // Each semester has its own SharePoint folder, saved in Settings (config.spBySem). Fall 2026-27 keeps its original folder if none is saved.
  function spFor(k) { var b = S.config.spBySem || {}; k = k || semKey(); return b[k] || null; }
  function defaultBase(k) { return LIB + "/" + (k || semKey()) + "/" + DEPT; }
  function basePath() { var sp = spFor(); return sp && sp.path ? sp.path : defaultBase(); }
  function hasFolder() { return !!spFor() || semKey() === FALL26; }
  function driveUri(path) { return "file:///" + DRIVE + "/" + path.slice(LIB.length + 1); }
  // Turn a SharePoint address (browser bar, "Copy link" with /:f:/r/, or a server-relative path) into the folder's path in the course-files library.
  function parseSpLink(raw) {
    var u = String(raw || "").trim(), path = "";
    if (!u) return { err: "Paste the folder address first." };
    if (/sharepoint\.com\/:[a-z]:\/[sg]\//i.test(u)) return { err: "This is a sharing link, which Folio cannot follow. Open the folder in SharePoint and copy the address from the browser bar instead." };
    var m = /[?&]id=([^&#]+)/.exec(u);
    if (m) path = decodeURIComponent(m[1].replace(/\+/g, " "));
    else if ((m = /sharepoint\.com\/:[a-z]:\/r(\/sites\/[^?#]+)/i.exec(u))) path = decodeURIComponent(m[1]);
    else if ((m = /sharepoint\.com(\/sites\/[^?#]+)/i.exec(u))) path = decodeURIComponent(m[1]);
    else if (/^\/sites\//.test(u)) path = u;
    else return { err: "That does not look like a SharePoint folder address." };
    path = path.replace(/\/+$/, "").replace(/\/Forms\/AllItems\.aspx$/i, "");
    if (path.toLowerCase().indexOf(LIB.toLowerCase() + "/") !== 0) return { err: "The folder must be in the course-files library (" + LIB + "). This link points somewhere else." };
    path = LIB + path.slice(LIB.length);
    return { path: path };
  }
  // A link whose path names a different semester (e.g. "Fall" while you are viewing Spring) is refused.
  function semMismatch(path, k) {
    var p = k.split("/"), segs = path.slice(LIB.length + 1).split("/"), other = SEMS.filter(function (x) { return x !== p[1]; });
    var bad = segs.filter(function (x) { return other.some(function (o) { return new RegExp("^" + o + "$", "i").test(x.trim()); }); })[0];
    var yr = segs.filter(function (x) { return /20\d\d\s*-\s*20\d\d/.test(x); })[0];
    if (bad) return "This link is for a " + bad.trim() + " folder, but you are setting up " + semLabelOf(k) + ".";
    if (yr && yr.replace(/\s/g, "") !== p[0].replace(/\s/g, "")) return "This link is for " + yr.trim() + ", but you are setting up " + semLabelOf(k) + ".";
    return "";
  }
  async function checkSpLink(raw) {
    var k = semKey(), r = parseSpLink(raw); if (r.err) return r;
    var mm = semMismatch(r.path, k); if (mm) return { err: mm + " Switch to that semester with the menu at the top, or paste this semester's link." };
    if (!S.mcp) return { err: "Checking the link needs the Microsoft 365 connector." };
    var list; try { list = await listWithRetry(driveUri(r.path)); } catch (e) { return { err: "SharePoint could not open that folder (" + esc((e && e.code) || "error") + "). Check the address and your access." }; }
    var dept = list.folders.filter(function (f) { return f.name.trim().toLowerCase() === DEPT.toLowerCase(); })[0];
    if (dept) { r.path += "/" + dept.name; try { list = await listWithRetry(driveUri(r.path)); } catch (e) { return { err: "SharePoint could not open the Department of Education folder inside that link." }; } }
    var progs = list.folders.map(function (f) { return f.name; });
    if (!progs.length) return { err: "That folder has no sub-folders. Folio expects the Department of Education folder, with one folder per program." };
    return { path: r.path, programs: progs, link: String(raw).trim() };
  }
  function folderUrl(path) { return SITE + "/sites/CAS_CourseFiles/Shared%20Documents/Forms/AllItems.aspx?id=" + encodeURIComponent(path); }
  function banner(kind, html) { var b = $("banner"); if (!html) { b.hidden = true; return; } b.className = "banner " + (kind || ""); b.innerHTML = html; b.hidden = false; }

  /* ---------- derived ---------- */
  function rawScan(c) { return c.scan || { items: [], issues: [], assessments: [], state: "not_started", score: 0, files: 0 }; }
  // Samples required per assessment and per level (High/Medium/Low), from the class size.
  function need(c) { var n = +c.enrolled || 0; return n > 60 ? 3 : n >= 30 ? 2 : 1; }
  var effMemo = new WeakMap();
  function scan(c) {
    var s = rawScan(c);
    if (s.state === "not_started") return s;
    var key = (+c.enrolled || 0) + "|" + JSON.stringify(s.syllabus || null) + "|" + JSON.stringify(c.accepted || null);
    var m = effMemo.get(s); if (m && m.key === key) return m.val;
    var items = (s.items || []).map(function (i) { return Object.assign({}, i); });
    var issues = [], notes = (s.notes || []).slice();
    // Older scans listed unsorted graded files as issues; they are accepted notes now.
    (s.issues || []).forEach(function (x) { if (/graded files? not placed in a High \/ Medium \/ Low folder\.$/.test(x)) notes.push(x + " Accepted as uploaded."); else issues.push(x); });
    var byK = {}; items.forEach(function (i) { byK[i.k] = i; });
    var nd = need(c), short = [];
    (s.assessments || []).forEach(function (a) {
      var unl = a.unl || 0, zeros = (a.h ? 0 : 1) + (a.m ? 0 : 1) + (a.l ? 0 : 1);
      if (zeros > unl) return; // already reported as missing
      var lv = [], gap = 0; [["High", a.h], ["Medium", a.m], ["Low", a.l]].forEach(function (p) { if (p[1] < nd) { lv.push(p[0] + " " + p[1]); gap += nd - p[1]; } });
      if (!lv.length) return;
      // Graded files that are uploaded but not sorted/named by level count towards the shortfall.
      if (unl >= gap) { if (nd > 1) notes.push(a.name + ": " + nd + " samples per level needed; the shortfall is covered by " + unl + " graded file" + (unl === 1 ? "" : "s") + " not sorted by level. Accepted as uploaded."); return; }
      short.push(a.name + ": " + nd + " samples per level needed, found " + lv.join(", ") + (unl ? " (plus " + unl + " not sorted by level)" : "") + ".");
    });
    if (short.length && byK[4]) {
      if (byK[4].s === "ok") byK[4].s = "part";
      byK[4].note = (byK[4].note ? byK[4].note + " " : "") + "Class of " + c.enrolled + " needs " + nd + " High, " + nd + " Medium and " + nd + " Low samples per assessment.";
      short.forEach(function (x) { issues.push("Samples \u2013 " + x); });
    }
    var sy = s.syllabus;
    if (sy && sy.read && byK[2]) {
      var miss = []; if (!sy.campus) miss.push("campus"); if (!sy.section) miss.push("section");
      if (miss.length) {
        if (byK[2].s === "ok") byK[2].s = "part";
        byK[2].note = (byK[2].note ? byK[2].note + " " : "") + "The " + miss.join(" and ") + " could not be found in the current syllabus.";
        issues.push("Syllabus: the current syllabus must state the " + miss.join(" and ") + " (" + (sy.file || "current syllabus") + ").");
      }
    }
    // Reviewer overrides: an item the reviewer accepted counts as done.
    var acc = c.accepted || {};
    var ISSUE_OF = { 2: /^Syllabus/, 4: /^(Samples|Assessment)|not placed in a High/ };
    Object.keys(acc).forEach(function (k) {
      var it = byK[k]; if (!it || !acc[k]) return;
      it.acceptedFrom = it.s; it.accepted = acc[k];
      if (it.s !== "ok") { it.s = "ok"; it.note = "Accepted by the reviewer" + (acc[k].at ? " on " + fmtDate(acc[k].at.slice(0, 10)) : "") + ". " + (it.note || ""); }
      if (ISSUE_OF[k]) issues = issues.filter(function (x) { return !ISSUE_OF[k].test(x); });
    });
    var ok = items.filter(function (i) { return i.s === "ok"; }).length, part = items.filter(function (i) { return i.s === "part"; }).length;
    var val = Object.assign({}, s, { items: items, issues: issues, notes: notes, state: ok === 7 ? "complete" : "in_progress", score: Math.round(((ok + part * 0.5) / 7) * 100) });
    effMemo.set(s, { key: key, val: val });
    return val;
  }

  // What still needs the instructor's attention (drives the reminder lines in emails).
  function sampleShort(c) { var sc = scan(c); if (sc.state === "not_started") return true; var it = (sc.items || []).filter(function (i) { return i.k === 4; })[0]; return !it || it.s !== "ok" || (sc.issues || []).some(function (x) { return /^Samples/.test(x); }); }
  function syllabusShort(c) { var sc = scan(c); if (sc.state === "not_started") return true; return (sc.issues || []).some(function (x) { return /^Syllabus: the current syllabus must state/.test(x); }); }
  // Ready-made reviewer comments. "accept" ones are recorded as notes (not open issues) and accept the item as done.
  var PRESETS = [
    { id: "named", sec: 4, accept: true, text: "Graded samples are uploaded, but the files are named after the students (or another name) rather than as in the model course file. Accepted as correct; no action needed." },
    { id: "place", sec: 0, accept: true, text: "The files are uploaded, but not in the expected sub-folder or not named as in the model course file. Accepted as correct; no action needed." },
    { id: "campus", sec: 2, accept: true, text: "The syllabus template has no Campus field; the section is stated. Accepted as correct; no action needed." },
    { id: "manual", sec: 0, accept: true, text: "Checked manually by the reviewer. Accepted as correct; no action needed." },
    { id: "tidy", sec: 4, accept: false, note: true, text: "Optional, for clarity: please place the graded samples in the High, Medium and Low folders next time." },
    { id: "missing", sec: 0, accept: false, text: "Please upload the missing item(s) listed in this course file." }
  ];
  function openComments(cid) { return S.comments.filter(function (m) { return m.courseId === cid && m.kind === "issue" && m.status !== "resolved"; }); }
  function dueFor(c) { var d = semDue(), t = eTerm(c); return t === SHORT ? d.F || "" : t === "Term A" ? d.A : t === "Term B" ? d.B : t === "Full semester" ? d.F : ""; }
  // Display status: an empty course file is fine until its due date.
  function dstate(c) { var s = scan(c).state; if (s !== "not_started") return s; return isOverdue(c) ? "overdue" : "not_due"; }
  function isOverdue(c) { var d = dueFor(c); return d && scan(c).state !== "complete" && new Date(d + "T23:59:59") < new Date(); }
  // Complete by the checker but not yet approved: waiting for the reviewer.
  function isReady(c) { return scan(c).state === "complete" && ((c.review || {}).status || "pending") !== "approved"; }
  var PROG_ORDER = ["PG Diploma in Teaching", "M.Ed. Educational Leadership", "M.Ed. Educational Technologies & AI", "PhD in Education"];
  function progRank(p) { var i = PROG_ORDER.indexOf(p); return i < 0 ? 99 : i; }
  // "EDC471(E) - EDUC510 - EDUC501 PC" -> main code + cross-listed codes
  function splitCode(name) { var p = String(name || "").split(/\s+[-+]\s+/); if (p.length < 2 || !p.slice(1).every(function (x) { return /^[A-Za-z]{2,5}\s?\d{2,3}/.test(x); })) return { main: String(name || ""), rest: [] }; return { main: p[0], rest: p.slice(1) }; }
  function daysTo(d) { if (!d) return null; var t = new Date(); t.setHours(0, 0, 0, 0); return Math.round((new Date(d + "T00:00:00") - t) / 86400000); }
  function semCourses() { var k = semKey(); return S.courses.filter(function (c) { return !c.gone && !c.removed && (c.semester || "2026 - 2027/Fall") === k; }); }
  function norm(s) { return String(s || "").toLowerCase().replace(/[^a-z\s]/g, " ").replace(/\s+/g, " ").trim(); }
  function mine(c) {
    if (S.myEmail && c.email && c.email.toLowerCase() === S.myEmail) return true;
    if (!S.myName || !c.instructor) return false;
    var parts = norm(c.instructor).split(" ").filter(function (w) { return w.length > 2; });
    return parts.length > 0 && parts.every(function (w) { return S.myName.indexOf(w) >= 0; });
  }
  function visible() {
    var f = S.filters, t = f.text.trim().toLowerCase();
    return semCourses().filter(function (c) {
      if (S.scope === "mine" && !mine(c)) return false;
      if (f.prog && c.programShort !== f.prog) return false;
      if (f.term) { if (f.term === "-" ? eTerm(c) : eTerm(c) !== f.term) return false; }
      if (f.inst && (c.instructor || "") !== f.inst) return false;
      var st = scan(c).state, rv = (c.review || {}).status || "pending";
      if (f.state === "open" && !openComments(c.id).length) return false;
      if (f.state === "revise" && rv !== "revise") return false;
      if (f.state === "approved" && rv !== "approved") return false;
      if (f.state === "ready" && !isReady(c)) return false;
      if (f.state === "incomplete" && st === "complete") return false;
      if (["complete", "in_progress", "not_due", "overdue"].indexOf(f.state) >= 0 && dstate(c) !== f.state) return false;
      if (t && (c.course + " " + (c.instructor || "") + " " + (c.section || "")).toLowerCase().indexOf(t) < 0) return false;
      return true;
    });
  }

  /* ---------- render: summary ---------- */
  function renderSummary() {
    var base = semCourses().filter(function (c) { return S.scope !== "mine" || mine(c); });
    var n = base.length, comp = 0, prog = 0, none = 0, late = 0, oc = 0, appr = 0, ready = 0;
    base.forEach(function (c) { var s = dstate(c); if (s === "complete") comp++; else if (s === "in_progress") prog++; else if (s === "overdue") late++; else none++; oc += openComments(c.id).length; if ((c.review || {}).status === "approved") appr++; if (isReady(c)) ready++; });
    $("sumTitle").textContent = S.scope === "mine" ? "Your course files" : "Department overview";
    $("sumNote").textContent = n ? Math.round(comp / n * 100) + "% complete · " + appr + " of " + n + " approved" : "";
    var k = [["", n, "Course files", ""], ["complete", comp, "Complete", "ok"], ["in_progress", prog, "In progress", "part"], ["not_due", none, "Not due yet", "none"],
      ["ready", ready, "Ready to review", "ready"], ["approved", appr, "Approved", "appr"], ["overdue", late, "Overdue", "late"], ["open", oc, "Open comments", "kcm"]];
    $("kpis").innerHTML = k.map(function (x) { return '<button type="button" class="kpi ' + x[3] + (x[0] && !x[1] ? " zero" : "") + '" data-kpi="' + x[0] + '" aria-pressed="' + (S.filters.state === x[0] && x[0] !== "" ? "true" : "false") + '"' + (x[0] ? ' title="Show only: ' + x[2].toLowerCase() + '"' : ' title="Show all course files"') + "><b>" + x[1] + "</b><span>" + x[2] + "</span></button>"; }).join("");
    var progs = {};
    base.forEach(function (c) { var p = progs[c.programShort] = progs[c.programShort] || { ok: 0, part: 0, none: 0, n: 0, appr: 0 }; var s = scan(c).state; p.n++; if (s === "complete") p.ok++; else if (s === "in_progress") p.part++; else p.none++; if ((c.review || {}).status === "approved") p.appr++; });
    var keys = Object.keys(progs).sort(function (a, b) { return progRank(a) - progRank(b); });
    $("bars").innerHTML = keys.length ? keys.map(function (p) {
      var v = progs[p], w = function (x) { return (x / v.n * 100).toFixed(2) + "%"; };
      return '<button type="button" class="barrow" data-prog="' + esc(p) + '" aria-pressed="' + (S.filters.prog === p) + '" title="' + v.ok + " complete, " + v.part + " in progress, " + v.none + " nothing uploaded yet, " + v.appr + ' approved"><span class="lab">' + esc(p) + '</span><span class="pct"><b>' + v.ok + "</b> / " + v.n + " complete" + (v.appr ? " · " + v.appr + " approved" : "") + '</span><span class="stack">' + (v.ok ? '<i class="s-ok" style="width:' + w(v.ok) + '"></i>' : "") + (v.part ? '<i class="s-part" style="width:' + w(v.part) + '"></i>' : "") + "</span></button>";
    }).join("") : '<p class="note">No course files yet.</p>';
    celebrate(base);
  }

  /* ---------- sound and celebration ---------- */
  var AC = null;
  function soundOn() { try { return localStorage.getItem("folio-sound") !== "off"; } catch (e) { return true; } }
  function setSound(on, quiet) { try { localStorage.setItem("folio-sound", on ? "on" : "off"); } catch (e) {} var b = $("btnSound"); b.setAttribute("aria-pressed", on); b.textContent = on ? "Sound on" : "Sound off"; b.title = on ? "Sounds are on. Press to mute." : "Sounds are off. Press to turn them on."; if (on && !quiet) fx("click"); }
  function tone(f, t0, dur, type, vol) {
    var o = AC.createOscillator(), g = AC.createGain(); o.type = type || "sine"; o.frequency.setValueAtTime(f, AC.currentTime + t0);
    g.gain.setValueAtTime(0.0001, AC.currentTime + t0); g.gain.exponentialRampToValueAtTime(vol || 0.12, AC.currentTime + t0 + 0.015); g.gain.exponentialRampToValueAtTime(0.0001, AC.currentTime + t0 + dur);
    o.connect(g); g.connect(AC.destination); o.start(AC.currentTime + t0); o.stop(AC.currentTime + t0 + dur + 0.05);
  }
  // Short synthesized cues (no audio files): success chime, approve, click, error.
  function fx(kind) {
    if (!soundOn()) return;
    try { AC = AC || new (window.AudioContext || window.webkitAudioContext)(); if (AC.state === "suspended") AC.resume(); } catch (e) { return; }
    if (kind === "success") { tone(659.3, 0, 0.35, "sine", 0.14); tone(987.8, 0.12, 0.5, "sine", 0.12); tone(1318.5, 0.24, 0.6, "triangle", 0.05); }
    else if (kind === "approve") { tone(784, 0, 0.18, "sine", 0.1); tone(1174.7, 0.07, 0.28, "sine", 0.08); }
    else if (kind === "error") { tone(329.6, 0, 0.3, "triangle", 0.12); tone(246.9, 0.16, 0.45, "triangle", 0.1); }
    else if (kind === "party") { [523.3, 659.3, 784, 1046.5].forEach(function (f, i) { tone(f, i * 0.08, 0.4, "triangle", 0.08); }); }
    else tone(1400, 0, 0.05, "square", 0.025);
  }
  function scanDone(ok) { fx(ok ? "success" : "error"); }
  function celebrate(base) {
    if (!base.length || !S.db) return;
    var key = "folio-cele-" + semKey(), done; try { done = JSON.parse(localStorage.getItem(key) || "[]"); } catch (e) { done = []; }
    var groups = {}; base.forEach(function (c) { (groups[c.programShort] = groups[c.programShort] || []).push(c); });
    var hits = Object.keys(groups).filter(function (p) { return groups[p].every(function (c) { return scan(c).state === "complete"; }); });
    if (base.every(function (c) { return (c.review || {}).status === "approved"; })) hits.push("__all_approved");
    var fresh = hits.filter(function (h) { return done.indexOf(h) < 0; });
    if (!fresh.length) return;
    try { localStorage.setItem(key, JSON.stringify(done.concat(fresh))); } catch (e) {}
    fx("party");
    toast(fresh.indexOf("__all_approved") >= 0 ? "Every course file in " + semLabel() + " is approved!" : fresh[0] + " is 100% complete!");
  }

  /* ---------- render: due-date strip ---------- */
  function renderDue() {
    var box = $("dueStrip"), d = semDue(), base = semCourses().filter(function (c) { return S.scope !== "mine" || mine(c); });
    var defs = isShort() ? [["F", SHORT]] : [["A", "Term A"], ["B", "Term B"], ["F", "Full semester"]], cards = [];
    defs.forEach(function (x) {
      var date = d[x[0]]; if (!date) return;
      var cs = base.filter(function (c) { return eTerm(c) === x[1]; }); if (!cs.length) return;
      var left = cs.filter(function (c) { return scan(c).state !== "complete"; }).length, days = daysTo(date);
      if (days < -30) return;
      cards.push({ term: x[1], date: date, days: days, n: cs.length, left: left });
    });
    cards.sort(function (a, b) { return a.date.localeCompare(b.date); });
    if (!cards.length) { box.hidden = true; box.innerHTML = ""; return; }
    box.hidden = false;
    box.innerHTML = cards.map(function (x) {
      var dt = new Date(x.date + "T00:00:00"), cls = x.left === 0 ? "done" : x.days < 0 ? "late" : x.days <= 7 ? "soon" : "";
      var when = x.left === 0 ? "All complete" : x.days < 0 ? Math.abs(x.days) + " day" + (Math.abs(x.days) === 1 ? "" : "s") + " overdue" : x.days === 0 ? "Due today" : x.days === 1 ? "Due tomorrow" : "In " + x.days + " days";
      var on = S.filters.term === x.term && S.filters.state === (x.days < 0 ? "overdue" : "incomplete");
      var line = x.left === 0 ? "All " + x.n + " course files are complete." : x.left + " of " + x.n + " course file" + (x.n === 1 ? "" : "s") + " not complete yet" + (x.days < 0 ? "" : ". Click to list them.");
      return '<button type="button" class="due ' + cls + '" data-due="' + esc(x.term) + '" data-late="' + (x.days < 0 ? 1 : 0) + '" aria-pressed="' + on + '"' + (x.left === 0 ? " disabled" : "") + '><span class="cal"><i>' + dt.toLocaleDateString("en-GB", { month: "short" }) + "</i><b>" + dt.getDate() + '</b></span><span class="txt"><b>' + esc(x.term === SHORT ? semLabel() + " (6 weeks)" : x.term) + " due " + esc(fmtDate(x.date)) + "</b><span>" + esc(line) + '</span></span><span class="when">' + when + "</span></button>";
    }).join("");
  }

  /* ---------- render: grid ---------- */
  function cellHTML(it, st) {
    if (st === "not_started") return '<span class="dot none" title="Not started">·</span>';
    var m = { ok: ["ok", "✓"], part: ["part", "!"], miss: ["miss", "✕"] }[it ? it.s : "miss"];
    return '<span class="dot ' + m[0] + '" title="' + esc(it ? it.note : "") + '">' + m[1] + "</span>";
  }
  function stateLabel(s) { return { complete: "Complete", in_progress: "In progress", not_started: "Not started", not_due: "Not due yet", overdue: "Overdue" }[s] || s; }
  var RV_RANK = { revise: 0, pending: 1, approved: 2 };
  function sortVal(c, key) {
    if (key === "inst") return (c.instructor || "~").toLowerCase();
    if (key === "term") return (eTerm(c) || "~") + "|" + (dueFor(c) || "9999");
    if (key === "done") return scan(c).score;
    if (key === "review") { var rv = (c.review || {}).status || "pending"; return isReady(c) ? -1 : RV_RANK[rv]; }
    if (key === "cm") return openComments(c.id).length;
    return (c.course + " " + (c.section || "")).toLowerCase();
  }
  function sorted(list) {
    var k = S.sort.key, dir = S.sort.dir;
    return list.slice().sort(function (a, b) {
      var pa = progRank(a.programShort), pb = progRank(b.programShort); if (pa !== pb) return pa - pb;
      if (a.programShort !== b.programShort) return String(a.programShort).localeCompare(String(b.programShort));
      var va = sortVal(a, k), vb = sortVal(b, k), r = typeof va === "number" ? va - vb : String(va).localeCompare(String(vb));
      if (!r) r = (a.course + a.section).localeCompare(b.course + b.section);
      return r * dir;
    });
  }
  function reviewHTML(c) {
    var rv = (c.review || {}).status || "pending";
    if (rv === "approved") return '<span class="state approved">✓ Approved</span>';
    if (rv === "revise") return '<span class="state revise">Needs revision</span>' + (isReady(c) ? '<div class="csub">Now complete: recheck</div>' : "");
    if (isReady(c)) return '<span class="state ready">Ready to review</span>';
    return '<span class="state pending">Not reviewed</span>';
  }
  function renderGrid() {
    var list = visible();
    var tb = $("rows");
    Array.prototype.forEach.call(document.querySelectorAll(".grid th[data-sort]"), function (th) { th.setAttribute("aria-sort", th.dataset.sort === S.sort.key ? (S.sort.dir > 0 ? "ascending" : "descending") : "none"); th.querySelector(".arr").textContent = th.dataset.sort === S.sort.key && S.sort.dir < 0 ? "▼" : "▲"; });
    S.order = [];
    if (!S.db) { tb.innerHTML = '<tr><td colspan="13" class="empty">Open this page in Claude while signed in to see the course files.</td></tr>'; return; }
    if (!semCourses().length) { tb.innerHTML = '<tr><td colspan="13" class="empty">No course files recorded for ' + esc(semLabel()) + " yet." + (S.isLead ? " Press <b>Scan SharePoint</b> to read the folders." : "") + "</td></tr>"; return; }
    if (!list.length) { tb.innerHTML = '<tr><td colspan="13" class="empty">' + (S.scope === "mine" ? "No course files are linked to your email yet. Switch to <b>All</b> to browse." : "No course files match these filters. " + '<button class="btn small" type="button" id="btnClear2">Clear filters</button>') + "</td></tr>"; return; }
    list = sorted(list);
    var groups = [], gi = {};
    list.forEach(function (c) { var g = gi[c.programShort]; if (!g) { g = gi[c.programShort] = { name: c.programShort, items: [] }; groups.push(g); } g.items.push(c); });
    var html = "";
    groups.forEach(function (g) {
      var open = !S.collapsed[g.name], n = g.items.length, ok = 0, part = 0, ready = 0, late = 0;
      g.items.forEach(function (c) { var s = scan(c).state; if (s === "complete") ok++; else if (s === "in_progress") part++; if (isReady(c)) ready++; if (isOverdue(c)) late++; });
      html += '<tr class="prog"><td colspan="13"><button type="button" class="proghead" data-group="' + esc(g.name) + '" aria-expanded="' + open + '"><svg class="chev" viewBox="0 0 24 24" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg><span class="pn">' + esc(g.name) + '</span><span class="pc">' + ok + " / " + n + ' complete</span><span class="mini" aria-hidden="true"><i style="width:' + (ok / n * 100) + '%;background:var(--ok)"></i><i style="width:' + (part / n * 100) + '%;background:var(--part)"></i></span><span class="chips">' + (ready ? '<span class="state ready">' + ready + " ready to review</span>" : "") + (late ? '<span class="state overdue">' + late + " overdue</span>" : "") + "</span></button></td></tr>";
      if (!open) return;
      g.items.forEach(function (c) {
        S.order.push(c.id);
        var sc = scan(c), st = sc.state, oc = openComments(c.id).length, due = dueFor(c), late = isOverdue(c);
        var items = {}; (sc.items || []).forEach(function (i) { items[i.k] = i; });
        var cells = "", mcells = "";
        if (st === "not_started") {
          cells = '<td class="empty7" colspan="7"><span class="awaiting' + (late ? " late" : "") + '">' + (late ? "Nothing uploaded · overdue since " + esc(fmtDate(due)) : "Nothing uploaded yet") + "</span></td>";
        } else {
          for (var k = 1; k <= 7; k++) { var dh = cellHTML(items[k], st); cells += '<td class="cell">' + dh + "</td>"; mcells += dh; }
          mcells = '<td class="cells-m">' + mcells + "</td>";
        }
        var col = st === "complete" ? "var(--ok)" : st === "in_progress" ? "var(--part)" : "var(--none)";
        var code = splitCode(c.course);
        var sub = st === "complete" ? '<span class="state complete">Complete</span>' : st === "in_progress" ? '<span class="state in_progress">In progress</span>' + (late ? '<span class="state overdue">Overdue</span>' : "") : "";
        if (c.section && !c.instructor) sub = esc(c.section) + " " + sub;
        html += '<tr data-id="' + esc(c.id) + '" tabindex="0"' + (S.open === c.id ? ' class="current"' : "") + ">" +
          '<td class="c-course"><div class="ccode">' + esc(code.main) + "</div>" + (code.rest.length ? '<div class="cx">with ' + esc(code.rest.join(", ")) + "</div>" : "") + (sub ? '<div class="csub">' + sub + "</div>" : "") + "</td>" +
          '<td class="c-inst">' + (c.instructor ? esc(c.instructor) : '<span class="muted">Not set</span>') + "</td>" +
          '<td class="c-term">' + (eTerm(c) ? esc(eTerm(c) === SHORT ? "Short (6 wks)" : eTerm(c)) : '<span class="muted">Term not set</span>') + (due ? "<small>due " + esc(fmtDate(due).replace(/ \d{4}$/, "")) + "</small>" : "") + "</td>" + cells + mcells +
          '<td class="c-done">' + (sc.score ? '<div class="ring" style="--p:' + sc.score + ";--c:" + col + '" title="' + sc.score + '% done"><b>' + sc.score + "%</b></div>" : '<span class="zero">0%</span>') + "</td>" +
          '<td class="c-rev">' + reviewHTML(c) + "</td>" +
          '<td class="c-cm">' + (oc ? '<span class="cm">' + oc + " open</span>" : '<span class="muted">—</span>') + "</td></tr>";
      });
    });
    tb.innerHTML = html;
  }
  function filtersActive() { var f = S.filters; return !!(f.prog || f.term || f.state || f.inst || f.text.trim()); }
  function renderFilters() {
    var cs = semCourses();
    var progs = Array.from(new Set(cs.map(function (c) { return c.programShort; }))).sort(function (a, b) { return progRank(a) - progRank(b); });
    var insts = Array.from(new Set(cs.map(function (c) { return c.instructor; }).filter(Boolean))).sort();
    fillSelect($("fProg"), "All programs", progs, S.filters.prog);
    fillSelect($("fInst"), "All instructors", insts, S.filters.inst);
    var ft = $("fTerm"), want = terms().join("|");
    if (ft.dataset.terms !== want) { ft.innerHTML = '<option value="">All terms</option>' + terms().map(function (t) { return "<option>" + t + "</option>"; }).join("") + (isShort() ? "" : '<option value="-">Term not set</option>'); ft.dataset.terms = want; if (terms().indexOf(S.filters.term) < 0) S.filters.term = ""; }
    ft.value = S.filters.term; $("fState").value = S.filters.state;
    if ($("fText").value !== S.filters.text) $("fText").value = S.filters.text;
    $("btnClear").hidden = !filtersActive();
    var base = cs.filter(function (c) { return S.scope !== "mine" || mine(c); }).length, shown = visible().length;
    $("countLine").textContent = base ? (shown === base ? base + " course files" : "Showing " + shown + " of " + base) : "";
  }
  function clearFilters() { S.filters = { prog: "", term: "", state: "", inst: "", text: "" }; renderAll(); }
  function fillSelect(el, first, vals, cur) { el.innerHTML = '<option value="">' + first + "</option>" + vals.map(function (v) { return "<option" + (v === cur ? " selected" : "") + ">" + esc(v) + "</option>"; }).join(""); }
  function renderHeader() {
    var so = semOptions(), ss = $("semSel"), cur = semKey();
    ss.innerHTML = so.map(function (o) { return '<option value="' + esc(o.k) + '"' + (o.k === cur ? " selected" : "") + ">" + esc(semLabelOf(o.k)) + (o.n ? " (" + o.n + ")" : "") + "</option>"; }).join("");
    var ls = semLastScan(), sl = $("scanLine");
    sl.innerHTML = (ls ? "Checked against SharePoint " + esc(fmtDT(ls.at)) : "Not scanned yet") + (hasFolder() ? "" : ' · <span class="where" style="color:var(--miss)">SharePoint folder not set</span>');
    sl.title = hasFolder() ? "Scans read " + basePath().slice(LIB.length + 1) : "Set this semester's SharePoint folder in Settings";
    $("btnScan").title = sl.title;
    sl.classList.toggle("stale", !ls || (Date.now() - new Date(ls.at)) > 3 * 86400000);
    $("leadActions").hidden = !S.isLead; $("scanBox").hidden = !S.isLead;
    var sp = $("scanProg"), cur = sp.value, pf = programFolders();
    if (sp.options.length !== pf.length + 1) { sp.innerHTML = '<option value="">All programs</option>' + pf.map(function (p) { return '<option value="' + esc(p) + '">' + esc(CF.programShort(p)) + "</option>"; }).join(""); sp.value = cur; }
    $("scopeSeg").hidden = !S.courses.some(mine);
    Array.prototype.forEach.call($("scopeSeg").querySelectorAll("button"), function (b) { b.setAttribute("aria-pressed", b.dataset.scope === S.scope); });
  }
  function renderAll() { renderHeader(); renderFilters(); renderDue(); renderSummary(); renderGrid(); if (S.open && S.open !== "__new" && S.open !== "__panel") renderSheet(); if (S.open === "__panel" && S._panelIs === "rem") renderReminders(); }

  /* ---------- detail sheet ---------- */
  function course(id) { return S.courses.find(function (c) { return c.id === id; }); }
  function markCurrent() { Array.prototype.forEach.call(document.querySelectorAll("#rows tr[data-id]"), function (r) { r.classList.toggle("current", r.dataset.id === S.open); }); }
  function openSheet(id) { var was = S.open && S.open.indexOf("__") !== 0; S.open = id; $("sheet").hidden = false; $("scrim").hidden = false; renderSheet(); markCurrent(); try { if (!was) $("sheet").querySelector(".x").focus(); } catch (e) {} }
  function closeSheet() { var id = S.open; S.open = null; $("sheet").hidden = true; $("scrim").hidden = true; markCurrent(); var r = id && document.querySelector('#rows tr[data-id="' + (window.CSS && CSS.escape ? CSS.escape(id) : id) + '"]'); if (r) try { r.focus(); } catch (e) {} }
  // Previous / next course file in the order the grid shows them.
  function stepSheet(d) { var i = S.order.indexOf(S.open); if (i < 0) return; var j = i + d; if (j < 0 || j >= S.order.length) return; openSheet(S.order[j]); var r = document.querySelector('#rows tr[data-id="' + S.order[j] + '"]'); if (r && r.scrollIntoView) r.scrollIntoView({ block: "nearest" }); }
  // Comments by the reviewer always show the reviewer's name from Settings (the signed-in account may be a shared one).
  function authorName(m) { return m && m.role === "lead" ? (S.config.reviewerName || "Mousa Alhatti") : personName(m && m.by); }
  function personName(uid) { return (S.names[uid] && S.names[uid].name) || "Someone"; }

  function programFolders() {
    var m = {}; S.courses.forEach(function (c) { if (c.program) m[c.program] = 1; });
    ["Professional Postgraduate Diploma in Teaching (English)", "MEd in Educational Leadership", "MEd Educational Technologies and AI", "Doctor of Philosophy in Education"].forEach(function (p) { m[p] = 1; });
    return Object.keys(m);
  }
  function inp(id, val, ph, w) { return '<input id="' + id + '" value="' + esc(val || "") + '" placeholder="' + esc(ph || "") + '" style="border:1px solid var(--line);background:var(--surface);border-radius:8px;padding:7px 9px;width:' + (w || "100%") + '">'; }
  function renderAddSheet() {
    var h = '<div class="sheet-head"><div class="row"><div style="display:grid;gap:4px"><span class="eyebrow">New course file</span><h2 id="dTitle" style="font-size:1.3rem">Add a course file</h2></div><button class="x" type="button" aria-label="Close">×</button></div></div>' +
      '<div class="sheet-body"><div class="settings-grid">' +
      '<label for="nProg">Program<select id="nProg">' + programFolders().map(function (p) { return "<option>" + esc(p) + "</option>"; }).join("") + "</select></label>" +
      '<label for="nCourse">Course folder name' + inp("nCourse", "", "e.g. EDUL680B") + "</label>" +
      '<label for="nSection">Instructor/section folder (optional)' + inp("nSection", "", "e.g. Politis,Yurgos") + "</label>" +
      '<label for="nInst">Instructor' + inp("nInst", "", "Name") + "</label>" +
      '<label for="nEmail">Email' + inp("nEmail", "", "name@adu.ac.ae") + "</label>" +
      '<label for="nTerm">Term<select id="nTerm"><option value="">Not set</option>' + terms().map(function (t) { return "<option" + (isShort() ? " selected" : "") + ">" + t + "</option>"; }).join("") + "</select></label>" +
      '<label for="nStud">Students<input id="nStud" type="number" min="0" placeholder="?"></label></div>' +
      '<p class="note">Use the exact folder names you use (or will use) on SharePoint, so the next scan links this course file to its folder. Folio does not create the SharePoint folder.</p>' +
      '<div class="actions"><button class="btn primary" type="button" id="nCreate">Add course file</button></div></div>';
    $("sheet").innerHTML = h;
  }
  function renderSheet() {
    if (S.open === "__new") return renderAddSheet();
    if (S.open === "__panel") return;
    var c = course(S.open); if (!c) { closeSheet(); return; }
    var sc = scan(c), rv = (c.review || {}).status || "pending";
    var keepDraft = {}; if (S._mailFor === S.open) ["eSubj", "eBody", "eTo", "eCc"].forEach(function (k) { var el = $(k); if (el) keepDraft[k] = el.value; }); ["cText", "cSec", "dCourse", "dSection"].forEach(function (k) { var el = $(k); if (el) keepDraft[k] = el.value; });
    var items = {}; (sc.items || []).forEach(function (i) { items[i.k] = i; });
    var cms = S.comments.filter(function (m) { return m.courseId === c.id; }).sort(function (a, b) { return (a.at || "").localeCompare(b.at || ""); });
    var issuesBySec = {}; cms.forEach(function (m) { if (m.kind === "issue" && m.status !== "resolved") issuesBySec[m.sec || 0] = (issuesBySec[m.sec || 0] || 0) + 1; });
    var lead = S.isLead;
    var termSel = lead ? '<select id="mTerm" aria-label="Term"><option value="">Term not set</option>' + terms().map(function (t) { return "<option" + (eTerm(c) === t ? " selected" : "") + ">" + t + "</option>"; }).join("") + "</select>" : esc(c.term || "Term not set");
    var instEd = lead ? '<div class="edit-grid"><input id="mInst" value="' + esc(c.instructor || "") + '" placeholder="Instructor name" aria-label="Instructor"><input id="mEmail" value="' + esc(c.email || "") + '" placeholder="email@adu.ac.ae" aria-label="Instructor email"><input id="mStud" type="number" min="0" value="' + esc(c.enrolled || "") + '" placeholder="Students" aria-label="Students" title="Students enrolled"><button class="btn small" type="button" id="mSave">Save</button></div>' : '<div class="meta">' + esc(c.instructor || "Instructor not set") + (c.email ? ' · <span class="copyable">' + esc(c.email) + "</span>" : "") + (c.enrolled ? " · " + esc(c.enrolled) + " students" : "") + "</div>";
    var pos = S.order.indexOf(c.id), code = splitCode(c.course);
    var nav = pos >= 0 ? '<button class="navb" type="button" id="sPrev" aria-label="Previous course file" title="Previous (↑)"' + (pos === 0 ? " disabled" : "") + '><svg viewBox="0 0 24 24"><path d="m15 18-6-6 6-6"/></svg></button><span class="pos">' + (pos + 1) + " / " + S.order.length + '</span><button class="navb" type="button" id="sNext" aria-label="Next course file" title="Next (↓)"' + (pos === S.order.length - 1 ? " disabled" : "") + '><svg viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg></button>' : "";

    var h = '<div class="sheet-head"><div class="row"><div style="display:grid;gap:3px;min-width:0"><span class="eyebrow">' + esc(c.programShort) + '</span><h2 id="dTitle" style="font-size:1.3rem">' + esc(code.main) + (c.section ? ' <span class="muted" style="font-weight:500;font-size:.88rem">· ' + esc(c.section) + "</span>" : "") + "</h2>" + (code.rest.length ? '<span class="cx">Cross-listed with ' + esc(code.rest.join(", ")) + "</span>" : "") + '</div><div class="nav">' + nav + '<button class="x" type="button" aria-label="Close" title="Close (Esc)">×</button></div></div>' +
      '<div class="meta"><span class="state ' + dstate(c) + '">' + stateLabel(dstate(c)) + " · " + sc.score + "%</span>" + reviewHTML(c) + "<span>" + termSel + "</span>" + (dueFor(c) ? "<span>Due " + fmtDate(dueFor(c)) + "</span>" : "") + '<a href="' + esc(folderUrl(c.path)) + '" target="_blank" rel="noopener">Open folder in SharePoint ↗</a></div>' +
      (lead ? '<div class="editrow">' + instEd + '<span class="acts"><button class="btn small" type="button" id="btnRescan">Rescan</button><button class="btn small" type="button" id="btnExpCourse">Export…</button></span></div>' : instEd) + "</div>";

    h += '<div class="sheet-body">';
    if (c.scanNote) h += '<div class="banner warn">' + esc(c.scanNote) + "</div>";
    // review decision
    h += '<section><h3>Reviewer decision</h3>' +
      (lead ? '<div class="decide"><button type="button" class="btn ap" data-review="approved" aria-pressed="' + (rv === "approved") + '">✓ Approve</button><button type="button" class="btn rv" data-review="revise" aria-pressed="' + (rv === "revise") + '">Needs revision</button><button type="button" class="btn" data-review="pending" aria-pressed="' + (rv === "pending") + '">Not reviewed</button></div>' : '<div><span class="state ' + rv + '">' + ({ pending: "Not reviewed yet", approved: "Approved", revise: "Needs revision" }[rv]) + "</span></div>") +
      ((c.review || {}).at ? '<span class="note">Updated ' + fmtDT(c.review.at) + "</span>" : "") + "</section>";
    // Department rules
    var nd = need(c);
    h += '<section><h3>Department rules</h3><div class="checks" style="padding:12px 14px;display:grid;gap:6px;font-size:.86rem">' +
      "<div>Graded samples required: <b>" + nd + " High, " + nd + " Medium, " + nd + " Low</b> per assessment " + (c.enrolled ? "(class of " + esc(c.enrolled) + ")" : '<span class="muted">(class size not recorded, so the minimum of 1 is applied)</span>') + "</div>" +
      "<div>Syllabus must state the section and campus" + (sc.syllabus ? (sc.syllabus.read ? " · " + esc(sc.syllabus.file || "") + ": campus " + (sc.syllabus.campus ? "✓" : "✕") + ", section " + (sc.syllabus.section ? "✓" : "✕") + ", office hours " + (sc.syllabus.office ? "✓" : "✕") + ", faculty name " + (sc.syllabus.faculty ? "✓" : "✕") : " · the current syllabus could not be read") : ' <span class="muted">(checked at the next scan)</span>') + "</div>" +
      (c.emailedAt ? "<div>Instructor emailed " + fmtDT(c.emailedAt) + "</div>" : "") + "</div></section>";

    // checklist
    h += '<section><h3>Checklist against the model course file</h3><div class="checks">';
    CF.SECTIONS.forEach(function (s) {
      var it = items[s.k];
      var cls = sc.state === "not_started" ? "none" : (it ? it.s : "miss");
      var sym = { ok: "✓", part: "!", miss: "✕", none: "·" }[cls];
      h += '<div class="ck"><span class="dot ' + cls + '">' + sym + '</span><div><h3>' + s.k + ". " + esc(s.label) + "</h3><p>" + esc(sc.state === "not_started" ? (isOverdue(c) ? "Nothing uploaded yet." : "Nothing uploaded yet. Not due until " + (dueFor(c) ? fmtDate(dueFor(c)) : "the end of the term") + ".") : (it ? it.note : "")) + "</p></div>" +
        '<div style="display:grid;gap:4px;justify-items:end">' + (issuesBySec[s.k] ? '<span class="cm cmc">● ' + issuesBySec[s.k] + "</span>" : "") + (lead ? '<button type="button" class="btn small" data-addc="' + s.k + '">Comment</button>' : "") +
        (lead && it && sc.state !== "not_started" ? (it.accepted ? '<button type="button" class="btn small" data-accept="' + s.k + '" data-on="0" title="Go back to the automatic result">Undo accept</button>' : (it.s !== "ok" ? '<button type="button" class="btn small" data-accept="' + s.k + '" data-on="1" title="Count this item as done (for example, checked manually)">Accept as done</button>' : "")) : "") + "</div></div>";
    });
    h += "</div></section>";

    // assessments
    if ((sc.assessments || []).length) {
      h += '<section><h3>Assessment evidence</h3><div style="overflow-x:auto;background:var(--surface);border:1px solid var(--line);border-radius:12px"><table class="amx"><thead><tr><th>Assessment</th><th>Instrument</th><th>Key / rubric</th><th>High</th><th>Medium</th><th>Low</th></tr></thead><tbody>';
      sc.assessments.forEach(function (a) {
        var yn = function (b) { return b ? '<span class="yes">✓</span>' : '<span class="no">✕</span>'; };
        var nn = function (n) { return n ? '<span class="yes">✓</span>' : '<span class="no">✕</span>'; };
        h += "<tr><td>" + esc(a.name) + "</td><td>" + yn(a.ins) + "</td><td>" + yn(a.key) + "</td><td>" + nn(a.h) + "</td><td>" + nn(a.m) + "</td><td>" + nn(a.l) + "</td></tr>";
      });
      h += '</tbody></table></div><span class="note">✓ under High / Medium / Low means at least one graded sample is in that folder. Student work is never listed.</span></section>';
    }

    // auto issues
    if ((sc.notes || []).length) {
      h += '<section><h3>Notes <span class="note" style="font-weight:400">(accepted, no action needed)</span></h3><ul class="issues" style="color:var(--muted)">' + sc.notes.map(function (x) { return "<li>" + esc(x) + "</li>"; }).join("") + "</ul></section>";
    }
    if ((sc.issues || []).length) {
      h += '<section><h3>Found by the automatic check</h3><ul class="issues">' + sc.issues.map(function (x) { return "<li>" + esc(x) + "</li>"; }).join("") + "</ul></section>";
    }

    // comments
    h += '<section><h3>Review comments</h3><div class="thread">';
    var roots = cms.filter(function (m) { return !m.parent; });
    if (!roots.length) h += '<p class="note">No comments yet.</p>';
    roots.forEach(function (m) {
      h += commentHTML(m, c);
      cms.filter(function (r) { return r.parent === m.id; }).forEach(function (r) { h += commentHTML(r, c, true); });
    });
    h += "</div>";
    if (S.db) {
      h += '<div class="composer">' + (lead ? '<div class="row"><label for="cPreset" class="note">Quick comment</label><select id="cPreset"><option value="">Choose a ready-made comment…</option>' + PRESETS.map(function (p) { return '<option value="' + p.id + '">' + esc((p.accept ? "Accepted: " : p.note ? "Note: " : "Issue: ") + p.text.replace(/ (Accepted as correct; no action needed\.|Please |Optional, for clarity: )/, " ").slice(0, 80)) + "…</option>"; }).join("") + "</select></div>" : "") + '<div class="row"><label for="cSec" class="note">About</label><select id="cSec"><option value="0">Whole course file</option>' + CF.SECTIONS.map(function (s) { return '<option value="' + s.k + '">' + s.k + ". " + esc(s.short) + "</option>"; }).join("") + '</select></div><textarea id="cText" placeholder="' + (lead ? "Write a comment or issue for the instructor…" : "Write a note to the reviewer…") + '"></textarea><div class="row"><button class="btn primary" type="button" id="cPost">' + (lead ? "Add comment" : "Send note") + "</button></div></div>";
    }
    h += "</section>";

    // email
    if (lead) {
      var mail = buildEmail(c);
      h += '<section><h3>Email the instructor</h3><div class="composer">' +
        '<input id="eTo" value="' + esc(c.email || "") + '" placeholder="instructor@adu.ac.ae" aria-label="To">' +
        '<input id="eCc" value="" placeholder="Cc (optional)" aria-label="Cc">' +
        '<input id="eSubj" value="' + esc(mail.subject) + '" aria-label="Subject">' +
        '<textarea id="eBody" style="min-height:200px" aria-label="Message">' + esc(mail.body) + "</textarea>" +
        '<div class="row"><a class="btn primary" id="eOutlook" target="_blank" rel="noopener" href="#">Open in Outlook ↗</a><button class="btn" type="button" id="eCopy">Copy message</button><button class="btn" type="button" id="eMark">Mark as emailed</button></div>' +
        '<span class="note">Outlook opens a new message with this text in your browser. Check it and press Send there.</span></div></section>';
    }
    if (lead) {
      h += '<section><h3>Course details</h3><div class="settings-grid">' +
        '<label for="dCourse">Course name' + inp("dCourse", c.course) + "</label>" +
        '<label for="dSection">Section / folder label' + inp("dSection", c.section) + "</label>" +
        '<label for="dProg">Program<select id="dProg">' + programFolders().map(function (p) { return "<option" + (p === c.program ? " selected" : "") + ">" + esc(p) + "</option>"; }).join("") + "</select></label></div>" +
        '<div class="actions"><button class="btn" type="button" id="dSave">Save details</button><button class="btn" type="button" id="dRemove" style="color:var(--miss);border-color:var(--miss)">Remove from Folio</button></div>' +
        '<span class="note">Changing the name here does not rename the SharePoint folder. Removing hides this course file from Folio (scans will not bring it back); the SharePoint folder is not touched. You can restore it in Settings.</span></section>';
    }
    h += "</div>";
    $("sheet").innerHTML = h; S._mailFor = S.open;
    Object.keys(keepDraft).forEach(function (k) { var el = $(k); if (el && keepDraft[k]) el.value = keepDraft[k]; });
    syncOutlook();
  }
  function commentHTML(m, c, isReply) {
    var lead = S.isLead, mineC = S.me && m.by === S.me;
    var st = m.kind === "issue" ? '<span class="st-' + m.status + '">' + ({ open: "● Open", fixed: "◐ Instructor says fixed", resolved: "✓ Resolved" }[m.status] || "") + "</span>" : "";
    var sec = m.sec ? "Folder " + m.sec + " · " + CF.SECTIONS[m.sec - 1].short : "Whole course file";
    var acts = "";
    if (m.kind === "issue" && !isReply) {
      if (!lead && m.status === "open") acts += '<button class="btn small" type="button" data-cstat="fixed" data-cid="' + m.id + '">I have fixed this</button>';
      if (lead && m.status !== "resolved") acts += '<button class="btn small" type="button" data-cstat="resolved" data-cid="' + m.id + '">Resolve</button>';
      if (lead && m.status === "resolved") acts += '<button class="btn small" type="button" data-cstat="open" data-cid="' + m.id + '">Reopen</button>';
    }
    if (!isReply && S.db) acts += '<button class="btn small" type="button" data-reply="' + m.id + '">Reply</button>';
    if ((lead || mineC)) acts += '<button class="btn small" type="button" data-cdel="' + m.id + '">Delete</button>';
    return '<div class="cmt' + (isReply ? " reply" : "") + '"><div class="h"><b>' + esc(authorName(m)) + "</b><span>" + (m.role === "lead" ? "Reviewer" : "Instructor") + "</span><span>" + fmtDT(m.at) + "</span>" + (isReply ? "" : "<span>" + sec + "</span>") + st + '</div><div class="t">' + esc(m.text) + "</div>" + (acts ? '<div class="acts">' + acts + "</div>" : "") + '<div class="replybox" data-for="' + m.id + '" hidden></div></div>';
  }

  function salute(name) { name = (name || "").trim(); if (!name) return "Dear colleague,"; if (/^dr\.?\s/i.test(name)) return "Dear Dr. " + name.replace(/^dr\.?\s+/i, "").split(/\s+/)[0] + ","; return "Dear Dr. " + name.split(/\s+/)[0] + ","; }
  function buildEmail(c) {
    var sc = scan(c), lines = [], first = (c.instructor || "").split(" ")[0];
    var sem = semLabel();
    lines.push(salute(c.instructor), "");
    lines.push("Thank you for preparing the course file for " + c.course + (c.section && !c.instructor ? " (" + c.section + ")" : "") + ", " + sem + ". I have reviewed it against the department's model course file.", "");
    var status = sc.state === "not_started" ? "The course file folder is still empty." : "Current status: " + sc.score + "% complete.";
    lines.push(status);
    var miss = (sc.items || []).filter(function (i) { return i.s !== "ok"; });
    if (sc.state !== "not_started" && miss.length) {
      lines.push("", "Items to complete:");
      miss.forEach(function (i) { lines.push("- " + i.k + ". " + CF.SECTIONS[i.k - 1].short + ": " + i.note); });
    }
    if ((sc.issues || []).length) { lines.push("", "Details:"); sc.issues.forEach(function (x) { lines.push("- " + x); }); }
    var oc = openComments(c.id);
    if (oc.length) { lines.push("", "Reviewer comments:"); oc.forEach(function (m) { lines.push("- " + (m.sec ? CF.SECTIONS[m.sec - 1].short + ": " : "") + m.text); }); }
    if (sc.state === "complete" && !oc.length) lines.push("", "All seven sections are in place. Thank you.");
    var nd = need(c);
    var rem = [];
    if (sampleShort(c)) rem.push("upload " + nd + " High, " + nd + " Medium and " + nd + " Low graded sample" + (nd > 1 ? "s" : "") + " for each assessment (files can keep the student's name)");
    if (syllabusShort(c) && !(c.accepted && c.accepted[2])) rem.push("make sure the syllabus states the section and campus");
    if (rem.length) lines.push("", "Reminder: please " + rem.join(", and ") + ".");
    lines.push("", "Course file folder: " + folderUrl(c.path));
    if (S.config.link) lines.push("You can see the full review and reply to comments here: " + S.config.link);
    return { subject: "Course file review: " + c.course + " (" + sem + ")", body: lines.join("\n") };
  }
  function syncOutlook() {
    var a = $("eOutlook"); if (!a) return;
    var to = $("eTo").value.trim(), cc = ($("eCc") ? $("eCc").value.trim() : ""), su = $("eSubj").value, bo = $("eBody").value;
    a.href = "https://outlook.office.com/mail/deeplink/compose?to=" + encodeURIComponent(to) + (cc ? "&cc=" + encodeURIComponent(cc) : "") + "&subject=" + encodeURIComponent(su) + "&body=" + encodeURIComponent(bo);
  }

  /* ---------- writes ---------- */
  function upd(id, data) { return S.db.collection("courses").doc(id).update(data).catch(function (e) { toast("Could not save: " + (e && e.message || e)); throw e; }); }
  function addComment(c, sec, text, parent, kindOverride) {
    var kind = parent ? "reply" : (kindOverride || (S.isLead ? "issue" : "note"));
    var doc = { courseId: c.id, sec: +sec || 0, text: text, by: S.me, role: S.isLead ? "lead" : "instructor", at: nowISO(), kind: kind, status: kind === "issue" ? "open" : "", parent: parent || "" };
    return S.db.collection("comments").add(doc).then(function () { toast(S.isLead ? "Comment added" : "Note sent"); }).catch(function (e) { toast(e && e.code === "not_granted" ? "You need permission to comment. Ask the reviewer to share with 'can use'." : "Could not post the comment."); });
  }

  /* ---------- SharePoint scan via Microsoft 365 connector ---------- */
  var ENTRY = /\s*(.+?) \((folder|file)[^)]*?(\d+) bytes[^)]*\) (file:\/\/\/[^\/\s]+\/[A-Za-z0-9]{34})/g;
  function textOf(res) {
    var t = [];
    (res && res.content || []).forEach(function (b) { if (b && b.type === "text" && b.text) t.push(b.text); });
    if (!t.length && typeof (res && res.payload) === "string") t.push(res.payload);
    return t.join("\n");
  }
  async function listFolder(uri) {
    var res;
    try { res = await S.mcp.callTool(M365, "read_resource", { uri: uri }, { cache: false }); }
    catch (e) {
      // Microsoft Graph throttling arrives as a tool_error whose message says RATE_LIMITED.
      var msg = String((e && e.message) || "") + JSON.stringify((e && e.result) || "");
      if (/RATE_LIMITED|throttl|429|too many requests/i.test(msg)) { var r = new Error("Microsoft Graph asked us to slow down"); r.code = "rate_limited"; throw r; }
      throw e;
    }
    var out = { files: [], folders: [] };
    var txt = textOf(res), m; ENTRY.lastIndex = 0;
    while ((m = ENTRY.exec(txt))) {
      var nm = m[1].trim();
      if (m[2] === "folder") out.folders.push({ name: nm, size: +m[3], uri: m[4] });
      else out.files.push({ name: nm, size: +m[3], uri: m[4] });
    }
    if (!out.files.length && !out.folders.length && /bytes\)/.test(txt)) { var e = new Error("Could not read the folder listing"); e.code = "parse"; throw e; }
    return out;
  }
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  var cooldownUntil = 0, lastCall = 0, GAP = 350;
  async function throttle() {
    for (;;) {
      var now = Date.now(), next = Math.max(cooldownUntil, lastCall + GAP);
      if (now >= next) { lastCall = now; return; }
      await wait(next - now);
    }
  }
  async function listWithRetry(uri) {
    var last;
    for (var i = 0; i < 7; i++) {
      await throttle();
      try { return await listFolder(uri); }
      catch (e) {
        last = e;
        var fatal = ["server_not_connected", "needs_reauth", "not_in_manifest", "approval_required", "blocked_by_policy", "not_granted", "capability_disabled"].indexOf(e && e.code) >= 0;
        if (fatal) throw e;
        if (e && e.code === "rate_limited") {
          // Pause the whole scan, not just this folder, then try again.
          var pause = Math.min(30000, 4000 * Math.pow(2, i));
          cooldownUntil = Math.max(cooldownUntil, Date.now() + pause);
          S.scanPaused = Math.ceil(pause / 1000);
          continue;
        }
        if (e && e.code === "tool_error" && i >= 1) break;
        if (i >= 3) break;
        await wait((e && e.retryAfterMs) || (800 * Math.pow(2, i) + Math.random() * 400));
      }
    }
    throw last;
  }
  // Folders whose contents only need to be "not empty": teaching-material sub-folders and graded High/Medium/Low folders.
  function sizeOnly(parentName) {
    return /^\s*3(?![0-9])|material|teaching|instruction/i.test(parentName || "") && CF.secNum(parentName) !== 4 && !/assess|graded/i.test(parentName);
  }
  function isLevel(name) { return /\bhigh\b|medium|\bmed\b|\blow\b/i.test(name || ""); }
  function allF(n) { var o = (n.files || []).slice(); (n.folders || []).forEach(function (f) { o = o.concat(allF(f)); }); return o; }
  function findSyllabus(node) {
    var s2 = (node.folders || []).filter(function (f) { return CF.secNum(f.name) === 2; })[0] || (node.folders || []).filter(function (f) { return /syllab/i.test(f.name); })[0];
    if (!s2) return null;
    var cur = (s2.folders || []).filter(function (f) { return /current|this|present/i.test(f.name); })[0];
    var pool = (cur ? allF(cur) : (s2.files || [])).filter(function (f) { return f.uri && f.size > 0 && /\.(docx?|pdf)$/i.test(f.name); });
    pool.sort(function (a, b) { return (/\.docx?$/i.test(b.name) ? 1 : 0) - (/\.docx?$/i.test(a.name) ? 1 : 0); });
    return pool[0] || null;
  }
  async function readSyllabus(file, instructor) {
    var res, last;
    for (var i = 0; i < 6; i++) {
      await throttle();
      try { res = await S.mcp.callTool(M365, "read_resource", { uri: file.uri, startPage: 1, endPage: 3 }, { cache: false }); break; }
      catch (e) {
        last = e; var msg = String((e && e.message) || "") + JSON.stringify((e && e.result) || "");
        if (/RATE_LIMITED|throttl|429/i.test(msg)) { cooldownUntil = Math.max(cooldownUntil, Date.now() + Math.min(30000, 4000 * Math.pow(2, i))); continue; }
        return { read: false, file: file.name };
      }
    }
    if (!res) return { read: false, file: file.name };
    var t = textOf(res), last2 = norm(instructor || "").split(" ").filter(function (w) { return w.length > 2; }).pop();
    return {
      read: true, file: file.name,
      campus: /abu\s*dhabi|al\s*ain|dubai/i.test(t),
      section: /\bsection\b|\bsec\.?\s*[:#-]?\s*\d/i.test(t),
      office: /office\s*hours?/i.test(t),
      faculty: last2 ? norm(t).indexOf(last2) >= 0 : false
    };
  }
  async function scanSharePoint(opts) {
    opts = opts || {};
    if (S.scanning) return;
    if (!S.mcp) { banner("err", "Scanning needs the Microsoft 365 connector. Open this page in claude.ai with Microsoft 365 connected."); return; }
    S.scanning = true; $("btnScan").disabled = true;
    var done = 0, queued = 1, reusedN = 0;
    var show = function () { banner("", '<b>Reading SharePoint…</b><span class="note">' + done + " of " + queued + " folders" + (reusedN ? " · " + reusedN + " unchanged course files skipped" : "") + (Date.now() < cooldownUntil ? " · Microsoft asked us to slow down, pausing a few seconds" : "") + '</span><span class="progress"><i style="width:' + Math.round(done / Math.max(queued, 1) * 100) + '%"></i></span>'); };
    show();
    if (!hasFolder()) { banner("warn", "<b>No SharePoint folder is set for " + esc(semLabel()) + ".</b> <span>Open Settings and paste the link to this semester's Department of Education folder, so Folio does not read another semester's folder.</span>"); S.scanning = false; $("btnScan").disabled = false; return; }
    var deptUri = driveUri(basePath());
    var single = opts.courseId ? course(opts.courseId) : null;
    var errors = 0;
    try {
      var tree = { name: DEPT, files: [], folders: [] }, errCodes = {}, startJob;
      if (single && single.path.indexOf(basePath() + "/") !== 0) { var se = new Error("This course file was found in a different SharePoint folder from the one now set for " + semLabel() + ". Run a full scan of the semester."); se.code = "other_folder"; throw se; }
      if (single) {
        // Rescan one course file: build the path down to its folder and start there.
        var relS = single.path.slice(basePath().length + 1), parts = relS.split("/"), node = tree;
        for (var pi = 0; pi < parts.length; pi++) { var nn = { name: parts[pi], files: [], folders: [] }; node.folders.push(nn); node = nn; }
        startJob = { node: node, uri: deptUri + "/" + relS, depth: parts.length, size: 1, rel: relS, isStart: true };
      } else startJob = { node: tree, uri: deptUri, depth: 0, size: 1, rel: "", isStart: true };
      var pending = [startJob], softSkip = {}, failNames = [], sizeByRel = {};
      var byPath = {}; S.courses.forEach(function (c) { if (c.path) byPath[c.path] = c; });
      var active = 0;
      await new Promise(function (resolve, reject) {
        function pump() {
          if (!pending.length && !active) return resolve();
          while (active < 1 && pending.length) {
            var job = pending.shift(); active++;
            (function (job) {
              listWithRetry(job.uri).then(function (r) {
                job.node.files = r.files;
                if (job.level) {
                  // A High/Medium/Low folder: each file or sub-folder inside it is one graded sample.
                  job.node.files = r.files.concat(r.folders.map(function (f) { return { name: "(sample folder)", size: f.size }; }));
                  job.node.folders = []; return;
                }
                var isRootListing = r.folders.filter(function (f) { return CF.secNum(f.name); }).length >= 2;
                job.node.folders = r.folders.map(function (f) {
                  var n = { name: f.name, files: [], folders: [] }, cd = job.depth + 1, rel = job.rel ? job.rel + "/" + f.name : f.name;
                  if (cd === 1 && opts.prog && f.name !== opts.prog) { n.outOfScope = true; return n; }
                  if (cd === 2 || cd === 3) {
                    sizeByRel[rel] = f.size;
                    var ex = byPath[basePath() + "/" + rel];
                    var approved = ex && ex.review && ex.review.status === "approved";
                    if (!opts.full && ex && !ex.gone && !ex.removed && ex.scan && (approved || (f.size > 0 && ex.folderSize === f.size))) { n.reuse = ex.id; reusedN++; return n; }
                  }
                  if (cd === 3 && f.size === 0 && !isRootListing) return n; // empty instructor folder: nothing to read
                  if (job.depth >= 3 && sizeOnly(job.node.name)) { if (f.size > 0) n.files = [{ name: "(contents)", size: f.size }]; return n; }
                  if (job.depth >= 3 && isLevel(f.name) && f.size === 0) return n;
                  pending.push({ node: n, uri: f.uri, depth: cd, size: f.size, level: job.depth >= 3 && isLevel(f.name), rel: rel }); queued++; return n;
                });
              }).catch(function (e) {
                if (job.isStart) { reject(e); return; }
                if (job.size === 0 && job.depth >= 3) return;
                if (job.size === 0 && job.depth === 2) { softSkip[basePath() + "/" + job.rel] = 1; return; }
                if (failNames.length < 4) failNames.push(job.rel.split("/").slice(-2).join(" / ") + (e && e.message ? " (" + String(e.message).slice(0, 80) + ")" : ""));
                errors++; job.node.failed = true; var ec = (e && e.code) || "error"; errCodes[ec] = (errCodes[ec] || 0) + 1;
              }).then(function () { active--; done++; show(); pump(); });
            })(job);
          }
        }
        pump();
      });
      CF.setYear(curYear());
      // Keep only programs that were in scope.
      tree.folders = tree.folders.filter(function (p) { return !p.outOfScope; });
      var rows = CF.rows(tree, basePath());
      rows.forEach(function (r) { r.id = rid(r.path.slice(basePath().length + 1)); });
      if (single) rows = rows.filter(function (r) { return r.path === single.path; });
      function hasFail(n) { return !!(n.failed || (n.folders || []).some(hasFail)); }
      var failedPaths = {};
      (function mark(n, rel, depth) { (n.folders || []).forEach(function (f) { var r2 = rel ? rel + "/" + f.name : f.name; if (depth >= 1 && hasFail(f)) failedPaths[basePath() + "/" + r2] = 1; mark(f, r2, depth + 1); }); })(tree, "", 0);
      function rowFailed(r) { return Object.keys(failedPaths).some(function (p) { return r.path === p || r.path.indexOf(p + "/") === 0 || p.indexOf(r.path + "/") === 0; }); }
      var skipped = 0;
      function soft(r) { return Object.keys(softSkip).some(function (p) { return r.path === p || r.path.indexOf(p + "/") === 0; }); }
      var existing = {}; S.courses.forEach(function (c) { existing[c.id] = c; });
      function isReused(r) { return r.node && r.node.reuse; }
      // Read each current syllabus (only for course files that changed).
      var toRead = rows.filter(function (r) { return !isReused(r) && r.scan.state !== "not_started" && r.node && !rowFailed(r); });
      for (var j = 0; j < toRead.length; j++) {
        var rr = toRead[j], sf = findSyllabus(rr.node);
        banner("", "<b>Reading syllabi…</b><span class=\"note\">" + (j + 1) + " of " + toRead.length + "</span>");
        if (sf) rr.scan.syllabus = await readSyllabus(sf, (existing[rr.id] && existing[rr.id].instructor) || rr.instructorFromFolder);
      }
      var people = {}; S.people.forEach(function (p) { people[(p.name || "").toLowerCase()] = p; });
      var at = nowISO(), col = S.db.collection("courses"), updatedN = 0;
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i], ex = existing[r.id];
        if (isReused(r)) continue;
        if (soft(r)) { if (ex) continue; }
        if (rowFailed(r)) { skipped++; if (ex) await col.doc(r.id).update({ scanNote: "Last scan could not read every folder; showing the previous result." }); continue; }
        if (ex && ex.removed) continue; // removed by the reviewer: leave it hidden
        var relR = r.path.slice(basePath().length + 1);
        var scanData = { gone: false, manual: false, scanNote: "", scan: r.scan, scannedAt: at, path: r.path, program: r.program, programShort: r.programShort, course: r.course, section: r.section, semester: semKey(), folderSize: sizeByRel[relR] != null ? sizeByRel[relR] : (single && ex ? ex.folderSize || 0 : 0) };
        updatedN++;
        if (ex) await col.doc(r.id).update(scanData);
        else {
          var nm = r.instructorFromFolder, p = people[(nm || "").toLowerCase()];
          await col.doc(r.id).set(Object.assign(scanData, { instructor: p ? p.name : nm, email: p ? p.email : "", term: isShort() ? SHORT : /diploma/i.test(r.program) ? "Full semester" : "", review: { status: "pending" } }));
        }
      }
      if (single && !rows.length) { banner("warn", "<b>The folder for this course file was not found on SharePoint.</b> <span>It may have been renamed or moved. Run a program scan to pick up the new name.</span>"); S.scanning = false; $("btnScan").disabled = false; return; }
      // Course files whose folder no longer exists on SharePoint are hidden (kept in the database with their comments).
      if (!errors && !single) {
        var seen = {}; rows.forEach(function (r) { seen[r.id] = 1; });
        for (var g = 0; g < S.courses.length; g++) {
          var oc = S.courses[g];
          if (opts.prog && oc.program !== opts.prog) continue;
          if (!seen[oc.id] && !oc.gone && !oc.manual && (oc.semester || "2026 - 2027/Fall") === semKey()) await col.doc(oc.id).update({ gone: true });
        }
      }
      var cfg = Object.assign({}, S.config);
      if (!single) {
        cfg.lastScanBySem = Object.assign({}, cfg.lastScanBySem || {});
        if (semKey() !== FALL26 && !cfg.lastScanBySem[FALL26] && cfg.lastScan) cfg.lastScanBySem[FALL26] = cfg.lastScan;
        cfg.lastScanBySem[semKey()] = { at: at, folders: done, errors: errors, scope: opts.prog || "All programs", full: !!opts.full };
        cfg.lastScanByProg = Object.assign({}, cfg.lastScanByProg || {});
        (opts.prog ? [opts.prog] : tree.folders.map(function (p) { return p.name; })).forEach(function (p) { cfg.lastScanByProg[p] = at; });
        await S.db.collection("config").doc("settings").set(cfg);
      }
      var codes = Object.keys(errCodes).map(function (k) { return k + " ×" + errCodes[k]; }).join(", ");
      banner(errors ? "warn" : "", errors
        ? "<b>" + (rows.length - skipped) + " of " + rows.length + " course files updated.</b> <span>SharePoint did not answer for " + errors + " folders (" + esc(codes) + "), so " + skipped + " course files keep their previous result. Press Scan again to retry them.</span>" + (failNames.length ? '<span class="note" style="flex-basis:100%">Not read: ' + esc(failNames.join("; ")) + "</span>" : "")
        : single ? "<b>Course file rescanned.</b>"
        : "<b>Scan complete" + (opts.prog ? " for " + esc(CF.programShort(opts.prog)) : "") + ".</b> <span class=\"note\">" + updatedN + " course files checked, " + reusedN + " unchanged or approved and skipped, " + done + " folders read.</span>");
      setTimeout(function () { if (!errors) banner(); }, 7000);
      scanDone(!errors);
    } catch (e) {
      var code = e && e.code;
      var msg = code === "server_not_connected" ? "Add the Microsoft 365 connector in claude.ai Settings → Connectors, then scan again."
        : code === "needs_reauth" ? "Reconnect Microsoft 365 in claude.ai Settings → Connectors, then scan again."
        : code === "not_in_manifest" || code === "approval_required" ? "This page needs your permission to read SharePoint through Microsoft 365. Press Scan again and choose Allow."
        : code === "tool_error" ? (single ? "SharePoint could not open this course file's folder. It may have been renamed or moved." : "SharePoint could not open the folder set for " + esc(semLabel()) + " (" + esc(basePath().slice(LIB.length + 1)) + "). Check the SharePoint folder link in Settings.")
        : "The scan stopped: " + esc(e && e.message || String(e));
      banner("err", msg); scanDone(false);
    }
    S.scanning = false; $("btnScan").disabled = false;
  }

  /* ---------- events ---------- */
  document.addEventListener("click", function (ev) {
    var t = ev.target.closest("button, a, tr[data-id], th[data-sort]"); if (!t) return;
    if (t.matches("tr[data-id]")) return openSheet(t.dataset.id);
    if (t.matches("th[data-sort]")) { var sk = t.dataset.sort; S.sort = S.sort.key === sk ? { key: sk, dir: -S.sort.dir } : { key: sk, dir: sk === "done" || sk === "cm" ? -1 : 1 }; return renderGrid(); }
    if (t.dataset.group) { S.collapsed[t.dataset.group] = !S.collapsed[t.dataset.group]; try { localStorage.setItem("folio-collapsed", JSON.stringify(S.collapsed)); } catch (e) {} return renderGrid(); }
    if (t.dataset.due) { var late = t.dataset.late === "1", stv = late ? "overdue" : "incomplete", on = S.filters.term === t.dataset.due && S.filters.state === stv; S.filters.term = on ? "" : t.dataset.due; S.filters.state = on ? "" : stv; if (!on) { S.collapsed = {}; } return renderAll(); }
    if (t.id === "btnClear" || t.id === "btnClear2") return clearFilters();
    if (t.id === "sPrev") return stepSheet(-1);
    if (t.id === "sNext") return stepSheet(1);
    if (t.dataset.kpi !== undefined) { S.filters.state = t.dataset.kpi === "" ? "" : (S.filters.state === t.dataset.kpi ? "" : t.dataset.kpi); if (t.dataset.kpi === "") { S.filters.prog = ""; S.filters.term = ""; } return renderAll(); }
    if (t.dataset.prog !== undefined) { S.filters.prog = S.filters.prog === t.dataset.prog ? "" : t.dataset.prog; return renderAll(); }
    if (t.dataset.scope) { S.scope = t.dataset.scope; return renderAll(); }
    if (t.id === "btnScan") return scanSharePoint({ prog: $("scanProg").value, full: $("scanFull").checked });
    if (t.id === "btnReport") return downloadReport();
    if (t.id === "btnExport") return openExport();
    if (t.id === "sigRemove") { var nc = Object.assign({}, S.config); delete nc.signature; S.db.collection("config").doc("settings").set(nc).then(function () { toast("Signature removed"); setTimeout(fillSettings, 300); }); return; }
    if (t.id === "btnExpCourse") return openExport(S.open);
    if (t.dataset.expset) { S.exp[t.dataset.expset] = t.dataset.val; if (t.dataset.expset === "scope") S.exp.detail = t.dataset.val === "instructor"; return renderExport(); }
    if (t.id === "expGo") return runExport();
    if (t.id === "btnRemind") return renderReminders();
    if (t.dataset.remcopy) { var gtxt = reminderBody(S._remGroups[+t.dataset.remcopy]); try { navigator.clipboard.writeText(gtxt).then(function () { toast("Copied"); }); } catch (e) {} return; }
    if (t.dataset.remsent) { var gg = S._remGroups[+t.dataset.remsent]; Promise.all(gg.files.map(function (c) { return upd(c.id, { emailedAt: nowISO() }); })).then(function () { toast("Marked as sent"); fx("click"); }); return; }
    if (t.id === "impApply") return applyImport();
    if (t.dataset.impadd) return prefillAdd(S._imp.missing[+t.dataset.impadd]);
    if (t.dataset.impopen) { closeSheet(); return openSheet(t.dataset.impopen); }
    if (t.id === "btnRescan") { var rc = course(S.open); if (rc) scanSharePoint({ courseId: rc.id, full: true }); return; }
    if (t.id === "btnAdd") { S.open = "__new"; $("sheet").hidden = false; $("scrim").hidden = false; renderAddSheet(); return; }
    if (t.id === "nCreate") return createCourse();
    if (t.dataset.restore) { upd(t.dataset.restore, { removed: false }).then(function () { toast("Restored"); fillSettings(); }); return; }
    if (t.id === "btnSound") { setSound(!soundOn()); return; }
    if (t.id === "btnSettings") { var p = $("settingsPanel"); p.hidden = !p.hidden; fillSettings(); return; }
    if (t.id === "btnCloseSettings") { $("settingsPanel").hidden = true; return; }
    if (t.id === "btnSaveSettings") return saveSettings();
    if (t.id === "spCheck") return runSpCheck();
    if (t.classList.contains("x")) return closeSheet();
    var c = course(S.open);
    if (!c) return;
    if (t.dataset.review) { var rvv = t.dataset.review; upd(c.id, { review: { status: rvv, at: nowISO(), by: S.me } }).then(function () { toast(rvv === "approved" ? "Course file approved" : "Decision saved"); fx(rvv === "approved" ? "approve" : "click"); }); return; }
    if (t.id === "dSave") { var pr = $("dProg").value; upd(c.id, { course: $("dCourse").value.trim() || c.course, section: $("dSection").value.trim(), program: pr, programShort: CF.programShort(pr) }).then(function () { toast("Details saved"); }); return; }
    if (t.id === "dRemove") {
      if (t.dataset.confirm !== "1") { t.dataset.confirm = "1"; t.textContent = "Click again to remove"; return; }
      upd(c.id, { removed: true, removedAt: nowISO() }).then(function () { toast("Removed from Folio"); closeSheet(); }); return;
    }
    if (t.id === "mSave") { upd(c.id, { instructor: $("mInst").value.trim(), email: $("mEmail").value.trim(), enrolled: +$("mStud").value || 0 }).then(function () { toast("Saved"); }); return; }
    if (t.dataset.accept) {
      var accs = Object.assign({}, c.accepted || {}), k = t.dataset.accept;
      if (t.dataset.on === "1") accs[k] = { at: nowISO(), by: S.me }; else delete accs[k];
      upd(c.id, { accepted: accs }).then(function () { toast(t.dataset.on === "1" ? "Item accepted as done" : "Back to the automatic result"); });
      return;
    }
    if (t.dataset.addc) { var sel = $("cSec"); if (sel) { sel.value = t.dataset.addc; $("cText").focus(); } return; }
    if (t.id === "cPost") {
      var tx = $("cText").value.trim(); if (!tx) return;
      var pr = S._preset && S._preset.text === tx ? S._preset : null, secv = $("cSec").value;
      addComment(c, secv, tx, null, pr && (pr.accept || pr.note) ? "note" : null).then(function () {
        var el = $("cText"); if (el) el.value = ""; var ps = $("cPreset"); if (ps) ps.value = "";
        if (pr && pr.accept && +secv) { var accs = Object.assign({}, c.accepted || {}); accs[secv] = { at: nowISO(), by: S.me, why: pr.id }; upd(c.id, { accepted: accs }).then(function () { toast("Comment added and item " + secv + " accepted as done"); }); }
        S._preset = null;
      });
      return;
    }
    if (t.dataset.cstat) { S.db.collection("comments").doc(t.dataset.cid).update({ status: t.dataset.cstat, statusAt: nowISO(), statusBy: S.me }).then(function () { toast(t.dataset.cstat === "fixed" ? "Reviewer will recheck this item" : "Updated"); fx(t.dataset.cstat === "resolved" ? "approve" : "click"); }).catch(function () { toast("Could not update the comment."); }); return; }
    if (t.dataset.cdel) {
      if (t.dataset.confirm !== "1") { t.dataset.confirm = "1"; t.textContent = "Confirm delete"; return; }
      S.db.collection("comments").doc(t.dataset.cdel).delete().catch(function () { toast("Could not delete."); }); return;
    }
    if (t.dataset.reply) {
      var box = document.querySelector('.replybox[data-for="' + t.dataset.reply + '"]');
      box.hidden = false; box.innerHTML = '<div class="composer"><textarea placeholder="Write a reply…"></textarea><div class="row"><button class="btn primary small" type="button" data-sendreply="' + t.dataset.reply + '">Send reply</button></div></div>';
      box.querySelector("textarea").focus(); return;
    }
    if (t.dataset.sendreply) { var ta = t.closest(".composer").querySelector("textarea"); var v = ta.value.trim(); if (!v) return; addComment(c, 0, v, t.dataset.sendreply); return; }
    if (t.id === "eCopy") {
      var txt = "To: " + $("eTo").value + ($("eCc").value ? "\nCc: " + $("eCc").value : "") + "\nSubject: " + $("eSubj").value + "\n\n" + $("eBody").value;
      try { navigator.clipboard.writeText(txt).then(function () { toast("Message copied"); }, function () { $("eBody").select(); toast("Select-all done: press Ctrl+C"); }); } catch (e) { $("eBody").select(); }
      return;
    }
    if (t.id === "eMark") { var to = $("eTo").value.trim(); upd(c.id, { emailedAt: nowISO(), email: to || c.email || "" }).then(function () { toast("Recorded as emailed"); fx("click"); }); return; }
  });
  document.addEventListener("keydown", function (ev) {
    var typing = ev.target.matches && ev.target.matches("input, textarea, select, [contenteditable]");
    if (ev.key === "Escape" && S.open) { if (typing && ev.target.value) return; closeSheet(); return; }
    if (ev.key === "Enter" && ev.target.matches && ev.target.matches("tr[data-id]")) { openSheet(ev.target.dataset.id); return; }
    if (typing || ev.altKey || ev.ctrlKey || ev.metaKey) return;
    if ((ev.key === "ArrowDown" || ev.key === "ArrowUp") && S.open && S.open.indexOf("__") !== 0) { ev.preventDefault(); stepSheet(ev.key === "ArrowDown" ? 1 : -1); return; }
    if ((ev.key === "ArrowDown" || ev.key === "ArrowUp") && !S.open && ev.target.matches && ev.target.matches("tr[data-id]")) {
      ev.preventDefault(); var rs = Array.prototype.slice.call(document.querySelectorAll("#rows tr[data-id]")), ri = rs.indexOf(ev.target), nx = rs[ri + (ev.key === "ArrowDown" ? 1 : -1)]; if (nx) nx.focus();
    }
  });
  document.addEventListener("input", function (ev) {
    if (["eTo", "eCc", "eSubj", "eBody"].indexOf(ev.target.id) >= 0) syncOutlook();
    if (ev.target.id === "fText") { S.filters.text = ev.target.value; renderGrid(); $("btnClear").hidden = !filtersActive(); var b0 = semCourses().filter(function (c) { return S.scope !== "mine" || mine(c); }).length, v0 = visible().length; $("countLine").textContent = v0 === b0 ? b0 + " course files" : "Showing " + v0 + " of " + b0; }
  });
  document.addEventListener("change", function (ev) {
    var id = ev.target.id;
    if (id === "fProg") S.filters.prog = ev.target.value;
    else if (id === "semSel") { switchSem(ev.target.value); return; }
    else if (id === "setYear" || id === "setSem") { switchSem($("setYear").value + "/" + $("setSem").value); return; }
    else if (id === "sigFile") { if (ev.target.files[0]) saveSignature(ev.target.files[0]); ev.target.value = ""; return; }
    else if (id === "expTarget") { S.exp.target = ev.target.value; renderExport(); return; }
    else if (ev.target.dataset && ev.target.dataset.expchk) { S.exp[ev.target.dataset.expchk] = ev.target.checked; renderExport(); return; }
    else if (id === "fTerm") S.filters.term = ev.target.value;
    else if (id === "fState") S.filters.state = ev.target.value;
    else if (id === "fInst") S.filters.inst = ev.target.value;
    else if (ev.target.dataset && ev.target.dataset.remf) { S.remF[ev.target.dataset.remf] = ev.target.checked; renderReminders(); return; }
    else if (id === "impFile") { if (ev.target.files[0]) { $("settingsPanel").hidden = true; importSchedule(ev.target.files[0]); ev.target.value = ""; } return; }
    else if (id === "cPreset") { var pp = PRESETS.filter(function (p) { return p.id === ev.target.value; })[0]; S._preset = pp || null; if (pp) { $("cText").value = pp.text; if (pp.sec) $("cSec").value = String(pp.sec); } return; }
    else if (id === "mTerm") { var c = course(S.open); if (c) upd(c.id, { term: ev.target.value }).then(function () { toast("Term saved"); }); return; }
    else return;
    renderAll();
  });
  $("scrim").addEventListener("click", closeSheet);

  function createCourse() {
    var prog = $("nProg").value, cname = $("nCourse").value.trim(), sec = $("nSection").value.trim();
    if (!cname) { toast("Enter the course folder name"); return; }
    var rel = prog + "/" + cname + (sec ? "/" + sec : ""), id = rid(rel), ex = course(id);
    var doc = { program: prog, programShort: CF.programShort(prog), course: cname, section: sec, path: basePath() + "/" + rel, semester: semKey(), manual: true, gone: false, removed: false,
      instructor: $("nInst").value.trim(), email: $("nEmail").value.trim(), term: $("nTerm").value, enrolled: +$("nStud").value || 0,
      scan: { items: CF.SECTIONS.map(function (s) { return { k: s.k, s: "miss", note: "Nothing uploaded yet." }; }), issues: [], assessments: [], files: 0, state: "not_started", score: 0 }, review: { status: "pending" } };
    var p = ex ? S.db.collection("courses").doc(id).update({ removed: false, gone: false, instructor: doc.instructor || ex.instructor, email: doc.email || ex.email, term: doc.term || ex.term, enrolled: doc.enrolled || ex.enrolled }) : S.db.collection("courses").doc(id).set(doc);
    p.then(function () { toast(ex ? "Course file restored" : "Course file added"); closeSheet(); if (course(id)) openSheet(id); else S.pendingOpen = id; }).catch(function () { toast("Could not add the course file."); });
  }
  function fillSettings() {
    var c = S.config, d = semDue();
    $("dueNote").textContent = "Due dates for " + semLabel() + " (the semester you are viewing).";
    var sp = spFor(); $("spTitle").textContent = "SharePoint folder for " + semLabel(); $("setSp").value = sp ? sp.link || "" : ""; S._spChecked = null;
    $("spNote").innerHTML = sp ? '<span class="ok">✓ Set</span> · scans read <span class="mono">' + esc(sp.path.slice(LIB.length + 1)) + "</span>" + (sp.programs ? " (" + sp.programs.length + " program folders)" : "") + (sp.at ? " · checked " + fmtDT(sp.at) : "")
      : semKey() === FALL26 ? 'Not set: scans use the original folder <span class="mono">' + esc(defaultBase().slice(LIB.length + 1)) + "</span>. Paste a link to change it."
      : '<span class="bad">Not set.</span> Scanning is blocked for ' + esc(semLabel()) + " until you paste this semester's link.";
    // The semester picker in Settings switches the semester being set up (the same as the menu at the top).
    var ys = []; semOptions().forEach(function (o) { var y = o.k.split("/")[0]; if (ys.indexOf(y) < 0) ys.push(y); });
    $("setYear").innerHTML = ys.map(function (y) { return '<option value="' + esc(y) + '"' + (y === curYear() ? " selected" : "") + ">" + esc(shortYear(y)) + "</option>"; }).join("");
    $("setSem").value = curSem();
    var isDef = (c.year || "2026 - 2027") === curYear() && (c.semester || "Fall") === curSem();
    $("setDefault").checked = isDef; $("setDefault").disabled = isDef;
    $("setDefault").parentNode.title = isDef ? "This is already the default. To change it, pick another semester and tick this box there." : "";
    $("semNote").textContent = (isShort() ? semLabel() + " is a short six-week semester: one due date for every course file, no Term A or Term B." : semLabel() + " has Term A, Term B and full-semester due dates.") + (isDef ? " Folio opens on this semester by default." : "");
    var rem = S.courses.filter(function (c) { return c.removed && (c.semester || "2026 - 2027/Fall") === semKey(); });
    $("removedList").innerHTML = rem.length ? rem.map(function (c) { return '<div style="display:flex;gap:8px;align-items:center;justify-content:space-between;border-bottom:1px solid var(--line);padding:4px 0"><span>' + esc(c.course + (c.section ? " · " + c.section : "")) + '</span><button class="btn small" type="button" data-restore="' + esc(c.id) + '">Restore</button></div>'; }).join("") : '<span class="note">None.</span>';
    var sh = isShort(); $("lblDueA").hidden = sh; $("lblDueB").hidden = sh; $("lblDueFText").textContent = sh ? semLabel().split(" ")[0] + " due date (6-week semester)" : "Full-semester due date";
    $("dueA").value = d.A || ""; $("dueB").value = d.B || ""; $("dueF").value = d.F || ""; $("setLink").value = c.link || ""; $("sigPrev").innerHTML = c.signature ? '<img alt="Saved signature" style="max-height:48px;max-width:220px" src="' + esc(c.signature) + '">' : '<span class="note" style="color:#5B6763">No signature saved</span>'; $("sigRemove").hidden = !c.signature; $("setRevName").value = c.reviewerName || "Mousa Alhatti"; $("setRevEmail").value = c.reviewerEmail || "mousa.alhatti@adu.ac.ae";
  }
  // Signature: trim the white space and make the background transparent, then save it with the settings.
  function saveSignature(file) {
    var fr = new FileReader();
    fr.onload = function () {
      var img = new Image();
      img.onload = function () {
        var sc = Math.min(1, 900 / Math.max(img.width, img.height)), w = Math.round(img.width * sc), h = Math.round(img.height * sc);
        var cv = document.createElement("canvas"); cv.width = w; cv.height = h; var g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, w, h); g.drawImage(img, 0, 0, w, h);
        var d = g.getImageData(0, 0, w, h), p = d.data, x0 = w, y0 = h, x1 = -1, y1 = -1;
        for (var y = 0; y < h; y++) for (var x = 0; x < w; x++) { var i = (y * w + x) * 4, lum = 0.3 * p[i] + 0.59 * p[i + 1] + 0.11 * p[i + 2], a = Math.max(0, Math.min(255, Math.round((235 - lum) * 255 / 180))); p[i] = 24; p[i + 1] = 32; p[i + 2] = 30; p[i + 3] = a; if (a > 60) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } }
        if (x1 < 0) { toast("No signature found in that image."); return; }
        g.putImageData(d, 0, 0); x0 = Math.max(0, x0 - 8); y0 = Math.max(0, y0 - 8); x1 = Math.min(w - 1, x1 + 8); y1 = Math.min(h - 1, y1 + 8);
        var out = document.createElement("canvas"); out.width = x1 - x0 + 1; out.height = y1 - y0 + 1; out.getContext("2d").drawImage(cv, x0, y0, out.width, out.height, 0, 0, out.width, out.height);
        var url = out.toDataURL("image/png");
        S.db.collection("config").doc("settings").set(Object.assign({}, S.config, { signature: url })).then(function () { toast("Signature saved"); setTimeout(fillSettings, 300); }).catch(function () { toast("Could not save the signature."); });
      };
      img.onerror = function () { toast("That file could not be read as an image."); };
      img.src = fr.result;
    };
    fr.readAsDataURL(file);
  }
  async function saveSettings() {
    var raw = $("setSp").value.trim(), cur = spFor(), spNext;
    if (raw && (!cur || raw !== cur.link)) {
      var chk = S._spChecked && S._spChecked.link === raw ? S._spChecked : await runSpCheck();
      if (!chk || chk.err) { toast("Settings not saved: fix the SharePoint link first."); return; }
      spNext = { link: raw, path: chk.path, programs: chk.programs, at: nowISO(), by: S.me };
    }
    if (!raw && cur) { if (!S._spClear) { S._spClear = 1; $("spNote").innerHTML = '<span class="bad">The link box is empty.</span> Press Save again to remove this semester\'s SharePoint folder.'; return; } }
    S._spClear = 0;
    var spBy = Object.assign({}, S.config.spBySem || {}); if (spNext) spBy[semKey()] = spNext; else if (!raw) delete spBy[semKey()];
    var dueBy = Object.assign({}, S.config.dueBySem || {}); dueBy[semKey()] = isShort() ? { F: $("dueF").value } : { A: $("dueA").value, B: $("dueB").value, F: $("dueF").value };
    var next = Object.assign({}, S.config, { year: $("setDefault").checked ? curYear() : (S.config.year || "2026 - 2027"), semester: $("setDefault").checked ? curSem() : (S.config.semester || "Fall"), dueBySem: dueBy, spBySem: spBy, link: $("setLink").value.trim(), reviewerName: $("setRevName").value.trim(), reviewerEmail: $("setRevEmail").value.trim() });
    if (semKey() === FALL26) next.due = dueBy[FALL26];
    var defChanged = next.year !== S.config.year || next.semester !== S.config.semester;
    S.db.collection("config").doc("settings").set(next).then(function () { toast("Settings saved"); fx("click"); $("settingsPanel").hidden = true; if (defChanged) toast("Settings saved. Folio now opens on " + semLabel() + " by default."); }).catch(function () { toast("Could not save settings."); });
  }

  async function runSpCheck() {
    var n = $("spNote"), b = $("spCheck"); n.textContent = "Checking the folder on SharePoint…"; b.disabled = true;
    var r = await checkSpLink($("setSp").value); b.disabled = false;
    if (r.err) { n.innerHTML = '<span class="bad">✕ ' + esc(r.err) + "</span>"; S._spChecked = null; return r; }
    n.innerHTML = '<span class="ok">✓ Folder found</span> · <span class="mono">' + esc(r.path.slice(LIB.length + 1)) + "</span> · " + r.programs.length + " program folders: " + esc(r.programs.map(CF.programShort).join(", ")) + ". Press Save settings to use it for " + esc(semLabel()) + ".";
    S._spChecked = r; return r;
  }

  /* ---------- shared: spreadsheet library (loaded only when needed) ---------- */
  function loadXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise(function (res, rej) {
      var s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = function () { res(window.XLSX); }; s.onerror = function () { rej(new Error("Could not load the spreadsheet library.")); };
      document.head.appendChild(s);
    });
  }
  function itemSym(c, k) { var sc = scan(c); if (sc.state === "not_started") return "–"; var it = (sc.items || []).filter(function (i) { return i.k === k; })[0]; return it ? { ok: "✓", part: "!", miss: "✕" }[it.s] : "✕"; }
  function sheetModal(title, eyebrow, bodyHtml) {
    S.open = "__panel"; $("sheet").hidden = false; $("scrim").hidden = false;
    $("sheet").innerHTML = '<div class="sheet-head"><div class="row"><div style="display:grid;gap:4px"><span class="eyebrow">' + esc(eyebrow) + '</span><h2 id="dTitle" style="font-size:1.3rem">' + esc(title) + '</h2></div><button class="x" type="button" aria-label="Close">×</button></div></div><div class="sheet-body">' + bodyHtml + "</div>";
  }

  /* ---------- 8. Report (Excel) ---------- */
  async function downloadReport() {
    var dl = S.downloads;
    if (!dl) { toast("Downloads are not available in this view."); return; }
    toast("Preparing the report…");
    var X; try { X = await loadXLSX(); } catch (e) { toast(e.message); return; }
    var cs = semCourses(), sem = semLabel(), today = new Date().toISOString().slice(0, 10);
    var order = ["PG Diploma in Teaching", "M.Ed. Educational Leadership", "M.Ed. Educational Technologies & AI", "PhD in Education"];
    var progs = {}; cs.forEach(function (c) { var k = c.programShort; var p = progs[k] = progs[k] || { n: 0, complete: 0, in_progress: 0, not_due: 0, overdue: 0, approved: 0, score: 0, issues: 0 }; p.n++; p[dstate(c)]++; p.score += scan(c).score; if ((c.review || {}).status === "approved") p.approved++; p.issues += openComments(c.id).length; });
    var sum = [["Folio course file report", ""], ["Department", "Department of Education"], ["Semester", sem], ["Generated", today], ["Last SharePoint check", semLastScan() ? fmtDT(semLastScan().at) : "—"], [],
      ["Program", "Course files", "Complete", "In progress", "Not due yet", "Overdue", "Approved", "Completion rate", "Average score", "Open comments"]];
    var tot = { n: 0, complete: 0, in_progress: 0, not_due: 0, overdue: 0, approved: 0, score: 0, issues: 0 };
    Object.keys(progs).sort(function (a, b) { return order.indexOf(a) - order.indexOf(b); }).forEach(function (k) {
      var p = progs[k]; Object.keys(tot).forEach(function (f) { tot[f] += p[f]; });
      sum.push([k, p.n, p.complete, p.in_progress, p.not_due, p.overdue, p.approved, p.n ? Math.round(p.complete / p.n * 100) + "%" : "–", p.n ? Math.round(p.score / p.n) + "%" : "–", p.issues]);
    });
    sum.push(["All programs", tot.n, tot.complete, tot.in_progress, tot.not_due, tot.overdue, tot.approved, tot.n ? Math.round(tot.complete / tot.n * 100) + "%" : "–", tot.n ? Math.round(tot.score / tot.n) + "%" : "–", tot.issues]);
    sum.push([], ["Rules applied", "Seven course-file folders as in the model course file; graded samples per assessment and level: 1 (<30 students), 2 (30–60), 3 (>60); syllabus must state section and campus."]);
    var det = [["Program", "Course", "Section / folder", "Instructor", "Email", "Term", "Due date", "Students", "Samples per level", "Status", "Score %", "1 Checklist", "2 Syllabus", "3 Materials", "4 Assessment", "5 Attendance", "6 CRC", "7 SET", "Reviewer decision", "Decision date", "Open comments", "Instructor emailed", "Last scanned"]];
    var iss = [["Program", "Course", "Instructor", "Item", "Detail"]];
    cs.slice().sort(function (a, b) { return order.indexOf(a.programShort) - order.indexOf(b.programShort) || (a.course + a.section).localeCompare(b.course + b.section); }).forEach(function (c) {
      var sc = scan(c), rv = (c.review || {});
      det.push([c.programShort, c.course, c.section || "", c.instructor || "", c.email || "", eTerm(c) || "", dueFor(c) || "", c.enrolled || "", need(c), stateLabel(dstate(c)), sc.score, itemSym(c, 1), itemSym(c, 2), itemSym(c, 3), itemSym(c, 4), itemSym(c, 5), itemSym(c, 6), itemSym(c, 7), ({ pending: "Not reviewed", approved: "Approved", revise: "Needs revision" })[rv.status || "pending"], rv.at ? rv.at.slice(0, 10) : "", openComments(c.id).length, c.emailedAt ? c.emailedAt.slice(0, 10) : "", c.scannedAt ? c.scannedAt.slice(0, 10) : ""]);
      if (sc.state !== "not_started") (sc.items || []).forEach(function (i) { if (i.s !== "ok") iss.push([c.programShort, c.course, c.instructor || "", i.k + ". " + CF.SECTIONS[i.k - 1].short, i.note]); });
      (sc.issues || []).forEach(function (x) { iss.push([c.programShort, c.course, c.instructor || "", "Automatic check", x]); });
      openComments(c.id).forEach(function (m) { iss.push([c.programShort, c.course, c.instructor || "", "Reviewer comment" + (m.sec ? " (" + CF.SECTIONS[m.sec - 1].short + ")" : ""), m.text]); });
    });
    var wb = X.utils.book_new();
    function add(name, aoa, widths) { var ws = X.utils.aoa_to_sheet(aoa); ws["!cols"] = widths.map(function (w) { return { wch: w }; }); X.utils.book_append_sheet(wb, ws, name); }
    add("Summary", sum, [34, 14, 10, 12, 12, 10, 10, 15, 14, 14]);
    add("Course files", det, [26, 34, 28, 22, 28, 13, 11, 9, 10, 12, 8, 9, 9, 9, 11, 10, 7, 7, 15, 12, 10, 12, 12]);
    add("Issues", iss.length > 1 ? iss : iss.concat([["", "", "", "", "No open issues"]]), [26, 34, 22, 26, 90]);
    var buf = X.write(wb, { type: "array", bookType: "xlsx" });
    try { await dl.save({ filename: "Folio report - Education - " + sem.replace(/[^\w\s-]/g, "-") + " - " + today + ".xlsx", data: new Blob([buf]) }); toast("Report saved"); }
    catch (e) { if (e && e.code !== "declined") toast("Could not save the report (" + (e.code || "error") + ")."); }
  }

/*@@EXPORT@@*/
  /* ---------- 9. Batch reminders ---------- */
  function reminderGroups(f) {
    var soon = Date.now() + 14 * 864e5, by = {};
    semCourses().forEach(function (c) {
      var st = dstate(c), d = dueFor(c), dueSoon = d && new Date(d) <= new Date(soon);
      var take = (f.overdue && st === "overdue") || (f.inprog && st === "in_progress") || (f.soon && st === "not_due" && dueSoon) || (f.overdue && st === "in_progress" && isOverdue(c));
      if (!take || !c.instructor) return;
      var k = (c.email || c.instructor).toLowerCase(); (by[k] = by[k] || { name: c.instructor, email: c.email, files: [] }).files.push(c);
    });
    return Object.keys(by).map(function (k) { return by[k]; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
  }
  function reminderBody(g) {
    var sem = semLabel(), L = [salute(g.name), "", "This is a friendly reminder about your course files for " + sem + ":", ""];
    g.files.forEach(function (c) {
      var sc = scan(c), nd = need(c);
      L.push("• " + c.course + (c.section && !/,/.test(c.section) ? " (" + c.section + ")" : "") + " — " + stateLabel(dstate(c)) + (dueFor(c) ? ", due " + fmtDate(dueFor(c)) : "") + (sc.state !== "not_started" ? ", " + sc.score + "% complete" : ""));
      if (sc.state === "not_started") L.push("   The folder is still empty.");
      else (sc.items || []).filter(function (i) { return i.s !== "ok"; }).forEach(function (i) { L.push("   - " + i.k + ". " + CF.SECTIONS[i.k - 1].short + ": " + i.note); });
      L.push((sampleShort(c) ? "   Graded samples needed: " + nd + " High, " + nd + " Medium, " + nd + " Low per assessment. " : "   ") + "Folder: " + folderUrl(c.path), "");
    });
    if (g.files.some(function (c) { return syllabusShort(c) && !(c.accepted && c.accepted[2]); })) L.push("Please also make sure each syllabus states the section and campus.");
    if (S.config.link) L.push("You can follow your course file status here: " + S.config.link);
    return L.join("\n");
  }
  function renderReminders() {
    S._panelIs = "rem";
    var f = S.remF || (S.remF = { overdue: true, inprog: true, soon: false });
    var gs = reminderGroups(f), sem = semLabel();
    var h = '<div class="meta" style="gap:14px">' +
      '<label><input type="checkbox" data-remf="overdue"' + (f.overdue ? " checked" : "") + "> Overdue</label>" +
      '<label><input type="checkbox" data-remf="inprog"' + (f.inprog ? " checked" : "") + "> In progress (incomplete)</label>" +
      '<label><input type="checkbox" data-remf="soon"' + (f.soon ? " checked" : "") + "> Not started, due within 14 days</label></div>" +
      '<p class="note" style="margin:0">' + gs.length + " instructor" + (gs.length === 1 ? "" : "s") + " to remind. Each email lists only that instructor’s course files and missing items. Open each one in Outlook, check it and press Send, then mark it as sent.</p>";
    if (!gs.length) h += '<p class="empty">Nobody matches these filters.</p>';
    gs.forEach(function (g, i) {
      var href = "https://outlook.office.com/mail/deeplink/compose?to=" + encodeURIComponent(g.email || "") + "&subject=" + encodeURIComponent("Reminder: course files " + sem) + "&body=" + encodeURIComponent(reminderBody(g));
      var last = g.files.map(function (c) { return c.emailedAt || ""; }).sort().pop();
      h += '<div class="cmt"><div class="h"><b>' + esc(g.name) + "</b><span>" + esc(g.email || "no email recorded") + "</span>" + (last ? "<span>last emailed " + fmtDT(last) + "</span>" : "") + '</div><div class="t">' + g.files.map(function (c) { return esc(c.course) + " — " + esc(stateLabel(dstate(c))); }).join("<br>") + '</div><div class="acts"><a class="btn small primary" target="_blank" rel="noopener" href="' + esc(href) + '">Open in Outlook ↗</a><button class="btn small" type="button" data-remcopy="' + i + '">Copy text</button><button class="btn small" type="button" data-remsent="' + i + '">Mark as sent</button></div></div>';
    });
    S._remGroups = gs;
    sheetModal("Email reminders", "Batch reminders · " + sem, h);
  }

  /* ---------- 6. Import the term schedule ---------- */
  var CAMP = { AD: "Abu Dhabi", AA: "Al Ain", DB: "Dubai" };
  function normCode(s) { return String(s || "").replace(/\(E\)/gi, "").replace(/[\s-]/g, "").toUpperCase(); }
  function rowCodes(c) {
    var name = c.course || "", campf = [];
    if (/^EDUL680\s*A\b/i.test(name)) return { codes: ["EDUL680A"], camp: (name.match(/\b(AD|AA|DB)\b/g) || []) };
    var codes = name.split(/\s+-\s+|\s*\+\s*/).map(normCode).filter(function (x) { if (CAMP[x]) { campf.push(x); return false; } return !!x; });
    return { codes: codes, camp: campf };
  }
  function nameKeys(c) {
    var k = [];
    if (c.section && /,/.test(c.section)) k.push(c.section.split(",")[0].trim().toLowerCase());
    var parts = norm(c.instructor || "").split(" ").filter(function (w) { return w.length > 2; });
    if (parts.length) { k.push(parts[parts.length - 1]); if (parts[0].length > 3) k.push(parts[0]); }
    return k;
  }
  function guessProgram(code) {
    var pf = programFolders();
    var pick = function (re) { return pf.filter(function (p) { return re.test(p); })[0] || ""; };
    if (/^EDUT/.test(code)) return pick(/technolog/i);
    if (/^EDUP/.test(code)) return pick(/philosophy/i);
    if (/^EDUL|^EDU5/.test(code)) return pick(/leadership/i);
    return pick(/diploma/i);
  }
  async function importSchedule(file) {
    var X; try { X = await loadXLSX(); } catch (e) { toast(e.message); return; }
    var wb = X.read(await file.arrayBuffer(), { cellDates: true });
    var raw = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
    var col = function (r, names) { for (var i = 0; i < names.length; i++) if (r[names[i]] !== undefined && r[names[i]] !== "") return r[names[i]]; return ""; };
    var lines = raw.map(function (r) {
      return { code: normCode(col(r, ["Descr", "Course", "Subject Code"])), section: String(col(r, ["Section"])), enr: +col(r, ["Tot Enrl", "Enrolled", "Total Enrolled"]) || 0, name: String(col(r, ["Name", "Instructor"])), campus: String(col(r, ["Campus"])).toUpperCase(), start: col(r, ["Start Date"]), end: col(r, ["End Date"]) };
    }).filter(function (l) { return /^ED/.test(l.code) && l.name && !/be assigned/i.test(l.name); });
    if (!lines.length) { toast("No Education courses were found in that file."); return; }
    var starts = lines.map(function (l) { return +new Date(l.start); }).filter(Boolean), semStart = Math.min.apply(null, starts);
    function termOf(ls) {
      var t = ls.map(function (l) { var s = +new Date(l.start), e = +new Date(l.end); if (!s || !e) return ""; if ((e - s) / 864e5 > 70) return "Full semester"; return s - semStart < 21 * 864e5 ? "Term A" : "Term B"; }).filter(Boolean);
      return isShort() ? SHORT : t[0] || "";
    }
    var used = {}, updates = [], notOn = [];
    semCourses().forEach(function (c) {
      var rc = rowCodes(c), keys = nameKeys(c);
      var m = lines.filter(function (l, i) {
        if (rc.codes.indexOf(l.code) < 0) return false;
        if (rc.camp.length && rc.camp.indexOf(l.campus) < 0) return false;
        var nl = l.name.toLowerCase();
        if (keys.length && !keys.some(function (k) { return nl.indexOf(k) >= 0; })) return false;
        return true;
      });
      m.forEach(function (l) { used[lines.indexOf(l)] = 1; });
      if (!m.length) { notOn.push(c); return; }
      var enr = m.reduce(function (s, l) { return s + l.enr; }, 0), camps = Array.from(new Set(m.map(function (l) { return CAMP[l.campus] || l.campus; }))).sort(), term = termOf(m);
      var ch = {};
      if ((+c.enrolled || 0) !== enr) ch.enrolled = enr;
      if (term && c.term !== term) ch.term = term;
      if (JSON.stringify(c.campuses || []) !== JSON.stringify(camps)) ch.campuses = camps;
      if (Object.keys(ch).length) updates.push({ c: c, ch: ch });
    });
    var missing = {}; lines.forEach(function (l, i) { if (used[i]) return; var k = l.code + "|" + l.name; (missing[k] = missing[k] || { code: l.code, name: l.name, enr: 0, camps: {} }); missing[k].enr += l.enr; missing[k].camps[CAMP[l.campus] || l.campus] = 1; });
    S._imp = { updates: updates, notOn: notOn, missing: Object.keys(missing).map(function (k) { return missing[k]; }), file: file.name, n: lines.length };
    renderImport();
  }
  function renderImport() {
    S._panelIs = "imp";
    var I = S._imp, h = '<p class="note" style="margin:0">' + esc(I.file) + ": " + I.n + " Education class lines read and compared with the " + esc(semLabel()) + " course files in Folio.</p>" +
      (hasFolder() ? '<p class="note" style="margin:0">SharePoint folder for ' + esc(semLabel()) + ': <span class="mono">' + esc(basePath().slice(LIB.length + 1)) + "</span></p>" : '<div class="banner warn"><b>No SharePoint folder is set for ' + esc(semLabel()) + '.</b> <span>Add it in Settings before scanning, so the course files are read from this semester\'s folder.</span></div>');
    h += '<section><h3>Updates to apply (' + I.updates.length + ")</h3>";
    if (!I.updates.length) h += '<p class="note">Class sizes, terms and campuses already match the schedule.</p>';
    else {
      h += '<div style="overflow-x:auto;background:var(--surface);border:1px solid var(--line);border-radius:12px"><table class="amx"><thead><tr><th></th><th>Course file</th><th>Students</th><th>Term</th><th>Campus</th></tr></thead><tbody>' + I.updates.map(function (u, i) {
        var c = u.c, ch = u.ch, f = function (k, cur) { return ch[k] !== undefined ? '<span class="muted">' + esc(cur || "—") + "</span> → <b>" + esc(Array.isArray(ch[k]) ? ch[k].join(", ") : ch[k]) + "</b>" : esc(Array.isArray(cur) ? (cur || []).join(", ") : (cur || "")); };
        return '<tr><td><input type="checkbox" data-impu="' + i + '" checked aria-label="Apply"></td><td>' + esc(c.course) + "<br><span class=\"note\">" + esc(c.instructor || c.section || "") + "</span></td><td>" + f("enrolled", c.enrolled) + "</td><td>" + f("term", c.term) + "</td><td>" + f("campuses", c.campuses) + "</td></tr>";
      }).join("") + '</tbody></table></div><div class="actions"><button class="btn primary" type="button" id="impApply">Apply selected updates</button></div>';
    }
    h += "</section>";
    h += '<section><h3>On the schedule but no course file in Folio (' + I.missing.length + ")</h3>";
    h += I.missing.length ? I.missing.map(function (m, i) { return '<div class="cmt"><div class="h"><b>' + esc(m.code) + "</b><span>" + esc(m.name) + "</span><span>" + esc(Object.keys(m.camps).join(", ")) + " · " + m.enr + " student" + (m.enr === 1 ? "" : "s") + '</span></div><div class="acts"><button class="btn small" type="button" data-impadd="' + i + '">Add course file…</button></div></div>'; }).join("") : '<p class="note">None.</p>';
    h += '<p class="note" style="margin:0">A course may also be missing because its SharePoint folder uses a different code or instructor name. Check the folder before adding.</p></section>';
    h += '<section><h3>In Folio but not on the schedule (' + I.notOn.length + ")</h3>";
    h += I.notOn.length ? I.notOn.map(function (c) { return '<div class="cmt"><div class="h"><b>' + esc(c.course) + "</b><span>" + esc(c.instructor || c.section || "") + '</span></div><div class="acts"><button class="btn small" type="button" data-impopen="' + esc(c.id) + '">Open</button></div></div>'; }).join("") : '<p class="note">None.</p>';
    h += "</section>";
    sheetModal("Schedule check", "Import schedule", h);
  }
  async function applyImport() {
    var I = S._imp, sel = Array.prototype.filter.call(document.querySelectorAll("[data-impu]"), function (el) { return el.checked; }).map(function (el) { return I.updates[+el.dataset.impu]; });
    for (var i = 0; i < sel.length; i++) await S.db.collection("courses").doc(sel[i].c.id).update(sel[i].ch);
    toast(sel.length + " course files updated from the schedule");
    I.updates = I.updates.filter(function (u) { return sel.indexOf(u) < 0; }); renderImport();
  }
  function prefillAdd(m) {
    S.open = "__new"; renderAddSheet();
    var prog = guessProgram(m.code); if (prog) $("nProg").value = prog;
    var folder = semCourses().filter(function (c) { return rowCodes(c).codes.indexOf(m.code) >= 0; })[0];
    $("nCourse").value = folder ? folder.course : m.code;
    $("nSection").value = /,/.test(m.name) ? m.name : "";
    var parts = m.name.split(","); $("nInst").value = parts.length > 1 ? (parts[1].trim().split(/\s+/)[0] + " " + parts[0].trim()) : m.name;
    var p = S.people.filter(function (q) { return norm(q.name) === norm($("nInst").value); })[0]; if (p) $("nEmail").value = p.email;
    $("nStud").value = m.enr || "";
  }

  /* ---------- boot ---------- */
  async function resolveNames() {
    if (!S.user) return;
    var ids = Array.from(new Set(S.comments.map(function (m) { return m.by; }).filter(Boolean))).filter(function (id) { return !S.names[id]; });
    if (!ids.length) return;
    try { var ps = await S.user.profiles(ids); Object.assign(S.names, ps); if (S.open) renderSheet(); } catch (e) {}
  }
  async function boot() {
    renderAll();
    var claude = window.claude;
    if (!claude || !claude.use) { renderGrid(); return; }
    var dbP = claude.use("db"), userP = claude.use("user"), mcpP = claude.use("mcp");
    S.user = await userP;
    if (S.user) {
      try { S.isLead = !!(await S.user.canEdit()); } catch (e) {}
      try { S.me = await S.user.id(); } catch (e) {}
      try { var me = await S.user.me(); S.myEmail = ((me && me.email) || "").toLowerCase(); S.myName = norm((me && me.name) || ""); S.myDisplayName = (me && me.name) || ""; } catch (e) {}
      if (!S.myName) { try { S.myName = norm(await S.user.name()); } catch (e) {} }
    }
    S.db = await dbP;
    if (!S.db) { renderAll(); return; }
    S.db.collection("courses").onSnapshot(function (snap) {
      S.courses = snap.docs.map(function (d) { return Object.assign({ id: d.id }, d.data()); });
      if (!S.isLead && S.courses.some(mine) && !S._scoped) { S.scope = "mine"; S._scoped = true; }
      if (S.pendingOpen && course(S.pendingOpen)) { var po = S.pendingOpen; S.pendingOpen = null; setTimeout(function () { openSheet(po); }, 0); }
      renderAll();
    }, function (e) { banner("err", "Could not load course files (" + esc(e.code) + ")."); });
    S.db.collection("comments").onSnapshot(function (snap) { S.comments = snap.docs.map(function (d) { return Object.assign({ id: d.id }, d.data()); }); renderAll(); resolveNames(); });
    S.db.collection("people").onSnapshot(function (snap) { S.people = snap.docs.map(function (d) { return Object.assign({ id: d.id }, d.data()); }); });
    S.db.collection("config").doc("settings").onSnapshot(function (d) { S.config = (d.exists && d.data()) || {}; CF.setYear(curYear()); renderAll(); });
    if (S.isLead) { S.mcp = await mcpP; try { S.downloads = await claude.use("downloads"); } catch (e) {} }
    renderAll();
  }
  setSound(soundOn(), true);
  boot();
})();
