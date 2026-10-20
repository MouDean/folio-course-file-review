/* ===== Course-file checker: shared logic (template = USS001 model course file) ===== */
var CF = (function () {
  var SECTIONS = [
    { k: 1, label: "Course checklist", short: "Checklist", kw: /check\s*list/i },
    { k: 2, label: "Course syllabus (current + previous)", short: "Syllabus", kw: /syllab/i },
    { k: 3, label: "All instructional materials", short: "Materials", kw: /material|teaching|instruction|presentation/i },
    { k: 4, label: "Assessment (instruments, keys/rubrics, High/Medium/Low samples)", short: "Assessment", kw: /assess/i },
    { k: 5, label: "Student attendance data", short: "Attendance", kw: /attend/i },
    { k: 6, label: "Comprehensive review (CRC)", short: "CRC", kw: /comprehensive|review|crc/i },
    { k: 7, label: "Summary of student feedback (SET)", short: "SET", kw: /\bset\b|evaluation|feedback/i }
  ];
  var TERM_RE = /(20\d\d)\s*[-_/ ]\s*(20)?(\d\d)/g;
  var YEAR_TAG = "2026-27";

  function allFiles(n) {
    var out = (n.files || []).slice();
    (n.folders || []).forEach(function (f) { out = out.concat(allFiles(f)); });
    return out;
  }
  function secNum(name) {
    var m = /^\s*([1-7])(?![0-9])\s*[.\-_)\s]*/.exec(name || "");
    return m ? +m[1] : 0;
  }
  function isRoot(n) {
    var nums = {}, kws = 0;
    (n.folders || []).forEach(function (f) {
      var k = secNum(f.name);
      if (k) nums[k] = 1;
      else if (SECTIONS.some(function (s) { return s.kw.test(f.name); })) kws++;
    });
    return Object.keys(nums).length + kws >= 2;
  }
  // The 1-7 folders may sit inside one or more wrapper folders; look a few levels down.
  function findRoot(n, depth) {
    if (isRoot(n)) return n;
    if (!depth) return null;
    var fs = n.folders || [];
    for (var i = 0; i < fs.length; i++) { var r = findRoot(fs[i], depth - 1); if (r) return r; }
    return null;
  }
  function findSection(root, s) {
    var fs = root.folders || [];
    var hit = fs.filter(function (f) { return secNum(f.name) === s.k; });
    if (!hit.length) hit = fs.filter(function (f) { return !secNum(f.name) && s.kw.test(f.name); });
    return hit[0] || null;
  }
  function ext(name) { var m = /\.([a-z0-9]+)$/i.exec(name || ""); return m ? m[1].toLowerCase() : ""; }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }

  var ROLE_INS = /instrument|brief|question|paper|task|guideline/i;
  var ROLE_KEY = /answer|marking|scheme|rubric|key|solution|worked/i;
  var ROLE_GRD = /graded|sample|student\s*work|scripts/i;
  var LV = [["h", /high/i], ["m", /medium|\bmed\b|mid/i], ["l", /\blow/i]];

  function levelOf(name) {
    for (var i = 0; i < LV.length; i++) if (LV[i][1].test(name)) return LV[i][0];
    return null;
  }
  function roleOf(name) {
    if (ROLE_GRD.test(name)) return "grd";
    if (levelOf(name)) return "lvl";
    if (/answer|marking|scheme|rubric|solution|worked/i.test(name)) return "key";
    if (ROLE_INS.test(name)) return "ins";
    return null;
  }
  function levels(node, acc) {
    (node.folders || []).forEach(function (f) {
      var l = levelOf(f.name);
      if (l) acc[l] += allFiles(f).length;
      else levels(f, acc);
    });
    return acc;
  }
  function checkAssessment(a) {
    var r = { name: a.name, ins: 0, key: 0, h: 0, m: 0, l: 0, loose: 0 };
    (a.files || []).forEach(function (f) {
      if (ROLE_KEY.test(f.name)) r.key++; else r.ins++;
      if (/rubric|guideline/i.test(f.name) && ROLE_KEY.test(f.name)) { /* instrument + rubric in one file */ r.ins = Math.max(r.ins, 1); }
    });
    var graded = null;
    (a.folders || []).forEach(function (f) {
      var role = roleOf(f.name);
      if (role === "ins") { r.ins += allFiles(f).length; allFiles(f).forEach(function (x) { if (/key|answer|rubric|marking/i.test(x.name)) r.key++; }); }
      else if (role === "key") r.key += allFiles(f).length;
      else if (role === "grd") graded = f;
    });
    var lv = levels(graded || a, { h: 0, m: 0, l: 0 });
    r.h = lv.h; r.m = lv.m; r.l = lv.l;
    if (graded) r.loose = (graded.files || []).length;
    return r;
  }

  function evaluate(root) {
    var items = [], issues = [], assessments = [];
    var total = allFiles(root).length;
    SECTIONS.forEach(function (s) {
      var f = findSection(root, s);
      var it = { k: s.k, s: "miss", note: "" };
      if (!f) { it.note = "Folder “" + s.k + ". " + s.short + "” not found."; items.push(it); return; }
      var files = allFiles(f);
      var n = files.length;
      if (s.k === 2) {
        var cur = null, prev = null;
        (f.folders || []).forEach(function (x) {
          if (/current|this|present/i.test(x.name)) cur = x;
          else if (/previous|prior|past|last|old/i.test(x.name)) prev = x;
        });
        var nc = cur ? allFiles(cur).length : 0, np = prev ? allFiles(prev).length : 0;
        if (cur || prev) {
          if (nc && np) {
            it.s = "ok";
            var prevOnlyNote = allFiles(prev).every(function (x) { return x.size === 0 || ext(x.name) === "txt"; });
            it.note = prevOnlyNote ? "Current syllabus found; previous folder holds a note (first delivery)." : "Current and previous syllabi found.";
          } else if (nc) { it.s = "part"; it.note = "Previous offering syllabus missing (add it, or a note if this is the first delivery)."; }
          else if (np) { it.s = "part"; it.note = "Current syllabus missing."; }
          else it.note = "Syllabus folders are empty.";
          var curFiles = cur ? allFiles(cur) : [];
          curFiles.forEach(function (x) {
            var m, bad = null; TERM_RE.lastIndex = 0;
            while ((m = TERM_RE.exec(x.name))) { var yy = m[1] + "-" + m[3]; if (yy !== YEAR_TAG) bad = yy; }
            if (bad) issues.push("Syllabus: the current syllabus file “" + x.name + "” mentions " + bad + ". Check it is the version for this semester.");
          });
        } else if (n >= 2) { it.s = "part"; it.note = plural(n, "file") + " found. Put them in “Current” and “Previous” sub-folders."; }
        else if (n === 1) { it.s = "part"; it.note = "Only one syllabus found. Add the syllabus from the most recent previous offering, in a “Previous” sub-folder."; }
        else it.note = "No syllabus uploaded.";
      } else if (s.k === 4) {
        var subs = f.folders || [];
        var setLike = subs.some(function (x) { var r = roleOf(x.name); return r === "grd" || r === "ins" || r === "key"; });
        var units = setLike ? [{ name: "All assessments", files: f.files, folders: subs }] : subs;
        if (!setLike && (f.files || []).length && !subs.length) units = [{ name: "Assessment", files: f.files, folders: [] }];
        units.forEach(function (u) { assessments.push(checkAssessment(u)); });
        if (!n) it.note = "No assessment evidence uploaded.";
        else if (!assessments.length) { it.s = "part"; it.note = "Files found, but not organised by assessment."; }
        else {
          var gaps = [];
          assessments.forEach(function (a) {
            var miss = [];
            if (!a.ins) miss.push("assessment instrument");
            if (!a.key) miss.push("answer key / marking scheme or rubric");
            var lv = []; if (!a.h) lv.push("High"); if (!a.m) lv.push("Medium"); if (!a.l) lv.push("Low");
            if (lv.length) miss.push("graded sample: " + lv.join(", "));
            if (miss.length) gaps.push(a.name + ": missing " + miss.join("; ") + ".");
            if (a.loose) issues.push(a.name + ": " + plural(a.loose, "graded file") + " not placed in a High / Medium / Low folder.");
          });
          it.s = gaps.length ? "part" : "ok";
          it.note = gaps.length ? gaps.length + " of " + assessments.length + " assessments incomplete." : plural(assessments.length, "assessment") + " with instrument, key/rubric and High, Medium and Low samples.";
          gaps.forEach(function (g) { issues.push("Assessment – " + g); });
        }
      } else if (s.k === 6) {
        var xl = files.filter(function (x) { return /^xls/.test(ext(x.name)); });
        if (xl.length) { it.s = "ok"; it.note = "CRC workbook: " + xl.map(function (x) { return x.name; }).join(", "); }
        else if (n) { it.s = "part"; it.note = "File found, but not the CRC Excel workbook."; }
        else it.note = "No comprehensive review uploaded.";
      } else if (s.k === 7) {
        if (n) { it.s = "ok"; it.note = files.some(function (x) { return /disclaim|not available|n\/?a/i.test(x.name); }) ? "SET placeholder/disclaimer uploaded (replace with results once released)." : "SET summary uploaded."; }
        else it.note = "No SET summary or disclaimer uploaded.";
      } else if (s.k === 1) {
        if (n) { it.s = "ok"; it.note = files.map(function (x) { return x.name; }).slice(0, 2).join(", "); }
        else it.note = "Checklist not uploaded.";
      } else {
        if (n) { it.s = "ok"; it.note = files.some(function (x) { return x.name === "(contents)"; }) ? "Uploaded." : plural(n, "file") + " uploaded."; }
        else it.note = "Folder is empty.";
      }
      items.push(it);
    });
    var ok = items.filter(function (i) { return i.s === "ok"; }).length;
    var part = items.filter(function (i) { return i.s === "part"; }).length;
    var state = total === 0 ? "not_started" : (ok === 7 ? "complete" : "in_progress");
    return { items: items, issues: issues, assessments: assessments.map(function (a) { return { name: a.name, ins: a.ins > 0, key: a.key > 0, h: a.h, m: a.m, l: a.l }; }), files: total, state: state, score: Math.round(((ok + part * 0.5) / 7) * 100) };
  }

  function emptyEval() {
    return { items: SECTIONS.map(function (s) { return { k: s.k, s: "miss", note: "Nothing uploaded yet." }; }), issues: [], assessments: [], files: 0, state: "not_started", score: 0 };
  }

  function programShort(name) {
    if (/philosophy|phd/i.test(name)) return "PhD in Education";
    if (/technolog/i.test(name)) return "M.Ed. Educational Technologies & AI";
    if (/leadership/i.test(name)) return "M.Ed. Educational Leadership";
    if (/diploma/i.test(name)) return "PG Diploma in Teaching";
    return name;
  }
  function hash(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); }
  function rowId(rel) { return (rel.replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60) + "-" + hash(rel)); }
  function prettyInstructor(folder) {
    if (!folder) return "";
    var m = /^([^,]+),\s*(.+)$/.exec(folder.replace(/\s*\(.*?\)\s*/g, " ").trim());
    if (m) return (m[2].split(/\s+/)[0] + " " + m[1]).trim();
    return folder.replace(/^(ms|mr|dr|mrs)\.?\s+/i, "").trim();
  }

  /* Turn the department tree into course-file rows. Tree: Department > Program > Course > [section/instructor] > 1..7 */
  function rows(dept, basePath) {
    var out = [];
    (dept.folders || []).forEach(function (prog) {
      (prog.folders || []).forEach(function (course) {
        var roots, kids = course.folders || [];
        if (isRoot(course)) roots = [{ node: course, sub: "" }];
        // A single wrapper folder (e.g. "EDUL520 Fall- Term A Ghanem/EDUL520 Term A Fall 2026") is still the course's own file.
        // (Only when the 1-7 folders sit deeper than that single folder; an instructor folder holding them directly keeps its own row.)
        else if (kids.length === 1 && !isRoot(kids[0]) && findRoot(kids[0], 3)) roots = [{ node: findRoot(kids[0], 3), sub: "" }];
        else if (kids.length) roots = kids.map(function (f) { return { node: findRoot(f, 3) || f, sub: f.name }; });
        else roots = [{ node: course, sub: "" }];
        roots.forEach(function (r) {
          var rel = prog.name + "/" + course.name + (r.sub ? "/" + r.sub : "");
          var ev = isRoot(r.node) ? evaluate(r.node) : (allFiles(r.node).length ? Object.assign(emptyEval(), { state: "in_progress", files: allFiles(r.node).length, issues: ["Files uploaded, but not in the 1–7 course-file folders."] }) : emptyEval());
          out.push({
            id: rowId(rel),
            program: prog.name,
            programShort: programShort(prog.name),
            course: course.name,
            section: r.sub,
            instructorFromFolder: prettyInstructor(/\d/.test(r.sub) && !/,/.test(r.sub) ? "" : r.sub),
            path: basePath + "/" + rel,
            node: r.node,
            scan: ev
          });
        });
      });
    });
    return out;
  }

  function setYear(y) { var m = /(20\d\d)\D+20(\d\d)/.exec(y || ""); if (m) YEAR_TAG = m[1] + "-" + m[2]; }
  return { rowId: rowId, setYear: setYear, SECTIONS: SECTIONS, evaluate: evaluate, rows: rows, isRoot: isRoot, secNum: secNum, programShort: programShort };
})();
if (typeof module !== "undefined") module.exports = CF;
