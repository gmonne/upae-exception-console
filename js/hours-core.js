// =====================================================================
// WEEKLY HOURS: upload the B2W "Employee Hours Weekly Summary", save it, check it against the rules
// =====================================================================
var hrsWeek = null;        // { fileName, weekNo, start, end, days:[iso...], records:[{id,name,days:{iso:{rt,ot,dt}}}], dupIds:[] }
var hrsSaved = null;       // rows already in SharePoint (null = not loaded / could not be read)
var hrsSavedError = null;
var hrsFindings = null;    // result of the last rule check
var hrsFilter = 'all';
var hrsPage = 1, hrsPageSize = 50, hrsListHidden = false;   // paging and show/hide for the findings list
var hrsBusy = false;
var DOW = ['sun','mon','tue','wed','thu','fri','sat'];
var DOW_LABEL = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
var HRS_EXCESSIVE = 16;    // a single day above this is flagged for a look
var HRS_TOL = 0.01;

function hrsRound(n){ return Math.round(n * 100) / 100; }
function hrsIso(d){ return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
function hrsFromIso(s){ var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function hrsPretty(iso){ var d = hrsFromIso(iso); return DOW_LABEL[d.getDay()] + ' ' + (d.getMonth() + 1) + '/' + d.getDate(); }
function hrsSerialToDate(n){ var d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
function hrsIsoWeek(d){
  var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var dn = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - dn);
  var y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - y0) / 86400000 + 1) / 7);
}
function hrsParseCell(v){
  var out = { rt: 0, ot: 0, dt: 0 };
  String(v == null ? '' : v).split(/\r?\n/).forEach(function(line){
    var m = line.trim().match(/^(-?\d+(?:\.\d+)?)\s*(RT|OT|DT)$/i);
    if(m) out[m[2].toLowerCase()] += parseFloat(m[1]);
  });
  return out;
}
function hrsDateFromHeader(c){
  var s = empClean(c);
  if(!s) return null;
  if(/^\d{5}(\.\d+)?$/.test(s)) return hrsSerialToDate(parseFloat(s));
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m) return new Date(+m[3], +m[1] - 1, +m[2]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function hrsParseSummary(rows, fileName){
  var hi = -1, idCol = -1, nameCol = 0;
  for(var r = 0; r < Math.min(rows.length, 30); r++){
    var cells = (rows[r] || []).map(function(c){ return empClean(c).toLowerCase(); });
    var k = cells.indexOf('employee id');
    if(k !== -1){ hi = r; idCol = k; break; }
  }
  if(hi === -1) throw new Error('Could not find the header row. The weekly summary needs an "Employee ID" column followed by one column per day.');
  var header = rows[hi], dayCols = [];
  for(var c = idCol + 1; c < header.length; c++){
    var d = hrsDateFromHeader(header[c]);
    if(d) dayCols.push({ col: c, iso: hrsIso(d), date: d });
  }
  if(!dayCols.length) throw new Error('No date columns were found after "Employee ID".');
  var filters = '';
  for(var f = 0; f < hi; f++) (rows[f] || []).forEach(function(c){ filters += ' ' + String(c || ''); });
  var rng = filters.match(/\[(\d{1,2}\/\d{1,2}\/\d{4})\s*-\s*(\d{1,2}\/\d{1,2}\/\d{4})\]/);
  var start = dayCols[0].date, end = dayCols[dayCols.length - 1].date;
  if(rng){ start = hrsDateFromHeader(rng[1]) || start; end = hrsDateFromHeader(rng[2]) || end; }
  var records = [], seen = {}, dupIds = [];
  for(var j = hi + 1; j < rows.length; j++){
    var row = rows[j] || [];
    var id = empClean(row[idCol]);
    var nm = empClean(row[nameCol]);
    if(!id) continue;                       // blank rows and the Total row have no ID
    var key = empKey(id);
    var rec = { id: id, name: nm, days: {} };
    dayCols.forEach(function(dc){
      var h = hrsParseCell(row[dc.col]);
      if(h.rt || h.ot || h.dt) rec.days[dc.iso] = h;
    });
    if(seen[key]){ dupIds.push(id); continue; }
    seen[key] = true;
    records.push(rec);
  }
  // the Monday inside the range decides the ISO week number (Sunday-start weeks belong to the Monday after)
  var mon = new Date(start); while(mon.getDay() !== 1 && mon <= end) mon.setDate(mon.getDate() + 1);
  return { fileName: fileName, start: hrsIso(start), end: hrsIso(end), weekNo: hrsIsoWeek(mon > end ? start : mon),
           days: dayCols.map(function(d){ return d.iso; }), records: records, dupIds: dupIds, hadRange: !!rng };
}

// The B2W "Employee Hours Detail" report: one line per job per employee per day.
// Columns are found by their header names, so a moved or added column does not break the import.
function hrsIsDetail(rows){
  for(var r = 0; r < Math.min(rows.length, 30); r++){
    var cells = (rows[r] || []).map(function(c){ return empClean(c).toLowerCase(); });
    if(cells.indexOf('employee id') !== -1 && cells.indexOf('job number') !== -1 && cells.indexOf('work start') !== -1) return true;
  }
  return false;
}
function hrsDateTimeFromCell(c){
  var s = empClean(c);
  if(!s) return null;
  if(/^\d{5}(\.\d+)?$/.test(s)){
    var n = parseFloat(s), frac = n - Math.floor(n), mins = Math.round(frac * 1440);
    return { date: hrsSerialToDate(Math.floor(n)), time: ('0' + Math.floor(mins / 60) % 24).slice(-2) + ':' + ('0' + mins % 60).slice(-2) };
  }
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if(m) return { date: new Date(+m[3], +m[1] - 1, +m[2]), time: m[4] ? ('0' + m[4]).slice(-2) + ':' + m[5] : '' };
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2}))?/);
  if(m) return { date: new Date(+m[1], +m[2] - 1, +m[3]), time: m[4] ? m[4] + ':' + m[5] : '' };
  return null;
}
function hrsParseDetail(rows, fileName){
  var hi = -1, col = {};
  for(var r = 0; r < Math.min(rows.length, 30); r++){
    var cells = (rows[r] || []).map(function(c){ return empClean(c).toLowerCase(); });
    if(cells.indexOf('employee id') !== -1 && cells.indexOf('job number') !== -1){ hi = r; cells.forEach(function(c, i){ if(c && !(c in col)) col[c] = i; }); break; }
  }
  if(hi === -1) throw new Error('Could not find the header row of the Employee Hours Detail report.');
  ['employee id', 'work start', 'rt hours', 'ot hours', 'dt hours'].forEach(function(k){
    if(!(k in col)) throw new Error('The detail report needs a "' + k.replace(/\b\w/g, function(c){ return c.toUpperCase(); }) + '" column.');
  });
  var filters = '';
  for(var f = 0; f < hi; f++) (rows[f] || []).forEach(function(c){ filters += ' ' + String(c || ''); });
  var sm = filters.match(/Start Date\s*-\s*(\d{1,2}\/\d{1,2}\/\d{4})/i), em = filters.match(/End Date\s*-\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  var recs = {}, order = [], lines = 0, minD = null, maxD = null, bad = 0, jobTitles = {};
  function num(v){ var x = parseFloat(empClean(v)); return isNaN(x) ? 0 : x; }
  for(var j = hi + 1; j < rows.length; j++){
    var row = rows[j] || [];
    var id = empClean(row[col['employee id']]);
    if(!id) continue;                                  // blank rows, "Employee: X" subtotals and the grand total have no ID
    var dt = hrsDateTimeFromCell(row[col['work start']]);
    if(!dt){ bad++; continue; }
    var iso = hrsIso(dt.date);
    if(!minD || iso < minD) minD = iso; if(!maxD || iso > maxD) maxD = iso;
    var key = empKey(id), rec = recs[key];
    if(!rec){ rec = recs[key] = { id: id, name: empClean(row[col['name']]), days: {}, jobs: {} }; order.push(key); }
    var rt = num(row[col['rt hours']]), ot = num(row[col['ot hours']]), d2 = num(row[col['dt hours']]);
    var h = rec.days[iso] || (rec.days[iso] = { rt: 0, ot: 0, dt: 0 });
    h.rt += rt; h.ot += ot; h.dt += d2;
    var job = empClean(row[col['job number']]), trk = empClean(row[col['tracking id']]), lab = empClean(row[col['labor type']]);
    if(job && !jobTitles[hrsJobBase(job)]) jobTitles[hrsJobBase(job)] = empClean(row[col['title']]);
    var list = rec.jobs[iso] || (rec.jobs[iso] = []), hit = null;
    list.forEach(function(x){ if(x.job === job && x.tracking === trk && x.laborType === lab && x.start === dt.time) hit = x; });
    if(!hit){ hit = { job: job, title: empClean(row[col['title']]), tracking: trk, laborType: lab, start: dt.time, rt: 0, ot: 0, dt: 0 }; list.push(hit); }
    hit.rt += rt; hit.ot += ot; hit.dt += d2; lines++;
  }
  if(!order.length) throw new Error('No hour lines were found in the detail report.');
  var start = sm ? hrsDateFromHeader(sm[1]) : hrsFromIso(minD), end = em ? hrsDateFromHeader(em[1]) : hrsFromIso(maxD);
  var days = [], d = new Date(start);
  for(var g = 0; d <= end && g < 40; g++){ days.push(hrsIso(d)); d.setDate(d.getDate() + 1); }
  var mon = new Date(start); while(mon.getDay() !== 1 && mon <= end) mon.setDate(mon.getDate() + 1);
  var records = order.map(function(k){
    var rec = recs[k];
    Object.keys(rec.days).forEach(function(i){ var x = rec.days[i]; x.rt = hrsRound(x.rt); x.ot = hrsRound(x.ot); x.dt = hrsRound(x.dt); if(!(x.rt || x.ot || x.dt)) delete rec.days[i]; });
    return rec;
  });
  return { fileName: fileName, source: 'detail', start: hrsIso(start), end: hrsIso(end), weekNo: hrsIsoWeek(mon > end ? start : mon),
           days: days, records: records, dupIds: [], hadRange: !!(sm && em), jobTitles: jobTitles, lines: lines, skippedLines: bad };
}
// 23-430~0328 and 25-615-0103 belong to jobs 23-430 and 25-615
function hrsJobBase(j){ j = empClean(j); var m = j.match(/^(\d{2}-\d{3})(?:[~-].*)?$/); return m ? m[1] : j; }
// "24-110 (8 RT + 1 OT); 23-604~0099 (13.5 OT)" for one employee-day
function hrsJobText(rec, iso){
  var list = rec && rec.jobs && rec.jobs[iso]; if(!list || !list.length) return '';
  var by = {}, order = [];
  list.forEach(function(x){ var o = by[x.job] || (by[x.job] = { job: x.job, rt: 0, ot: 0, dt: 0 }); if(order.indexOf(x.job) < 0) order.push(x.job); o.rt += x.rt; o.ot += x.ot; o.dt += x.dt; });
  return order.map(function(j){ var o = by[j]; return (o.job || '(no job)') + ' (' + hrsFmt(o) + ')'; }).join('; ');
}
function hrsJobCount(rec, iso){
  var list = rec && rec.jobs && rec.jobs[iso]; if(!list) return 0;
  var seen = {}; list.forEach(function(x){ seen[x.job] = 1; }); return Object.keys(seen).length;
}

// ---------------- turning a rule into something that can be calculated ----------------
function hrsDayIndex(tok){ return DOW.indexOf(String(tok || '').slice(0, 3).toLowerCase()); }
function hrsExpandDays(a, b){
  var i = hrsDayIndex(a); if(i < 0) return null;
  if(b == null) return [i];
  var j = hrsDayIndex(b); if(j < 0) return null;
  var out = [], k = i, guard = 0;
  while(guard++ < 8){ out.push(k); if(k === j) break; k = (k + 1) % 7; }
  return out;
}
function hrsSplitSegments(text){
  var segs = [], depth = 0, cur = '';
  String(text).split('').forEach(function(ch){
    if(ch === '(') depth++; if(ch === ')') depth = Math.max(0, depth - 1);
    if((ch === ',' || ch === ';') && depth === 0){ segs.push(cur); cur = ''; } else cur += ch;
  });
  if(cur.trim()) segs.push(cur);
  return segs;
}
// Reads one overtime tier. Returns { na:true } | { ok:true, daily:{dow:startHrs}, weekly:N|null } | { ok:false, why }
function hrsReadTier(label, trigger, threshold, daysText, stdDay){
  var trig = empClean(trigger).toLowerCase();
  if(!trig || /not applicable|^n\/?a$/.test(trig)) return { na: true };
  var out = { ok: true, daily: {}, weekly: null };
  var isWeekly = /week/.test(trig) && !/day-of-week/.test(trig) && !/daily/.test(trig);
  var thrNum = (String(threshold).match(/(\d+(?:\.\d+)?)/) || [])[1];
  if(isWeekly){
    if(!thrNum) return { ok: false, why: label + ' is a weekly rule but no hours threshold could be read from "' + threshold + '".' };
    out.weekly = parseFloat(thrNum);
    return out;
  }
  var segs = hrsSplitSegments(daysText).filter(function(s){ return s.trim(); });
  if(!segs.length) return { ok: false, why: label + ' has no applicable days filled in.' };
  var overMatch = String(threshold).match(/over\s+(\d+(?:\.\d+)?)/i);
  var defaultStart = overMatch ? parseFloat(overMatch[1]) : (stdDay || null);
  for(var i = 0; i < segs.length; i++){
    var m = segs[i].trim().match(/^([A-Za-z]{3,9})(?:\s*(?:[–—-]|to)\s*([A-Za-z]{3,9}))?\s*(?:\(([^)]*)\))?\s*$/);
    if(!m) return { ok: false, why: label + ' applicable days "' + empClean(daysText) + '" could not be read (stuck on "' + segs[i].trim() + '").' };
    var days = hrsExpandDays(m[1], m[2]);
    if(!days) return { ok: false, why: label + ' applicable days "' + empClean(daysText) + '" has an unknown day name.' };
    var note = (m[3] || '').toLowerCase(), start;
    if(/\ball\b/.test(note)) start = 0;
    else if(/over\s*(\d+(?:\.\d+)?)/.test(note)) start = parseFloat(note.match(/over\s*(\d+(?:\.\d+)?)/)[1]);
    else if(defaultStart != null) start = defaultStart;
    else return { ok: false, why: label + ' says nothing about when the tier starts (add "(over 8)" or "(all)" after the days, or "Over 8 hrs/day" in the threshold).' };
    days.forEach(function(d){ out.daily[d] = start; });
  }
  return out;
}
function hrsCompileRule(rule){
  if(!rule) return null;
  var F = rule.fields || {};
  if(rule.status !== 'structured') return { rule: rule, usable: false, why: 'Rule "' + rule.name + '" is not marked Structured yet (' + (rule.statusLabel || 'draft') + ').' };
  var std = parseFloat(F['Standard workday (hrs)']) || null;
  var t15 = hrsReadTier('OT 1.5x', F['OT 1.5 trigger type'], F['OT 1.5 threshold'], F['OT 1.5 applicable days'], std);
  var t2  = hrsReadTier('OT 2x',  F['OT 2x trigger type'],  F['OT 2x threshold'],  F['OT 2x applicable days'],  std);
  if(t15.ok === false) return { rule: rule, usable: false, why: 'Rule "' + rule.name + '": ' + t15.why };
  if(t2.ok === false)  return { rule: rule, usable: false, why: 'Rule "' + rule.name + '": ' + t2.why };
  if(t15.na && t2.na) return { rule: rule, usable: false, why: 'Rule "' + rule.name + '" has neither an OT 1.5x nor an OT 2x tier filled in.' };
  return { rule: rule, usable: true, t15: t15, t2: t2 };
}

// Splits a day's total into RT / OT(1.5x) / DT(2x) the way the rule says it should be.
function hrsExpectedWeek(comp, recDays, isoDays){
  var res = {}, cum = 0;
  isoDays.slice().sort().forEach(function(iso){
    var h = recDays[iso]; if(!h) return;
    var total = hrsRound(h.rt + h.ot + h.dt), dow = hrsFromIso(iso).getDay();
    var s15 = (!comp.t15.na && comp.t15.daily && comp.t15.daily[dow] != null) ? comp.t15.daily[dow] : Infinity;
    var s2  = (!comp.t2.na  && comp.t2.daily  && comp.t2.daily[dow]  != null) ? comp.t2.daily[dow]  : Infinity;
    var dt = Math.max(0, total - s2);
    var ot = s15 < s2 ? Math.max(0, Math.min(total, s2) - s15) : 0;
    var rt = total - dt - ot;
    // weekly thresholds: straight-time hours past N in the week move up a tier
    var wk15 = (!comp.t15.na && comp.t15.weekly != null) ? comp.t15.weekly : null;
    var wk2  = (!comp.t2.na  && comp.t2.weekly  != null) ? comp.t2.weekly  : null;
    var before = cum; cum += total;
    if(wk15 != null && rt > 0){ var over = Math.max(0, Math.min(rt, cum - Math.max(before, wk15))); rt -= over; ot += over; }
    if(wk2 != null){ var over2 = Math.max(0, Math.min(rt + ot, cum - Math.max(before, wk2))); var fromRt = Math.min(rt, over2); rt -= fromRt; ot -= (over2 - fromRt); dt += over2; }
    res[iso] = { rt: hrsRound(rt), ot: hrsRound(ot), dt: hrsRound(dt), total: total };
  });
  return res;
}

function hrsFindRule(unionName){
  var n = empClean(unionName).toLowerCase();
  if(!n) return null;
  return rules.filter(function(r){ return empClean(r.name).toLowerCase() === n; })[0] || null;
}

