// ---------------- the check ----------------
var hrsFindSaved = null, hrsFindError = null;   // WeeklyFindings rows already in SharePoint
var covPage = 1, covPageSize = 10;   // paging for the Rule coverage table
var hrsJobRows = [], jobPage = 1, jobPageSize = 50, jobEdit = null;   // every JobAssignments row, for the Project managers page
var hrsJobPMs = {}, hrsJobsError = null;           // job base number -> { spItemId, jobName, pmEmail, pmName }
var hrsPMs = [], hrsPMError = null;             // Project managers (name + email)
var hrsSel = {};                                // selected finding keys
var hrsStatusFilter = 'all';
var hrsLoadedOnce = false;
var hrsPmEdit = null;
var hrsAi = {};   // finding key -> { state:'loading'|'done'|'error', data, err, logId, feedback }
var STATUS_OPEN = 'Open', STATUS_SENT = 'Sent to PM', STATUS_DONE = 'Resolved';

function parTxt(x){ x = empClean(x); return x ? ' (' + x + ')' : ''; }
// Which rule was broken, in words, for one employee-day.
function hrsBroken(comp, rep, exp, iso){
  var F = comp.rule.fields || {}, out = [], dow = DOW_LABEL[hrsFromIso(iso).getDay()];
  var t15 = parTxt(F['OT 1.5 threshold'] || F['OT 1.5 applicable days']);
  var t2 = parTxt(F['OT 2x threshold'] || F['OT 2x applicable days']);
  if(rep.ot > exp.ot + HRS_TOL){
    out.push(comp.t15.na ? 'This union has no 1.5x overtime tier, but ' + rep.ot + ' hrs were reported as OT (1.5x). Overtime for ' + comp.rule.name + ' is paid at 2x' + t2 + '.'
                         : 'More 1.5x overtime was reported than the OT 1.5x tier allows on ' + dow + t15 + '.');
  }
  if(rep.ot < exp.ot - HRS_TOL) out.push('Hours that belong in the OT 1.5x tier on ' + dow + t15 + ' were reported at a different rate.');
  if(rep.dt < exp.dt - HRS_TOL) out.push('Hours that belong in the OT 2x tier on ' + dow + t2 + ' were reported at a lower rate.');
  if(rep.dt > exp.dt + HRS_TOL) out.push('Double time (2x) was reported on ' + dow + ', but the OT 2x tier does not reach that many hours there' + parTxt(F['OT 2x applicable days']) + '.');
  if(rep.rt > exp.rt + HRS_TOL) out.push('More regular (straight) time was reported than the rule allows' + parTxt(F['Regular threshold']) + '.');
  if(rep.rt < exp.rt - HRS_TOL) out.push('Fewer straight-time hours were reported than the rule gives' + parTxt(F['Regular threshold']) + '.');
  return out;
}
function hrsRunCheck(){
  var wk = hrsWeek; if(!wk) return null;
  var empByKey = {}; employees.forEach(function(e){ empByKey[empKey(e.id)] = e; });
  var compCache = {}, curRec = null;
  var out = { items: [], emp: { checked: 0, regular: 0, notListed: [], noLabor: [], noUnion: [], noRule: {}, hours: 0 }, rowsChecked: 0, matched: 0, byUnion: {} };
  function add(base, extra){
    var it = Object.assign({}, base, extra);
    if(it.iso && curRec){ it.jobs = hrsJobText(curRec, it.iso); it.jobCount = hrsJobCount(curRec, it.iso); }
    it.key = wk.start + '|' + empKey(it.id) + '|' + (it.iso || '-') + '|' + it.kind;
    it.status = STATUS_OPEN; it.sentTo = ''; it.sentDate = ''; it.note = ''; it.resReason = ''; it.resolvedBy = ''; it.resolvedDate = '';
    out.items.push(it);
  }
  wk.records.forEach(function(rec){
    curRec = rec;
    var total = 0; Object.keys(rec.days).forEach(function(i){ var h = rec.days[i]; total += h.rt + h.ot + h.dt; });
    out.emp.hours += total;
    var e = empByKey[empKey(rec.id)];
    var base = { id: rec.id, name: rec.name, union: '', ruleId: 'Data check', ruleName: 'Data check', cba: '' };
    Object.keys(rec.days).forEach(function(iso){
      var h = rec.days[iso], t = hrsRound(h.rt + h.ot + h.dt);
      if(t > HRS_EXCESSIVE) add(base, { kind: 'excessive', sev: 'warn', iso: iso, rep: h, msg: t + ' hours logged in one day (over ' + HRS_EXCESSIVE + ')' + (hrsJobCount(rec, iso) > 1 ? ' across ' + hrsJobCount(rec, iso) + ' jobs: possible second shift, verify.' : '.'), broken: ['Data check: a single day above ' + HRS_EXCESSIVE + ' hours is flagged for confirmation (not a union rule).'] });
      if(h.rt < 0 || h.ot < 0 || h.dt < 0) add(base, { kind: 'negative', sev: 'crit', iso: iso, rep: h, msg: 'Negative hours in the file.', broken: ['Data check: hours cannot be negative.'] });
    });
    if(!total) return;
    if(!e){ out.emp.notListed.push(rec); add(base, { kind: 'notlisted', sev: 'warn', iso: '', msg: 'Employee ID is not in the Employees list, so no union or rule could be applied.', broken: ['Data check: every employee in the hours report must exist on the Employees list.'] }); return; }
    var info = employeeUnionInfo(e);
    if(info.kind === 'regular'){ out.emp.regular++; return; }
    if(info.kind === 'nolabor'){ out.emp.noLabor.push(rec); add(base, { kind: 'nolabor', sev: 'warn', iso: '', msg: 'No default labor type on the Employees list, so no union rule could be applied.', broken: ['Data check: employees need a default labor type (or the Employee type) so their union can be found.'] }); return; }
    if(info.kind !== 'union'){ out.emp.noUnion.push(rec); return; }
    var unionName = empUnionName(e);
    if(!(unionName in compCache)) compCache[unionName] = hrsCompileRule(hrsFindRule(unionName)) || { usable: false, why: 'No rule exists yet for ' + unionName + '.' };
    var comp = compCache[unionName];
    if(!comp.usable){
      (out.emp.noRule[unionName] = out.emp.noRule[unionName] || { why: comp.why, n: 0 }).n++;
      return;
    }
    out.emp.checked++;
    var bu = out.byUnion[unionName] = out.byUnion[unionName] || { days: 0, bad: 0, emps: {} };
    var exp = hrsExpectedWeek(comp, rec.days, wk.days);
    Object.keys(exp).forEach(function(iso){
      var rep = rec.days[iso], x = exp[iso];
      out.rowsChecked++; bu.days++;
      var dRt = hrsRound(rep.rt - x.rt), dOt = hrsRound(rep.ot - x.ot), dDt = hrsRound(rep.dt - x.dt);
      if(Math.abs(dRt) > HRS_TOL || Math.abs(dOt) > HRS_TOL || Math.abs(dDt) > HRS_TOL){
        bu.bad++; bu.emps[empKey(rec.id)] = true;
        add(base, { kind: 'mismatch', sev: 'crit', iso: iso, union: unionName, rep: rep, exp: x, ruleId: comp.rule.id, ruleName: comp.rule.name, cba: comp.rule.cba || '',
          broken: hrsBroken(comp, rep, x, iso),
          msg: 'Reported ' + hrsFmt(rep) + '; the rule gives ' + hrsFmt(x) + '.' });
      } else out.matched++;
    });
  });
  out.hours = hrsRound(out.emp.hours);
  out.ranAt = new Date();
  hrsMergeSaved(out);
  return out;
}
function hrsFmt(h){
  var p = [];
  if(h.rt) p.push(hrsRound(h.rt) + ' RT'); if(h.ot) p.push(hrsRound(h.ot) + ' OT'); if(h.dt) p.push(hrsRound(h.dt) + ' DT');
  return p.join(' + ') || '0';
}
// carry review status over from what was saved earlier for the same finding
function hrsMergeSaved(f){
  var by = {}; (hrsFindSaved || []).forEach(function(r){ by[r.key] = r; });
  var seen = {};
  f.items.forEach(function(i){
    var sv = by[i.key]; seen[i.key] = true;
    if(sv){ if(!i.jobs && sv.jobs) i.jobs = sv.jobs; i.status = sv.status || STATUS_OPEN; i.sentTo = sv.sentTo || ''; i.sentDate = sv.sentDate || ''; i.note = sv.note || ''; i.spItemId = sv.spItemId; i.resReason = sv.resReason || ''; i.resolvedBy = sv.resolvedBy || ''; i.resolvedDate = sv.resolvedDate || '';
      i.pmResponse = sv.pmResponse || ''; i.pmNote = sv.pmNote || ''; i.pmBy = sv.pmBy || ''; i.pmDate = sv.pmDate || '';
      if(sv.aiExplanation){
        i.ai = { explanation: sv.aiExplanation, likelyCause: sv.aiLikelyCause, suggestedAction: sv.aiSuggestedAction, evidence: sv.aiEvidence ? sv.aiEvidence.split(' | ') : [],
                 confidence: sv.aiConfidence || 'low', model: sv.aiModel, date: sv.aiDate, feedback: sv.aiFeedback || '' };
        if(!hrsAi[i.key] || hrsAi[i.key].state !== 'loading')
          hrsAi[i.key] = Object.assign({}, hrsAi[i.key], { state: 'done', saved: true, collapsed: (hrsAi[i.key] && hrsAi[i.key].collapsed != null) ? hrsAi[i.key].collapsed : true, feedback: sv.aiFeedback || (hrsAi[i.key] && hrsAi[i.key].feedback) || '',
            data: { stub: false, model: sv.aiModel, confidence: i.ai.confidence, explanation: i.ai.explanation, likelyCause: i.ai.likelyCause, suggestedAction: i.ai.suggestedAction, evidence: i.ai.evidence, needsMoreInfo: false } });
      }
    }
  });
  f.stale = (hrsFindSaved || []).filter(function(r){ return hrsWeek && r.weekStart === hrsWeek.start && !seen[r.key]; }).length;
}

// ---------------- saved rows (SharePoint) ----------------
function hrsKey(id, iso){ return empKey(id) + '|' + iso; }
function hrsBuildRows(wk){
  var rows = [];
  wk.records.forEach(function(rec){
    Object.keys(rec.days).forEach(function(iso){
      var h = rec.days[iso];
      rows.push({ key: hrsKey(rec.id, iso), employeeId: rec.id, employeeName: rec.name, workDate: iso, weekStart: wk.start, weekNo: wk.weekNo,
                  rt: hrsRound(h.rt), ot: hrsRound(h.ot), dt: hrsRound(h.dt) });
    });
  });
  return rows;
}
function hrsPlanSave(){
  var wk = hrsWeek, rows = hrsBuildRows(wk), cur = {};
  (hrsSaved || []).forEach(function(r){ cur[r.key] = r; });
  var fresh = [], changed = [], same = 0, seen = {};
  rows.forEach(function(r){
    seen[r.key] = true;
    var c = cur[r.key];
    if(!c) fresh.push(r);
    else if(Math.abs(c.rt - r.rt) > HRS_TOL || Math.abs(c.ot - r.ot) > HRS_TOL || Math.abs(c.dt - r.dt) > HRS_TOL || String(c.weekNo) !== String(r.weekNo)) changed.push({ cur: c, next: r });
    else same++;
  });
  var gone = (hrsSaved || []).filter(function(r){ return r.weekStart === wk.start && !seen[r.key]; });
  return { fresh: fresh, changed: changed, same: same, gone: gone, total: rows.length };
}
function hrsFindingRow(i){
  return { key: i.key, weekStart: hrsWeek.start, weekNo: hrsWeek.weekNo, id: i.id, name: i.name, union: i.union, iso: i.iso || '', kind: i.kind,
           ruleId: i.ruleId, ruleName: i.ruleName, broken: (i.broken || []).join(' | '), reported: i.rep ? hrsFmt(i.rep) : '', expected: i.exp ? hrsFmt(i.exp) : '',
           jobs: i.jobs || '', status: i.status, sentTo: i.sentTo, sentDate: i.sentDate, note: i.note, resReason: i.resReason || '', resolvedBy: i.resolvedBy || '', resolvedDate: i.resolvedDate || '',
           aiExplanation: i.ai ? i.ai.explanation : '', aiLikelyCause: i.ai ? i.ai.likelyCause : '', aiSuggestedAction: i.ai ? i.ai.suggestedAction : '',
           aiEvidence: i.ai ? (i.ai.evidence || []).join(' | ') : '', aiConfidence: i.ai ? i.ai.confidence : '', aiModel: i.ai ? i.ai.model : '', aiDate: i.ai ? i.ai.date : '', aiFeedback: i.ai ? (i.ai.feedback || '') : '' };
}
// create or update the WeeklyFindings rows for these items
async function hrsPersistFindings(items, onProgress, opts){
  opts = opts || {};
  if(!items.length) return { ok: 0, bad: [] };
  if(hrsFindError) return { ok: 0, bad: [{ error: 'WeeklyFindings list could not be read: ' + hrsFindError }] };
  if(typeof upaeBatchFindings !== 'function') return { ok: 0, bad: [{ error: 'SharePoint is not connected in this view.' }] };
  var savedOf = function(i){ return i.spItemId ? (hrsFindSaved || []).filter(function(x){ return x.spItemId === i.spItemId; })[0] : null; };
  var me = (typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : '');
  // 1. has somebody else changed these rows since this screen loaded them?
  if(!opts.force && typeof upaeFetchFindingStamps === 'function'){
    try{
      var ids = items.filter(function(i){ return i.spItemId; }).map(function(i){ return i.spItemId; });
      var stamps = ids.length ? await upaeFetchFindingStamps(ids) : {}, clash = [];
      items.forEach(function(i){
        var cur = savedOf(i), st = i.spItemId ? stamps[i.spItemId] : null;
        if(!cur || !st || !st.modified || !cur.modified) return;
        if(st.modified > cur.modified && String(st.modifiedBy).toLowerCase() !== String(me).toLowerCase()) clash.push({ item: i, by: st.modifiedBy || 'someone else', at: st.modified });
      });
      if(clash.length){
        var lines = clash.slice(0, 3).map(function(c){ return '• ' + c.item.name + ' (' + (c.item.iso ? hrsPretty(c.item.iso) : 'record') + ') – changed by ' + c.by + ' on ' + new Date(c.at).toLocaleString(); }).join('\n');
        var go = confirm(clash.length + ' record(s) were changed by someone else since you opened this week:\n\n' + lines + (clash.length > 3 ? '\n…' : '') + '\n\nOK = overwrite their change with yours.\nCancel = keep their change and reload the week.');
        if(!go){
          await hrsRefreshSaved();
          return { ok: 0, bad: [{ error: 'Cancelled – someone else changed these records, so the latest version was reloaded.' }], cancelled: true };
        }
      }
    } catch(err){ console.warn('UPAE: could not check for concurrent edits', err); }
  }
  var before = items.map(function(i){ var c = savedOf(i); return c ? { status: c.status, sentTo: c.sentTo } : null; });
  var ops = items.map(function(i){ var r = hrsFindingRow(i); return i.spItemId ? { kind: 'update', row: r, spItemId: i.spItemId } : { kind: 'create', row: r }; });
  var res = await upaeBatchFindings(ops, onProgress), ok = 0, bad = [], hist = [], today = new Date().toISOString();
  res.forEach(function(r, k){
    if(r && r.ok){
      ok++;
      var it = items[k], row = ops[k].row;
      if(!it.spItemId){ it.spItemId = r.spItemId; hrsFindSaved = hrsFindSaved || []; hrsFindSaved.push(Object.assign({}, row, { spItemId: r.spItemId })); }
      else { var cur = (hrsFindSaved || []).filter(function(x){ return x.spItemId === it.spItemId; })[0]; if(cur) Object.assign(cur, row); }
      var prev = before[k] ? before[k].status : 'Open';
      if(row.status !== prev){
        var detail = row.status === STATUS_DONE ? [row.resReason, row.note].filter(Boolean).join(' – ') : row.status === STATUS_SENT ? ('Sent to ' + row.sentTo) : '';
        hist.push({ key: row.key, weekStart: row.weekStart, id: row.id, name: row.name, from: prev, to: row.status, detail: detail, by: me });
      }
    } else bad.push({ error: r && r.error });
  });
  // 2. remember the new modified stamps so our own saves are not mistaken for somebody else's
  if(ok && typeof upaeFetchFindingStamps === 'function'){
    try{
      var ids2 = items.filter(function(i){ return i.spItemId; }).map(function(i){ return i.spItemId; }), st2 = await upaeFetchFindingStamps(ids2);
      items.forEach(function(i){ var c = savedOf(i), s = st2[i.spItemId]; if(c && s && s.modified){ c.modified = s.modified; c.modifiedBy = s.modifiedBy; } });
    } catch(err){ console.warn('UPAE: could not refresh modified stamps', err); }
  }
  // 3. change history (never blocks the save)
  var histErr = '';
  if(hist.length && typeof upaeLogHistory === 'function'){
    try{ var hr = await upaeLogHistory(hist); var hb = hr.filter(function(x){ return !x.ok; }); if(hb.length) histErr = hb[0].error || 'unknown error'; }
    catch(err){ histErr = err.message || String(err); }
    if(histErr) console.warn('UPAE: FindingHistory not written:', histErr);
  }
  if(typeof window.upaeDashInvalidate === 'function') window.upaeDashInvalidate();
  return { ok: ok, bad: bad, histErr: histErr };
}

// ---------------- saved weeks ----------------
function hrsSavedWeeks(){
  var by = {};
  (hrsSaved || []).forEach(function(r){
    var w = by[r.weekStart] = by[r.weekStart] || { start: r.weekStart, weekNo: r.weekNo, emps: {}, hours: 0, rows: 0 };
    w.emps[empKey(r.employeeId)] = true; w.hours += r.rt + r.ot + r.dt; w.rows++;
  });
  return Object.keys(by).sort().reverse().map(function(k){ by[k].n = Object.keys(by[k].emps).length; return by[k]; });
}
function hrsAddDays(iso, n){ var d = hrsFromIso(iso); d.setDate(d.getDate() + n); return hrsIso(d); }
function hrsOpenSaved(start){
  var rows = (hrsSaved || []).filter(function(r){ return r.weekStart === start; });
  if(!rows.length) return;
  var by = {}, order = [];
  rows.forEach(function(r){
    var k = empKey(r.employeeId);
    if(!by[k]){ by[k] = { id: r.employeeId, name: r.employeeName, days: {} }; order.push(k); }
    by[k].days[r.workDate] = { rt: r.rt, ot: r.ot, dt: r.dt };
  });
  var days = []; for(var i = 0; i < 7; i++) days.push(hrsAddDays(start, i));
  hrsWeek = { fileName: 'Saved in SharePoint', start: start, end: days[6], weekNo: rows[0].weekNo, days: days,
              records: order.sort().map(function(k){ return by[k]; }), dupIds: [], hadRange: false, fromSaved: true };
  hrsSel = {}; hrsFilter = 'all'; hrsStatusFilter = 'all'; hrsPage = 1;
  document.getElementById('hrsFileName').textContent = 'Week ' + hrsWeek.weekNo + ' (saved)';
  hrsFindings = hrsRunCheck();
  renderHoursPanel();
}
async function hrsRefreshSaved(){
  var st = document.getElementById('hrsSavedStatus'); if(st) st.textContent = 'Reading saved weeks from SharePoint…';
  hrsSaved = null; hrsSavedError = null; hrsFindSaved = null; hrsFindError = null; hrsPMError = null;
  if(typeof upaeFetchWeeklyHours !== 'function'){ hrsSavedError = hrsFindError = hrsPMError = 'not connected to SharePoint'; }
  else {
    try{ hrsSaved = await upaeFetchWeeklyHours(); } catch(err){ console.warn('UPAE: WeeklyHours', err); hrsSavedError = err.message || String(err); }
    try{ hrsFindSaved = await upaeFetchFindings(); } catch(err){ console.warn('UPAE: WeeklyFindings', err); hrsFindError = err.message || String(err); }
    try{ hrsPMs = await upaeFetchProjectManagers(); } catch(err){ console.warn('UPAE: ProjectManagers', err); hrsPMError = err.message || String(err); }
    hrsJobsError = null; hrsJobPMs = {}; hrsJobRows = [];
    try{
      hrsJobRows = typeof upaeFetchJobAssignments === 'function' ? await upaeFetchJobAssignments() : [];
      hrsJobRows.forEach(function(a){ if(a.active && a.job && a.pmEmail) hrsJobPMs[hrsJobBase(a.job)] = a; });
    } catch(err){ console.warn('UPAE: JobAssignments', err); hrsJobsError = err.message || String(err); }
    hrsSums = null; hrsSumsError = null;
    try{ hrsSums = typeof upaeFetchSummaries === 'function' ? await upaeFetchSummaries() : []; } catch(err){ console.warn('UPAE: WeeklySummaries', err); hrsSumsError = err.message || String(err); hrsSums = []; }
  }
  hrsLoadedOnce = true;
  if(hrsFindings && hrsWeek) hrsMergeSaved(hrsFindings);
  renderHoursSaved(); renderHoursPanel(); renderPMs(); renderJobs();
}

// ---------------- screen ----------------
var HRS_KINDS = {
  mismatch: 'Hours differ from rule', notlisted: 'Not in Employees list', nolabor: 'No default labor type',
  excessive: 'Very long day', negative: 'Negative hours'
};
function fmtMD(iso){ var d = hrsFromIso(iso); return (d.getMonth() + 1) + '/' + d.getDate(); }
function renderHoursKpis(){
  var el = document.getElementById('hoursKpis'); if(!el) return;
  var wk = hrsWeek, f = hrsFindings;
  function kpi(label, val, sub, cls){ return '<div class="kpi"><div class="k-label">' + label + '</div><div class="k-value">' + val + '</div><div class="k-delta ' + (cls || 'flat') + '">' + sub + '</div></div>'; }
  if(!wk){
    el.innerHTML = kpi('Week', '—', 'no week loaded yet') + kpi('Employees in file', '—', '') + kpi('Total hours', '—', '') + kpi('Findings', '—', 'run the check after loading');
    return;
  }
  var hrs = 0; wk.records.forEach(function(r){ Object.keys(r.days).forEach(function(i){ var h = r.days[i]; hrs += h.rt + h.ot + h.dt; }); });
  var worked = wk.records.filter(function(r){ return Object.keys(r.days).length; }).length;
  var nCrit = f ? f.items.filter(function(i){ return i.kind === 'mismatch'; }).length : null;
  el.innerHTML =
    kpi('Week ' + (wk.weekNo || '?'), fmtMD(wk.start) + '–' + fmtMD(wk.end), 'week number can be changed below') +
    kpi('Employees with hours', worked, 'of ' + wk.records.length + ' in the report') +
    kpi('Total hours', hrsRound(hrs).toLocaleString(), 'RT + OT + DT') +
    (f ? kpi('Rule differences', nCrit, (f.rowsChecked ? f.rowsChecked + ' employee-days checked' : 'no days could be checked'), nCrit ? 'up' : 'flat')
       : kpi('Rule differences', '—', 'check not run yet'));
}

function renderHoursSaved(){
  var el = document.getElementById('hoursSaved'); if(!el) return;
  if(hrsSavedError){
    el.innerHTML = '<div class="emp-status err">Saved weeks could not be read from SharePoint (' + escapeHtml(hrsSavedError) + '). Check that the WeeklyHours list exists. Loading a file and the rule check still work.</div>';
    return;
  }
  if(!hrsLoadedOnce || hrsSaved == null){ el.innerHTML = '<div class="emp-status" id="hrsSavedStatus">Saved weeks load when you open this screen after signing in.</div>'; return; }
  var weeks = hrsSavedWeeks();
  var html = '<div class="hrs-week"><label>Saved weeks <select id="hrsSavedSel">' +
    (weeks.length ? weeks.map(function(w){ return '<option value="' + w.start + '"' + (hrsWeek && hrsWeek.start === w.start ? ' selected' : '') + '>Week ' + (w.weekNo || '?') + ' · ' + fmtMD(w.start) + '–' + fmtMD(hrsAddDays(w.start, 6)) + ' · ' + w.n + ' employees</option>'; }).join('') : '<option value="">(none saved yet)</option>') +
    '</select></label><button class="btn" type="button" id="hrsOpenSavedBtn"' + (weeks.length ? '' : ' disabled') + '>Open saved week</button>' +
    '<button class="btn ghost" type="button" id="hrsRefreshBtn">Refresh</button></div>';
  if(hrsFindError) html += '<div class="emp-status err">WeeklyFindings could not be read (' + escapeHtml(hrsFindError) + '). Findings can be checked and emailed, but review status will not be remembered until that list exists.</div>';
  el.innerHTML = html;
  var ob = document.getElementById('hrsOpenSavedBtn');
  if(ob) ob.addEventListener('click', function(){ var v = document.getElementById('hrsSavedSel').value; if(v) hrsOpenSaved(v); });
  document.getElementById('hrsRefreshBtn').addEventListener('click', hrsRefreshSaved);
}

function renderHoursPanel(){
  var el = document.getElementById('hoursWork'); if(!el) return;
  var wk = hrsWeek;
  if(!wk){ el.innerHTML = ''; renderHoursKpis(); renderHoursFindings(); renderHoursDash(); renderHoursSummary(); return; }
  var plan = hrsSaved ? hrsPlanSave() : null;
  var html = '<div class="hrs-week">' +
    '<label>Week no. <input type="number" id="hrsWeekNo" min="1" max="53" value="' + (wk.weekNo || '') + '"></label>' +
    '<label>Start <input type="date" id="hrsStart" value="' + wk.start + '"' + (wk.fromSaved ? ' disabled' : '') + '></label>' +
    '<label>End <input type="date" id="hrsEnd" value="' + wk.end + '"' + (wk.fromSaved ? ' disabled' : '') + '></label>' +
    '<span class="panel-sub">' + (wk.fromSaved ? 'Opened from SharePoint.' : (wk.hadRange ? 'Dates read from the report header.' : 'Dates taken from the day columns.')) + ' Week no. is the ISO week &mdash; change it if Payroll numbers weeks differently.</span></div>';
  html += '<div class="emp-sum"><span class="emp-pill"><b>' + wk.records.length + '</b>employees</span>' +
    (wk.dupIds.length ? '<span class="emp-pill chg"><b>' + wk.dupIds.length + '</b>repeated IDs skipped</span>' : '') +
    (plan ? '<span class="emp-pill new"><b>' + plan.fresh.length + '</b>new day rows</span><span class="emp-pill chg"><b>' + plan.changed.length + '</b>changed</span><span class="emp-pill"><b>' + plan.same + '</b>already saved</span>' : '') +
    '</div>';
  if(plan && plan.changed.length){
    html += '<details class="emp-details"><summary>Rows that differ from what was saved (' + plan.changed.length + ')</summary><div class="emp-detail-body">';
    plan.changed.slice(0, 100).forEach(function(c){
      html += '<div><span class="mono">' + escapeHtml(c.next.employeeId) + '</span> ' + hrsPretty(c.next.workDate) + ': <i>' + hrsFmt({ rt: c.cur.rt, ot: c.cur.ot, dt: c.cur.dt }) + '</i> &rarr; <b>' + hrsFmt({ rt: c.next.rt, ot: c.next.ot, dt: c.next.dt }) + '</b></div>';
    });
    html += '</div></details>';
  }
  if(plan && plan.gone.length) html += '<div class="emp-status">' + plan.gone.length + ' day row(s) saved earlier for this week are not in this file. They are left alone.</div>';
  var nSave = plan ? plan.fresh.length + plan.changed.length : 0;
  var canSave = !!hrsSaved && !hrsSavedError && !hrsBusy;
  html += '<div class="emp-actions">' +
    '<button class="btn primary" type="button" id="hrsCheckBtn">Run rule check</button>' +
    '<button class="btn" type="button" id="hrsSaveBtn"' + (canSave ? '' : ' disabled') + '>Save week ' + (wk.weekNo || '') + ' to SharePoint</button>' +
    '<button class="btn ghost" type="button" id="hrsDiscardBtn">Close week</button></div>' +
    '<div class="emp-status" id="hrsStatus">' + (plan && !nSave && hrsSaved ? 'Hours for this week are already saved.' : '') + '</div>';
  el.innerHTML = html;
  document.getElementById('hrsCheckBtn').addEventListener('click', function(){ hrsFindings = hrsRunCheck(); hrsPage = 1; renderHoursKpis(); renderHoursFindings(); renderHoursDash(); renderHoursSummary(); });
  document.getElementById('hrsSaveBtn').addEventListener('click', hrsSave);
  document.getElementById('hrsDiscardBtn').addEventListener('click', function(){ hrsWeek = null; hrsFindings = null; hrsSel = {}; document.getElementById('hrsFileName').textContent = 'No file chosen'; renderHoursPanel(); renderHoursSaved(); });
  document.getElementById('hrsWeekNo').addEventListener('change', function(){ wk.weekNo = parseInt(this.value, 10) || null; renderHoursKpis(); });
  document.getElementById('hrsStart').addEventListener('change', function(){ if(this.value){ wk.start = this.value; hrsFindings = null; renderHoursPanel(); } });
  document.getElementById('hrsEnd').addEventListener('change', function(){ if(this.value) wk.end = this.value; });
  renderHoursKpis();
  renderHoursFindings();
  renderHoursDash();
  renderHoursSummary();
}

// Phase 3a: what happened to similar findings before (same union, rule and kind, already resolved with a reason). Counts only, no AI.
function hrsSimilarOutcomes(i){
  if(i.status === STATUS_DONE || !i.ruleName || i.ruleId === 'Data check') return null;
  var c = {}, n = 0;
  (hrsFindSaved || []).forEach(function(r){
    if(r.key === i.key || r.status !== STATUS_DONE || !r.resReason) return;
    if(r.kind !== i.kind || (r.union || '') !== (i.union || '') || (r.ruleName || '') !== (i.ruleName || '')) return;
    c[r.resReason] = (c[r.resReason] || 0) + 1; n++;
  });
  if(n < 3) return null;
  var parts = Object.keys(c).map(function(k){ return [k, c[k]]; }).sort(function(a, b){ return b[1] - a[1]; });
  return { n: n, parts: parts.slice(0, 3), top: parts[0] };
}
function similarHtml(i){
  var o = hrsSimilarOutcomes(i); if(!o) return '';
  return '<div class="h-msg" style="font-size:12px;color:var(--ink-faint);">Past outcomes for this rule and union (' + o.n + ' resolved): ' + o.parts.map(function(p){ return escapeHtml(p[0]) + ' ' + p[1]; }).join(', ') + '. History only, not a decision.</div>';
}
function itemLine(i){ if(i.status === STATUS_DONE && i.resReason) return 'Resolved · ' + i.resReason + (i.resolvedBy ? ' · ' + i.resolvedBy : '') + (i.resolvedDate ? ' · ' + i.resolvedDate : '') + (i.note ? ' — ' + i.note : ''); return i.status === STATUS_SENT ? 'Sent to PM' + (i.sentTo ? ' · ' + i.sentTo : '') + (i.sentDate ? ' · ' + i.sentDate : '') + (i.pmResponse ? ' — PM answered: ' + i.pmResponse + (i.pmNote ? ' (' + i.pmNote + ')' : '') + (i.pmBy ? ' · ' + i.pmBy : '') : '') : i.status; }
function visibleFindings(){
  var f = hrsFindings; if(!f) return [];
  return f.items.filter(function(i){ return (hrsFilter === 'all' || i.kind === hrsFilter) && (hrsStatusFilter === 'all' || (hrsStatusFilter === 'pm' ? (i.pmResponse && i.status !== STATUS_DONE) : i.status === hrsStatusFilter)); })
    .sort(function(a, b){ return (a.sev === b.sev ? 0 : a.sev === 'crit' ? -1 : 1) || a.id.localeCompare(b.id) || (a.iso || '').localeCompare(b.iso || ''); });
}
function selectedItems(){ return hrsFindings ? hrsFindings.items.filter(function(i){ return hrsSel[i.key]; }) : []; }

function renderHoursFindings(){
  var el = document.getElementById('hoursFindings'); if(!el) return;
  var f = hrsFindings;
  if(!f){ el.innerHTML = hrsWeek ? '<div class="emp-status">Press <b>Run rule check</b> to compare each union employee&rsquo;s daily hours with their union&rsquo;s rule.</div>' : '<div class="emp-status">Load a week above to see findings.</div>'; return; }
  var counts = {}; f.items.forEach(function(i){ counts[i.kind] = (counts[i.kind] || 0) + 1; });
  var sc = {}; f.items.forEach(function(i){ sc[i.status] = (sc[i.status] || 0) + 1; if(i.pmResponse && i.status !== STATUS_DONE) sc.pm = (sc.pm || 0) + 1; });
  var e = f.emp;
  var html = '<div class="emp-sum">' +
    '<span class="emp-pill new"><b>' + e.checked + '</b>employees checked against a rule</span>' +
    '<span class="emp-pill"><b>' + f.matched + '</b>days match the rule</span>' +
    '<span class="emp-pill' + (counts.mismatch ? ' chg' : '') + '"><b>' + (counts.mismatch || 0) + '</b>days differ</span>' +
    '<span class="emp-pill"><b>' + e.regular + '</b>regular employees (no union rule)</span>' +
    (e.noUnion.length ? '<span class="emp-pill"><b>' + e.noUnion.length + '</b>labor type has no union</span>' : '') + '</div>';
  var nr = Object.keys(e.noRule);
  if(nr.length){
    html += '<details class="emp-details"><summary>Unions that could not be checked (' + nr.length + ') &mdash; ' + nr.reduce(function(s, k){ return s + e.noRule[k].n; }, 0) + ' employees</summary><div class="emp-detail-body">' +
      nr.sort().map(function(k){ return '<div><b>' + escapeHtml(k || '(no union)') + '</b> &middot; ' + e.noRule[k].n + ' employee(s) &mdash; ' + escapeHtml(e.noRule[k].why) + '</div>'; }).join('') + '</div></details>';
  }
  if(f.stale) html += '<div class="emp-status">' + f.stale + ' finding(s) saved for this week are no longer flagged under today&rsquo;s rules.</div>';
  html += '<div class="filters" style="padding:4px 18px 6px;"><button class="chip-filter' + (hrsFilter === 'all' ? ' active' : '') + '" type="button" data-hrs-filter="all">All (' + f.items.length + ')</button>' +
    Object.keys(HRS_KINDS).filter(function(k){ return counts[k]; }).map(function(k){
      return '<button class="chip-filter' + (hrsFilter === k ? ' active' : '') + '" type="button" data-hrs-filter="' + k + '">' + HRS_KINDS[k] + ' (' + counts[k] + ')</button>';
    }).join('') + '</div>' +
    '<div class="filters" style="padding:0 18px 10px;">' +
    ['all', STATUS_OPEN, STATUS_SENT, 'pm', STATUS_DONE].map(function(s){
      return '<button class="chip-filter' + (hrsStatusFilter === s ? ' active' : '') + '" type="button" data-hrs-status="' + s + '">' + (s === 'all' ? 'Any status' : (s === 'pm' ? 'PM answered' : s) + ' (' + (sc[s] || 0) + ')') + '</button>';
    }).join('') + '</div>';
  var list = visibleFindings();
  var sel = selectedItems();
  html += '<div class="hrs-bar"><label class="hrs-all"><input type="checkbox" id="hrsSelAll"> Select all shown (' + list.length + ')</label>' +
    '<span class="panel-sub">' + sel.length + ' selected</span>' +
    '<button class="btn primary" type="button" id="hrsMailBtn"' + (sel.length ? '' : ' disabled') + '>Email PM&hellip;</button>' +
    '<button class="btn" type="button" id="hrsDoneBtn"' + (sel.length ? '' : ' disabled') + '>Mark resolved</button>' +
    '<button class="btn ghost" type="button" id="hrsReopenBtn"' + (sel.length ? '' : ' disabled') + '>Reopen</button>' +
    '<button class="btn ghost" type="button" id="hrsListToggle">' + (hrsListHidden ? 'Show list' : 'Hide list') + '</button><button class="btn ghost" type="button" id="hrsAiShowAll">Show AI</button><button class="btn ghost" type="button" id="hrsAiHideAll">Hide AI</button>' +
    '<button class="btn ghost" type="button" id="hrsCsvBtn" style="margin-left:auto;">Download CSV</button></div>' +
    '<div class="emp-status" id="hrsActionStatus"></div>';
  if(!list.length) html += '<div class="emp-status">' + (f.items.length ? 'Nothing in this filter.' : 'No findings. Every checked day matches its rule.') + '</div>';
  else {
    var pages = Math.max(1, Math.ceil(list.length / hrsPageSize)); if(hrsPage > pages) hrsPage = pages; if(hrsPage < 1) hrsPage = 1;
    var pager = pagerHtml('hrs', list.length, hrsPage, hrsPageSize);
    html += hrsListHidden ? '<div class="emp-status">The list is hidden (' + list.length + ' finding' + (list.length === 1 ? '' : 's') + '). Press <b>Show list</b> to see it.</div>' : pager;
    if(!hrsListHidden) html += '<div class="hrs-row head"><div></div><div>ID</div><div>Employee</div><div>Day</div><div>Reported</div><div>Rule says</div><div>Rule broken</div></div>';
    if(!hrsListHidden) list.slice((hrsPage - 1) * hrsPageSize, hrsPage * hrsPageSize).forEach(function(i){
      html += '<div class="hrs-row ' + i.sev + '"><div><input type="checkbox" data-hrs-sel="' + escapeHtml(i.key) + '"' + (hrsSel[i.key] ? ' checked' : '') + '></div>' +
        '<div class="e-id">' + escapeHtml(i.id) + '</div><div class="e-name">' + escapeHtml(i.name) + (i.union ? '<span class="e-sub">' + escapeHtml(i.union) + '</span>' : '') + '</div>' +
        '<div>' + (i.iso ? hrsPretty(i.iso) : '—') + (i.jobs ? '<span class="e-sub">' + escapeHtml(i.jobs) + (i.jobCount > 1 ? ' · more than one job' : '') + '</span>' : '') + '</div><div>' + (i.rep ? hrsFmt(i.rep) : '—') + '</div><div>' + (i.exp ? hrsFmt(i.exp) : '—') + '</div>' +
        '<div class="h-msg"><b>' + escapeHtml(i.ruleId === 'Data check' ? 'Data check' : i.ruleId + ' · ' + i.ruleName) + '</b>' + (i.cba ? ' <span class="e-sub" style="display:inline;">' + escapeHtml(i.cba) + '</span>' : '') +
        '<div>' + escapeHtml(i.broken.join(' ')) + '</div>' + similarHtml(i) + '<div class="h-stat ' + (i.status === STATUS_DONE ? 'done' : i.status === STATUS_SENT ? 'sent' : '') + '">' + escapeHtml(itemLine(i)) + '</div>' +
        '<div class="hrs-ai-wrap"><button class="btn ghost hrs-ai-btn" type="button" data-hrs-ai="' + escapeHtml(i.key) + '"' + (hrsAi[i.key] && hrsAi[i.key].state === 'loading' ? ' disabled' : '') + '>' + (hrsAi[i.key] && hrsAi[i.key].state === 'done' ? 'Explain again' : 'Explain with AI') + '</button></div>' +
        '<div class="hrs-ai-box" data-hrs-ai-box="' + escapeHtml(i.key) + '">' + aiBoxHtml(i.key) + '</div></div></div>';
    });
    if(!hrsListHidden) html += pager;
  }
  el.innerHTML = html;
  Array.prototype.forEach.call(el.querySelectorAll('[data-hrs-filter]'), function(b){
    b.addEventListener('click', function(){ hrsFilter = b.getAttribute('data-hrs-filter'); hrsPage = 1; renderHoursFindings(); });
  });
  Array.prototype.forEach.call(el.querySelectorAll('[data-hrs-status]'), function(b){
    b.addEventListener('click', function(){ hrsStatusFilter = b.getAttribute('data-hrs-status'); hrsPage = 1; renderHoursFindings(); });
  });
  Array.prototype.forEach.call(el.querySelectorAll('[data-hrs-sel]'), function(c){
    c.addEventListener('change', function(){ var k = c.getAttribute('data-hrs-sel'); if(c.checked) hrsSel[k] = true; else delete hrsSel[k]; refreshSelBar(); });
  });
  var all = document.getElementById('hrsSelAll');
  all.checked = list.length > 0 && list.every(function(i){ return hrsSel[i.key]; });
  all.addEventListener('change', function(){ list.forEach(function(i){ if(all.checked) hrsSel[i.key] = true; else delete hrsSel[i.key]; }); renderHoursFindings(); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-hrs-ai]'), function(b){ b.addEventListener('click', function(){ hrsExplain(b.getAttribute('data-hrs-ai')); }); });
  wireAiBoxes(el);
  pagerWire(el, 'hrs', function(pg, size){ if(size){ hrsPageSize = size; hrsPage = 1; } else hrsPage = pg; renderHoursFindings(); el.scrollIntoView({ block: 'start' }); });
  document.getElementById('hrsListToggle').addEventListener('click', function(){ hrsListHidden = !hrsListHidden; renderHoursFindings(); });
  function setAllAi(c){ Object.keys(hrsAi).forEach(function(k){ if(hrsAi[k].state === 'done') hrsAi[k].collapsed = c; }); renderHoursFindings(); }
  document.getElementById('hrsAiShowAll').addEventListener('click', function(){ setAllAi(false); });
  document.getElementById('hrsAiHideAll').addEventListener('click', function(){ setAllAi(true); });
  document.getElementById('hrsCsvBtn').addEventListener('click', hrsDownloadCsv);
  document.getElementById('hrsMailBtn').addEventListener('click', openMailModal);
  document.getElementById('hrsDoneBtn').addEventListener('click', openResolveModal);
  document.getElementById('hrsReopenBtn').addEventListener('click', function(){ hrsSetStatus(STATUS_OPEN); });
}
function refreshSelBar(){
  var n = selectedItems().length;
  ['hrsMailBtn', 'hrsDoneBtn', 'hrsReopenBtn'].forEach(function(id){ var b = document.getElementById(id); if(b) b.disabled = !n; });
  var sp = document.querySelector('.hrs-bar .panel-sub'); if(sp) sp.textContent = n + ' selected';
}
async function hrsSetStatus(status){
  var items = selectedItems(); if(!items.length) return;
  items.forEach(function(i){ i.status = status; if(status === STATUS_OPEN){ i.sentTo = ''; i.sentDate = ''; i.resReason = ''; i.resolvedBy = ''; i.resolvedDate = ''; } });
  hrsSel = {};
  renderHoursFindings(); renderHoursDash();
  var r = await hrsPersistFindings(items);
  var st = document.getElementById('hrsActionStatus');
  if(st) st.textContent = r.cancelled ? r.bad[0].error : r.bad.length ? 'Updated on screen, but not saved to SharePoint: ' + (r.bad[0].error || 'unknown error') : 'Marked ' + items.length + ' finding(s) as ' + status + '.';
  if(st && r.bad.length) st.className = 'emp-status err';
}
function hrsDownloadCsv(){
  var f = hrsFindings; if(!f) return;
  function q(v){ v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  var lines = [['Week', 'Employee ID', 'Employee', 'Union', 'Day', 'Finding', 'Reported', 'Rule says', 'Rule', 'Rule broken', 'Status', 'Sent to', 'Sent date', 'Resolution reason', 'Note', 'Resolved by', 'Resolved date', 'Jobs'].join(',')];
  f.items.forEach(function(i){
    lines.push([hrsWeek.weekNo, i.id, i.name, i.union, i.iso, HRS_KINDS[i.kind], i.rep ? hrsFmt(i.rep) : '', i.exp ? hrsFmt(i.exp) : '', i.ruleName, i.broken.join(' '), i.status, i.sentTo, i.sentDate, i.resReason || '', i.note || '', i.resolvedBy || '', i.resolvedDate || '', i.jobs || ''].map(q).join(','));
  });
  var blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = 'UPAE_week_' + (hrsWeek.weekNo || 'x') + '_findings.csv';
  document.body.appendChild(a); a.click(); setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 500);
}


// ---------------- AI: explain a finding (Phase 1) ----------------
function aiBoxHtml(key){
  var a = hrsAi[key]; if(!a) return '';
  if(a.state === 'loading') return '<div class="ai-card"><span class="ai-tag">AI</span> Thinking&hellip;</div>';
  if(a.state === 'error') return '<div class="ai-card err"><span class="ai-tag">AI</span> ' + escapeHtml(a.err) + '</div>';
  var d = a.data;
  if(a.collapsed) return '<div class="ai-card ai-collapsed"><span class="ai-tag">AI</span><span class="ai-conf">AI explanation &middot; confidence: ' + escapeHtml(d.confidence) + '</span><button class="btn ghost" type="button" data-hrs-ai-toggle="' + escapeHtml(key) + '">Show</button></div>';
  return '<div class="ai-card"><div class="ai-head"><span class="ai-tag">AI</span><span class="ai-conf">Confidence: ' + escapeHtml(d.confidence) + '</span>' +
    (d.stub ? '<span class="ai-stub">Sample answer &mdash; the AI key is not set up yet</span>' : '') + '<button class="btn ghost ai-hide" type="button" data-hrs-ai-toggle="' + escapeHtml(key) + '">Hide</button></div>' +
    '<p><b>What happened.</b> ' + escapeHtml(d.explanation) + '</p>' +
    '<p><b>Likely cause.</b> ' + escapeHtml(d.likelyCause) + '</p>' +
    '<p><b>Suggested action.</b> ' + escapeHtml(d.suggestedAction) + '</p>' +
    (d.needsMoreInfo ? '<p class="ai-warn">The AI says it needs more information to be sure.</p>' : '') +
    (d.evidence && d.evidence.length ? '<details><summary>Rule text it relied on</summary>' + d.evidence.map(function(x){ return '<div class="ai-ev">' + escapeHtml(x) + '</div>'; }).join('') + '</details>' : '') +
    '<div class="ai-fb">Was this helpful? <button class="btn ghost" type="button" data-hrs-ai-fb="up" data-key="' + escapeHtml(key) + '"' + (a.feedback ? ' disabled' : '') + '>' + (a.feedback === 'Helpful' ? 'Helpful (thanks)' : 'Helpful') + '</button> ' +
    '<button class="btn ghost" type="button" data-hrs-ai-fb="down" data-key="' + escapeHtml(key) + '"' + (a.feedback ? ' disabled' : '') + '>' + (a.feedback === 'Not helpful' ? 'Not helpful (noted)' : 'Not helpful') + '</button></div>' +
    '<div class="ai-note">AI-generated from the rule text and this record. The rule check above is the source of truth.' + (a.saveNote ? ' ' + escapeHtml(a.saveNote) : '') + (d.model && d.model !== 'stub' ? ' Model: ' + escapeHtml(d.model) + '.' : '') + '</div></div>';
}
function wireAiBoxes(root){
  Array.prototype.forEach.call(root.querySelectorAll('[data-hrs-ai-toggle]'), function(b){
    b.addEventListener('click', function(){ var k = b.getAttribute('data-hrs-ai-toggle'); if(hrsAi[k]){ hrsAi[k].collapsed = !hrsAi[k].collapsed; refreshAiBox(k); } });
  });
  Array.prototype.forEach.call(root.querySelectorAll('[data-hrs-ai-fb]'), function(b){
    b.addEventListener('click', function(){ hrsAiFeedback(b.getAttribute('data-key'), b.getAttribute('data-hrs-ai-fb') === 'up' ? 'Helpful' : 'Not helpful'); });
  });
}
function refreshAiBox(key){
  var box = null;
  Array.prototype.forEach.call(document.querySelectorAll('[data-hrs-ai-box]'), function(b){ if(b.getAttribute('data-hrs-ai-box') === key) box = b; });
  if(!box) return;
  box.innerHTML = aiBoxHtml(key); wireAiBoxes(box);
  var btn = null;
  Array.prototype.forEach.call(document.querySelectorAll('[data-hrs-ai]'), function(b){ if(b.getAttribute('data-hrs-ai') === key) btn = b; });
  if(btn){ var a = hrsAi[key]; btn.disabled = !!(a && a.state === 'loading'); btn.textContent = a && a.state === 'done' ? 'Explain again' : 'Explain with AI'; }
}
async function hrsExplain(key){
  var item = hrsFindings && hrsFindings.items.filter(function(i){ return i.key === key; })[0];
  if(!item || (hrsAi[key] && hrsAi[key].state === 'loading')) return;
  var rule = item.kind === 'mismatch' ? hrsFindRule(item.union) : null;
  var rec = hrsWeek.records.filter(function(r){ return empKey(r.id) === empKey(item.id); })[0];
  var pattern = rec ? Object.keys(rec.days).sort().map(function(iso){ var h = rec.days[iso]; return { day: hrsPretty(iso), rt: h.rt, ot: h.ot, dt: h.dt }; }) : [];
  var payload = {
    week: { weekNo: hrsWeek.weekNo, start: hrsWeek.start, end: hrsWeek.end },
    finding: { kind: item.kind, id: item.id, name: item.name, union: item.union, iso: item.iso || '', dayName: item.iso ? hrsPretty(item.iso).split(' ')[0] : '',
               reported: item.rep ? hrsFmt(item.rep) : '', expected: item.exp ? hrsFmt(item.exp) : '', jobs: item.jobs || '', msg: item.msg, broken: item.broken },
    rule: rule ? { id: rule.id, name: rule.name, cba: rule.cba, fields: rule.fields } : null,
    weekPattern: pattern
  };
  hrsAi[key] = { state: 'loading' }; refreshAiBox(key);
  try{
    if(typeof upaeAiExplain !== 'function') throw new Error('The AI service is not available in this view.');
    var data = await upaeAiExplain(payload);
    hrsAi[key] = { state: 'done', data: data };
    refreshAiBox(key);
    if(!data.stub){
      item.ai = { explanation: data.explanation, likelyCause: data.likelyCause, suggestedAction: data.suggestedAction, evidence: data.evidence || [],
                  confidence: data.confidence, model: data.model, date: hrsIso(new Date()), feedback: '' };
      hrsPersistFindings([item]).then(function(r){
        if(r.bad.length && hrsAi[key]) { hrsAi[key].saveNote = 'Explanation not saved to the finding: ' + (r.bad[0].error || 'unknown error'); refreshAiBox(key); }
        else if(hrsAi[key]) { hrsAi[key].saveNote = 'Saved with the finding.'; refreshAiBox(key); }
      });
    }
    if(typeof upaeCreateAILog === 'function'){
      upaeCreateAILog({ purpose: 'explain', findingKey: key, weekStart: hrsWeek.start, userEmail: (typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : ''),
                        model: data.model, promptVersion: data.promptVersion, confidence: data.confidence, usage: data.usage,
                        output: JSON.stringify({ explanation: data.explanation, likelyCause: data.likelyCause, suggestedAction: data.suggestedAction, evidence: data.evidence, needsMoreInfo: data.needsMoreInfo, stub: !!data.stub }) })
        .then(function(it){ if(hrsAi[key]) hrsAi[key].logId = it && it.id; })
        .catch(function(err){ console.warn('UPAE: could not write AILog', err); });
    }
  } catch(err){
    console.error('UPAE: explain failed', err);
    if(typeof upaeLogFailedUsage === 'function') upaeLogFailedUsage('explain', err, hrsWeek && hrsWeek.start);
    hrsAi[key] = { state: 'error', err: err.message || String(err) }; refreshAiBox(key);
  }
}
async function hrsAiFeedback(key, value){
  var a = hrsAi[key]; if(!a || a.feedback) return;
  a.feedback = value; refreshAiBox(key);
  var it = hrsFindings && hrsFindings.items.filter(function(i){ return i.key === key; })[0];
  if(it && it.ai){ it.ai.feedback = value; hrsPersistFindings([it]).catch(function(err){ console.warn('UPAE: could not save AI feedback to the finding', err); }); }
  try{ if(a.logId && typeof upaeUpdateAILogFeedback === 'function') await upaeUpdateAILogFeedback(a.logId, value); }
  catch(err){ console.warn('UPAE: could not save AI feedback', err); }
}


// ---------------- resolve with a reason and note ----------------
var HRS_REASONS = ['Keying error', 'Approved overtime', 'Wrong day or date', 'Wrong labor type', 'Rule disagrees with practice', 'Duplicate entry', 'Other'];
var resOverlay = document.getElementById('hrsResOverlay');
function openResolveModal(){
  var items = selectedItems(); if(!items.length) return;
  document.getElementById('hrsResReason').innerHTML = '<option value="">— choose a reason —</option>' + HRS_REASONS.map(function(r){ return '<option>' + r + '</option>'; }).join('');
  document.getElementById('hrsResNote').value = '';
  document.getElementById('hrsResCount').textContent = items.length + ' finding' + (items.length === 1 ? '' : 's') + ' will be marked Resolved.';
  var st = document.getElementById('hrsResStatus'); st.textContent = ''; st.className = 'emp-status';
  resOverlay.hidden = false;
}
function closeResolveModal(){ resOverlay.hidden = true; }
async function applyResolve(){
  var reason = document.getElementById('hrsResReason').value, note = empClean(document.getElementById('hrsResNote').value);
  var st = document.getElementById('hrsResStatus');
  if(!reason){ st.className = 'emp-status err'; st.textContent = 'Choose a reason.'; return; }
  if(reason === 'Other' && !note){ st.className = 'emp-status err'; st.textContent = 'Add a short note when the reason is Other.'; return; }
  var items = selectedItems(), who = (typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : ''), today = hrsIso(new Date());
  items.forEach(function(i){ i.status = STATUS_DONE; i.resReason = reason; i.note = note || i.note || ''; i.resolvedBy = who; i.resolvedDate = today; });
  hrsSel = {};
  closeResolveModal();
  renderHoursFindings(); renderHoursDash();
  var r = await hrsPersistFindings(items);
  var st2 = document.getElementById('hrsActionStatus');
  if(st2){ st2.className = 'emp-status' + (r.bad.length ? ' err' : ''); st2.textContent = r.cancelled ? r.bad[0].error : r.bad.length ? 'Marked on screen, but not saved to SharePoint: ' + (r.bad[0].error || 'unknown error') + ' (the WeeklyFindings list needs the ResolutionReason, ResolvedBy and ResolvedDate columns).' : 'Resolved ' + items.length + ' finding(s): ' + reason + '.'; }
}
document.getElementById('hrsResApply').addEventListener('click', applyResolve);
document.getElementById('hrsResCancel').addEventListener('click', closeResolveModal);
document.getElementById('hrsResClose').addEventListener('click', closeResolveModal);

// ---------------- rule coverage ----------------
function hrsCoverageRows(){
  var onFile = {}, week = {}, weekHours = {}, noUnionHours = 0, totalHours = 0;
  var byId = {}; employees.forEach(function(e){ byId[empKey(e.id)] = e; var u = empUnionName(e); if(u) onFile[u] = (onFile[u] || 0) + 1; });
  if(hrsWeek) hrsWeek.records.forEach(function(r){
    var h = 0; Object.keys(r.days).forEach(function(i){ var d = r.days[i]; h += d.rt + d.ot + d.dt; });
    if(!h) return;
    totalHours += h;
    var e = byId[empKey(r.id)], u = e ? empUnionName(e) : '';
    if(u){ week[u] = (week[u] || 0) + 1; weekHours[u] = (weekHours[u] || 0) + h; } else noUnionHours += h;
  });
  var names = {}; unions.forEach(function(u){ names[u.name] = true; }); Object.keys(onFile).forEach(function(n){ names[n] = true; });
  var rows = Object.keys(names).map(function(n){
    var rule = hrsFindRule(n), comp = hrsCompileRule(rule);
    var status, why;
    if(!rule){ status = 'norule'; why = 'No rule in the Rule Library is named "' + n + '".'; }
    else if(comp.usable){ status = 'ok'; why = 'Rule ' + rule.id + ' is usable.'; }
    else { status = /not marked Structured|not structured/i.test(comp.why) ? 'draft' : 'unreadable'; why = comp.why; }
    return { union: n, onFile: onFile[n] || 0, week: week[n] || 0, hours: weekHours[n] || 0, status: status, why: why };
  });
  rows.sort(function(a, b){ return (b.week - a.week) || (b.onFile - a.onFile) || a.union.localeCompare(b.union); });
  var unionHours = 0, checkedHours = 0; rows.forEach(function(r){ unionHours += r.hours; if(r.status === 'ok') checkedHours += r.hours; });
  return { rows: rows, unionHours: unionHours, checkedHours: checkedHours, totalHours: totalHours, noUnionHours: noUnionHours };
}
var COV_LABEL = { ok: 'Usable', draft: 'Not Structured yet', unreadable: 'Cannot read rule', norule: 'No rule' };
function renderHoursCoverage(){
  var el = document.getElementById('hoursCoverage'); if(!el) return;
  var c = hrsCoverageRows(), usable = c.rows.filter(function(r){ return r.status === 'ok'; }).length;
  var html = '<div class="emp-sum"><span class="emp-pill new"><b>' + usable + '</b>of ' + c.rows.length + ' unions have a usable rule</span>' +
    (hrsWeek ? '<span class="emp-pill' + (c.unionHours && c.checkedHours < c.unionHours ? ' chg' : '') + '"><b>' + (c.unionHours ? Math.round(c.checkedHours / c.unionHours * 100) : 0) + '%</b>of this week&rsquo;s union hours can be checked</span>' +
      '<span class="emp-pill"><b>' + Math.round(c.noUnionHours).toLocaleString() + '</b>hrs with no union (regular or unassigned)</span>' : '<span class="emp-pill">Load a week to see hours covered</span>') + '</div>';
  html += '<div class="cov-row head"><div>Union</div><div>Employees</div><div>This week</div><div>Rule status</div><div>What is needed</div></div>';
  var covPages = Math.max(1, Math.ceil(c.rows.length / covPageSize)); if(covPage > covPages) covPage = covPages; if(covPage < 1) covPage = 1;
  c.rows.slice((covPage - 1) * covPageSize, covPage * covPageSize).forEach(function(r){
    html += '<div class="cov-row"><div class="e-name">' + escapeHtml(r.union) + '</div><div>' + r.onFile + '</div><div>' + (hrsWeek ? r.week + ' emp · ' + Math.round(r.hours).toLocaleString() + ' hrs' : '—') + '</div>' +
      '<div><span class="cov-pill ' + r.status + '">' + COV_LABEL[r.status] + '</span></div><div class="h-msg">' + (r.status === 'ok' ? '—' : escapeHtml(r.why)) + '</div></div>';
  });
  if(c.rows.length > 10) html += pagerHtml('cov', c.rows.length, covPage, covPageSize);
  el.innerHTML = html;
  pagerWire(el, 'cov', function(pg, size){ if(size){ covPageSize = size; covPage = 1; } else covPage = pg; renderHoursCoverage(); });
}

// ---------------- dashboard ----------------
function hbars(rows, cls){
  var max = Math.max.apply(null, rows.map(function(r){ return r.v; }).concat([1]));
  return rows.map(function(r){
    return '<div class="bar-row"><div class="b-lab" title="' + escapeHtml(r.l) + '">' + escapeHtml(r.l) + '</div><div class="b-track"><div class="b-fill ' + (r.cls || cls || '') + '" style="width:' + Math.max(2, Math.round(r.v / max * 100)) + '%"></div></div><div class="b-val">' + r.v + (r.sfx || '') + '</div></div>';
  }).join('');
}
function renderHoursDash(){
  renderHoursCoverage();
  var el = document.getElementById('hoursDash'); if(!el) return;
  var wk = hrsWeek, f = hrsFindings;
  if(!wk){
    var weeks0 = hrsSavedWeeks();
    el.innerHTML = weeks0.length ? '<div class="emp-status">Open a saved week or load a file to see this week&rsquo;s dashboard.</div>' + dashTrend() : '<div class="emp-status">The dashboard fills in once a week is loaded.</div>';
    return;
  }
  // hours mix by day
  var perDay = wk.days.map(function(iso){ var s = { iso: iso, rt: 0, ot: 0, dt: 0 }; wk.records.forEach(function(r){ var h = r.days[iso]; if(h){ s.rt += h.rt; s.ot += h.ot; s.dt += h.dt; } }); return s; });
  var maxDay = Math.max.apply(null, perDay.map(function(d){ return d.rt + d.ot + d.dt; }).concat([1]));
  var tRt = 0, tOt = 0, tDt = 0; perDay.forEach(function(d){ tRt += d.rt; tOt += d.ot; tDt += d.dt; });
  var tot = tRt + tOt + tDt || 1;
  var html = '<div class="dash-stats">' +
    '<div class="dstat"><b>' + hrsRound(tRt).toLocaleString() + '</b><span>RT hrs · ' + Math.round(tRt / tot * 100) + '%</span></div>' +
    '<div class="dstat"><b>' + hrsRound(tOt).toLocaleString() + '</b><span>OT (1.5x) hrs · ' + Math.round(tOt / tot * 100) + '%</span></div>' +
    '<div class="dstat"><b>' + hrsRound(tDt).toLocaleString() + '</b><span>DT (2x) hrs · ' + Math.round(tDt / tot * 100) + '%</span></div></div>';
  html += '<div class="dash-grid">';
  html += '<div class="dash-card"><h3>Hours by day</h3><div class="vstack">' + perDay.map(function(d){
    var t = d.rt + d.ot + d.dt;
    return '<div class="vcol" title="' + hrsPretty(d.iso) + ': ' + hrsRound(t) + ' hrs"><div class="vbar" style="height:' + Math.round(t / maxDay * 100) + '%">' +
      '<i class="s-dt" style="flex:' + d.dt + '"></i><i class="s-ot" style="flex:' + d.ot + '"></i><i class="s-rt" style="flex:' + d.rt + '"></i></div><div class="vlab">' + hrsPretty(d.iso).replace(' ', '<br>') + '</div></div>';
  }).join('') + '</div><div class="legend"><span><i class="s-rt"></i>RT</span><span><i class="s-ot"></i>OT 1.5x</span><span><i class="s-dt"></i>DT 2x</span></div></div>';
  if(f){
    var kinds = Object.keys(HRS_KINDS).map(function(k){ return { l: HRS_KINDS[k], v: f.items.filter(function(i){ return i.kind === k; }).length, cls: k === 'mismatch' ? 'crit' : 'warn' }; }).filter(function(r){ return r.v; });
    html += '<div class="dash-card"><h3>Findings by type</h3>' + (kinds.length ? hbars(kinds) : '<div class="emp-status">No findings.</div>') + '</div>';
    var unions = Object.keys(f.byUnion).map(function(u){ var b = f.byUnion[u]; return { l: u, v: b.bad, sfx: ' of ' + b.days + ' days', cls: b.bad ? 'crit' : 'good' }; }).sort(function(a, b){ return b.v - a.v; }).slice(0, 8);
    html += '<div class="dash-card"><h3>Rule differences by union</h3>' + (unions.length ? hbars(unions) : '<div class="emp-status">No union rules could be applied to this week.</div>') + '</div>';
    var ec = {}; f.items.filter(function(i){ return i.kind === 'mismatch'; }).forEach(function(i){ var o = ec[i.id] = ec[i.id] || { l: i.name + ' (' + i.id + ')', v: 0 }; o.v++; });
    var emps = Object.keys(ec).map(function(k){ return ec[k]; }).sort(function(a, b){ return b.v - a.v; }).slice(0, 8);
    html += '<div class="dash-card"><h3>Employees with the most differences</h3>' + (emps.length ? hbars(emps, 'crit') : '<div class="emp-status">None.</div>') + '</div>';
    var rc = {}; f.items.filter(function(i){ return i.kind === 'mismatch'; }).forEach(function(i){ var m = i.broken[0] || 'Other'; var key = m.replace(/\d+(\.\d+)? hrs/, 'N hrs').replace(/\(.*?\)/g, '').replace(/\bon (Sun|Mon|Tue|Wed|Thu|Fri|Sat)\b/, '').replace(/\s+/g, ' ').trim(); rc[key] = (rc[key] || 0) + 1; });
    var reasons = Object.keys(rc).map(function(k){ return { l: k.length > 70 ? k.slice(0, 68) + '…' : k, v: rc[k] }; }).sort(function(a, b){ return b.v - a.v; }).slice(0, 6);
    html += '<div class="dash-card wide"><h3>Most common ways the rules were broken</h3>' + (reasons.length ? hbars(reasons, 'warn') : '<div class="emp-status">None.</div>') + '</div>';
    var open = f.items.filter(function(i){ return i.status === STATUS_OPEN; }).length, sent = f.items.filter(function(i){ return i.status === STATUS_SENT; }).length, done = f.items.filter(function(i){ return i.status === STATUS_DONE; }).length;
    var tt = open + sent + done || 1;
    html += '<div class="dash-card"><h3>Review progress</h3><div class="progress"><i class="p-open" style="width:' + open / tt * 100 + '%"></i><i class="p-sent" style="width:' + sent / tt * 100 + '%"></i><i class="p-done" style="width:' + done / tt * 100 + '%"></i></div>' +
      '<div class="legend"><span><i class="p-open"></i>Open ' + open + '</span><span><i class="p-sent"></i>Sent to PM ' + sent + '</span><span><i class="p-done"></i>Resolved ' + done + '</span></div></div>';
    var rr = {}; f.items.forEach(function(i){ if(i.resReason) rr[i.resReason] = (rr[i.resReason] || 0) + 1; });
    var rrRows = Object.keys(rr).map(function(k){ return { l: k, v: rr[k], cls: 'good' }; }).sort(function(a, b){ return b.v - a.v; });
    if(rrRows.length) html += '<div class="dash-card"><h3>Why findings were resolved</h3>' + hbars(rrRows) + '</div>';
  } else html += '<div class="dash-card"><h3>Rule results</h3><div class="emp-status">Run the rule check to fill in the rest of the dashboard.</div></div>';
  html += '</div>' + dashTrend();
  el.innerHTML = html;
}
function dashTrend(){
  var weeks = hrsSavedWeeks().slice(0, 8).reverse();
  if(!weeks.length) return '';
  var fcount = {}; (hrsFindSaved || []).filter(function(r){ return r.kind === 'mismatch'; }).forEach(function(r){ fcount[r.weekStart] = (fcount[r.weekStart] || 0) + 1; });
  var hoursRows = weeks.map(function(w){ return { l: 'Wk ' + (w.weekNo || '?') + ' · ' + fmtMD(w.start), v: Math.round(w.hours) }; });
  var diffRows = weeks.map(function(w){ return { l: 'Wk ' + (w.weekNo || '?') + ' · ' + fmtMD(w.start), v: fcount[w.start] || 0, cls: 'crit' }; });
  return '<div class="dash-grid" style="margin-top:12px;"><div class="dash-card"><h3>Total hours, saved weeks</h3>' + hbars(hoursRows) + '</div>' +
    '<div class="dash-card"><h3>Rule differences, saved weeks</h3>' + hbars(diffRows, 'crit') + '<div class="emp-status" style="padding:6px 0 0;">Counts the findings saved for each week.</div></div></div>';
}

// ---------------- weekly analyst summary (AI phase 2) ----------------
var hrsSums = null, hrsSumsError = null;   // saved summaries from SharePoint
var hrsSum = null;                          // { weekStart, state:'idle'|'loading'|'done'|'error', data, cands, id, feedback, note, saved, err, collapsed }

// Every figure here is computed the same way as the dashboard above; the AI only writes words around them.
function hrsWeekStats(){
  var wk = hrsWeek, f = hrsFindings; if(!wk || !f) return null;
  var tRt = 0, tOt = 0, tDt = 0;
  wk.days.forEach(function(iso){ wk.records.forEach(function(r){ var h = r.days[iso]; if(h){ tRt += h.rt; tOt += h.ot; tDt += h.dt; } }); });
  var tot = tRt + tOt + tDt || 1;
  var cov = hrsCoverageRows();
  var mism = f.items.filter(function(i){ return i.kind === 'mismatch'; });
  var ec = {}; mism.forEach(function(i){ var o = ec[i.id] = ec[i.id] || { employee: i.name + ' (' + i.id + ')', differences: 0 }; o.differences++; });
  var noRt = function(i){ return !!(i.rep && i.rep.rt === 0 && (i.rep.ot + i.rep.dt) > 0); };
  var rr = {}; f.items.forEach(function(i){ if(i.resReason) rr[i.resReason] = (rr[i.resReason] || 0) + 1; });
  var cnt = function(st){ return f.items.filter(function(i){ return i.status === st; }).length; };
  var prior = hrsSavedWeeks().filter(function(w){ return w.start < wk.start; }).slice(0, 4).map(function(w){
    var ot = 0, all = 0; (hrsSaved || []).forEach(function(r){ if(r.weekStart === w.start){ ot += r.ot; all += r.rt + r.ot + r.dt; } });
    return { weekNo: w.weekNo, start: w.start, totalHours: Math.round(w.hours), ruleDifferences: (hrsFindSaved || []).filter(function(r){ return r.kind === 'mismatch' && r.weekStart === w.start; }).length, otSharePct: all ? Math.round(ot / all * 100) : null };
  });
  var rank = {}; rank[STATUS_OPEN] = 0; rank[STATUS_SENT] = 1; rank[STATUS_DONE] = 2;
  var dev = function(i){ return i.rep && i.exp ? Math.abs(i.rep.rt - i.exp.rt) + Math.abs(i.rep.ot - i.exp.ot) + Math.abs(i.rep.dt - i.exp.dt) : 0; };
  var cands = f.items.slice().sort(function(a, b){ return (rank[a.status] - rank[b.status]) || ((b.kind === 'mismatch') - (a.kind === 'mismatch')) || (dev(b) - dev(a)); }).slice(0, 25);
  var candPayload = cands.map(function(i){ return { key: i.key, employee: i.name + ' (' + i.id + ')', union: i.union, day: i.iso ? hrsPretty(i.iso) : '', reported: i.rep ? hrsFmt(i.rep) : '', ruleSays: i.exp ? hrsFmt(i.exp) : '',
    kind: HRS_KINDS[i.kind] || i.kind, status: i.status, noRtDay: i.kind === 'mismatch' && noRt(i), hasAiNote: !!i.ai, repeatEmployeeCount: ec[i.id] ? ec[i.id].differences : 0 }; });
  var map = {}; candPayload.forEach(function(c){ map[c.key] = c; });
  var jb = {}, pmw = {}, noPm = {};
  f.items.forEach(function(i){
    itemJobBases(i).forEach(function(b){
      var jp = hrsJobPMs[b], o = jb[b] = jb[b] || { job: b, name: jobTitleOf(b), pm: jp ? (jp.pmName || jp.pmEmail) : '', findings: 0, open: 0, sentToPm: 0, pmAnswered: 0, resolved: 0 };
      o.findings++; if(i.status === STATUS_OPEN) o.open++; if(i.status === STATUS_SENT) o.sentToPm++; if(i.status === STATUS_DONE) o.resolved++; if(i.pmResponse && i.status !== STATUS_DONE) o.pmAnswered++;
      if(!jp) noPm[b] = 1;
      if(jp && i.status === STATUS_SENT){ var q = pmw[o.pm] = pmw[o.pm] || { pm: o.pm, waiting: 0, answered: 0 }; if(i.pmResponse) q.answered++; else q.waiting++; }
    });
  });
  var byJob = Object.keys(jb).map(function(k){ return jb[k]; }).sort(function(a, b){ return (b.open + b.sentToPm) - (a.open + a.sentToPm) || b.findings - a.findings; }).slice(0, 12);
  var pmAns = { correct: 0, needsCorrection: 0, question: 0 };
  f.items.forEach(function(i){ if(i.pmResponse && i.status !== STATUS_DONE){ var r = String(i.pmResponse).toLowerCase(); if(r.indexOf('correct as') === 0) pmAns.correct++; else if(r.indexOf('needs') === 0) pmAns.needsCorrection++; else pmAns.question++; } });
  candPayload.forEach(function(c, n){ var i = cands[n]; var bs = itemJobBases(i); c.jobs = bs.join(', '); c.pm = bs.map(function(b){ var jp = hrsJobPMs[b]; return jp ? (jp.pmName || jp.pmEmail) : ''; }).filter(Boolean).join(', '); c.pmResponse = i.pmResponse || ''; });
  return { cands: map, payload: {
    source: hrsWeek.source === 'detail' ? 'detail' : 'summary',
    byJob: hrsWeek.source === 'detail' ? byJob : [], pmWaiting: hrsWeek.source === 'detail' ? Object.keys(pmw).map(function(k){ return pmw[k]; }).sort(function(a, b){ return b.waiting - a.waiting; }).slice(0, 10) : [],
    jobsWithoutPm: hrsWeek.source === 'detail' ? Object.keys(noPm).length : 0, pmAnswers: pmAns,
    week: { weekNo: wk.weekNo, start: wk.start, end: wk.end },
    totals: { employeesInFile: wk.records.length, rt: hrsRound(tRt), ot: hrsRound(tOt), dt: hrsRound(tDt), totalHours: hrsRound(tRt + tOt + tDt), otSharePct: Math.round(tOt / tot * 100), dtSharePct: Math.round(tDt / tot * 100) },
    coverage: { employeesChecked: f.emp.checked, daysChecked: f.rowsChecked, daysMatching: f.matched, checkedHoursPct: cov.unionHours ? Math.round(cov.checkedHours / cov.unionHours * 100) : 0,
      unionsSkipped: cov.rows.filter(function(r){ return r.status !== 'ok' && r.week > 0; }).map(function(r){ return { union: r.union, employees: r.week }; }) },
    findings: {
      total: f.items.length,
      byKind: Object.keys(HRS_KINDS).map(function(k){ return { kind: HRS_KINDS[k], count: f.items.filter(function(i){ return i.kind === k; }).length }; }).filter(function(r){ return r.count; }),
      byStatus: { open: cnt(STATUS_OPEN), sentToPm: cnt(STATUS_SENT), resolved: cnt(STATUS_DONE) },
      ruleDifferencesByUnion: Object.keys(f.byUnion).map(function(u){ var b = f.byUnion[u]; return { union: u, differences: b.bad, daysChecked: b.days }; }).sort(function(a, b){ return b.differences - a.differences; }).slice(0, 8),
      employeesWithMostDifferences: Object.keys(ec).map(function(k){ return ec[k]; }).sort(function(a, b){ return b.differences - a.differences; }).slice(0, 8),
      noRtDayDifferences: mism.filter(noRt).length,
      resolvedReasons: Object.keys(rr).map(function(k){ return { reason: k, count: rr[k] }; }).sort(function(a, b){ return b.count - a.count; })
    },
    priorWeeks: prior, candidates: candPayload } };
}
function hrsSumLatestSaved(start){
  var list = (hrsSums || []).filter(function(s){ return s.weekStart === start; }).sort(function(a, b){ return String(b.at).localeCompare(String(a.at)); });
  return list[0] || null;
}
function hrsSumText(d, cands){
  var t = ['WEEK ' + (hrsWeek.weekNo || '') + ' (' + hrsWeek.start + ' to ' + hrsWeek.end + '): ' + d.headline, '', 'WHAT STANDS OUT'];
  d.keyChanges.forEach(function(x){ t.push('- ' + x); });
  t.push('', 'REVIEW FIRST');
  d.riskRanked.forEach(function(r, n){ var c = cands && cands[r.key]; t.push((n + 1) + '. [' + r.risk + '] ' + (c ? c.employee + (c.day ? ', ' + c.day : '') + ' (reported ' + c.reported + '; rule ' + c.ruleSays + ')' : r.key) + ': ' + r.reason); });
  if(d.jobFocus && d.jobFocus.length){ t.push('', 'JOBS TO LOOK AT'); d.jobFocus.forEach(function(j){ t.push('- Job ' + j.job + ': ' + j.note); }); }
  t.push('', 'SUGGESTED FOLLOW-UPS'); d.suggestedFollowUps.forEach(function(x){ t.push('- ' + x); });
  t.push('', 'LIMITS OF THIS DATA', d.dataLimits || '', '', 'AI-generated from the week\'s rule-check results. The rule check and the dashboard are the source of truth.');
  return t.join('\n');
}
function renderHoursSummary(){
  var el = document.getElementById('hoursSummary'); if(!el) return;
  if(!hrsWeek){ el.innerHTML = '<div class="emp-status">Load or open a week, run the rule check, then write the summary.</div>'; return; }
  if(hrsSum && hrsSum.weekStart !== hrsWeek.start) hrsSum = null;
  if(!hrsSum){
    var sv = hrsSumLatestSaved(hrsWeek.start);
    if(sv){ try{ var o = JSON.parse(sv.json); hrsSum = { weekStart: hrsWeek.start, state: 'done', data: Object.assign({ model: sv.model, promptVersion: sv.promptVersion }, o.data), cands: o.cands || {}, id: sv.spItemId, feedback: sv.feedback, saved: true, note: 'Saved summary from ' + (sv.at ? new Date(sv.at).toLocaleString() : 'SharePoint') + (sv.generatedBy ? ' by ' + sv.generatedBy : '') + '.' }; } catch(e){ hrsSum = null; } }
  }
  var busy = hrsSum && hrsSum.state === 'loading';
  var html = '<div class="emp-actions" style="padding:0 0 8px;"><button class="btn primary" type="button" id="hrsSumBtn"' + (!hrsFindings || busy ? ' disabled' : '') + '>' + (hrsSum && hrsSum.state === 'done' ? 'Write it again' : 'Write weekly summary') + '</button>' +
    '<span class="panel-sub">' + (hrsFindings ? 'A short AI briefing built from this week&rsquo;s rule-check results and the last saved weeks.' : 'Run the rule check first.') + '</span></div>';
  if(hrsSum && hrsSum.state === 'loading') html += '<div class="ai-card"><span class="ai-tag">AI</span> Writing the summary&hellip;</div>';
  else if(hrsSum && hrsSum.state === 'error') html += '<div class="ai-card err"><span class="ai-tag">AI</span> ' + escapeHtml(hrsSum.err) + '</div>';
  else if(hrsSum && hrsSum.state === 'done'){
    var d = hrsSum.data, cands = hrsSum.cands || {};
    if(hrsSum.collapsed) html += '<div class="ai-card ai-collapsed"><span class="ai-tag">AI</span><span class="ai-conf">Weekly summary &middot; confidence: ' + escapeHtml(d.confidence) + '</span><button class="btn ghost" type="button" id="hrsSumToggle">Show</button></div>';
    else html += '<div class="ai-card"><div class="ai-head"><span class="ai-tag">AI</span><span class="ai-conf">Confidence: ' + escapeHtml(d.confidence) + '</span>' + (d.stub ? '<span class="ai-stub">Sample answer &mdash; the AI key is not set up yet</span>' : '') +
      '<button class="btn ghost ai-hide" type="button" id="hrsSumToggle">Hide</button></div>' +
      '<p style="font-size:15px;"><b>' + escapeHtml(d.headline) + '</b></p>' +
      (d.keyChanges.length ? '<p><b>What stands out</b></p><ul class="sum-list">' + d.keyChanges.map(function(x){ return '<li>' + escapeHtml(x) + '</li>'; }).join('') + '</ul>' : '') +
      (d.riskRanked.length ? '<p><b>Review first</b></p>' + d.riskRanked.map(function(r){ var c = cands[r.key];
        return '<div class="sum-risk"><span class="cov-pill ' + (r.risk === 'high' ? 'unreadable' : r.risk === 'medium' ? 'draft' : 'ok') + '">' + escapeHtml(r.risk) + '</span> <b>' + (c ? escapeHtml(c.employee) + (c.day ? ' · ' + escapeHtml(c.day) : '') : escapeHtml(r.key)) + '</b>' +
          (c ? ' <span class="h-msg">reported ' + escapeHtml(c.reported || '—') + (c.ruleSays ? ' · rule says ' + escapeHtml(c.ruleSays) : '') + '</span>' : '') + '<div>' + escapeHtml(r.reason) + (c && c.noRtDay ? ' <i>(day with no RT: verify possible second shift or other job)</i>' : '') +
          ' <button class="btn ghost" type="button" data-sum-go="' + escapeHtml(r.key) + '" style="padding:2px 8px;font-size:12px;">Show in list</button></div></div>'; }).join('') : '') +
      (d.jobFocus && d.jobFocus.length ? '<p><b>Jobs to look at</b></p><ul class="sum-list">' + d.jobFocus.map(function(j){ return '<li><b>Job ' + escapeHtml(j.job) + '</b>' + (jobTitleOf(j.job) ? ' &middot; ' + escapeHtml(jobTitleOf(j.job)) : '') + ': ' + escapeHtml(j.note) + '</li>'; }).join('') + '</ul>' : '') +
      (d.suggestedFollowUps.length ? '<p><b>Suggested follow-ups</b></p><ul class="sum-list">' + d.suggestedFollowUps.map(function(x){ return '<li>' + escapeHtml(x) + '</li>'; }).join('') + '</ul>' : '') +
      (d.dataLimits ? '<p class="ai-warn"><b>Limits of this data.</b> ' + escapeHtml(d.dataLimits) + '</p>' : '') +
      '<div class="emp-actions" style="padding:6px 0 0;"><button class="btn" type="button" id="hrsSumSave"' + (hrsSum.saved || d.stub ? ' disabled' : '') + '>' + (hrsSum.saved ? 'Saved' : 'Save summary to SharePoint') + '</button><button class="btn ghost" type="button" id="hrsSumCopy">Copy as text</button></div>' +
      '<div class="ai-fb">Was this helpful? <button class="btn ghost" type="button" data-sum-fb="Helpful"' + (hrsSum.feedback ? ' disabled' : '') + '>' + (hrsSum.feedback === 'Helpful' ? 'Helpful (thanks)' : 'Helpful') + '</button> <button class="btn ghost" type="button" data-sum-fb="Not helpful"' + (hrsSum.feedback ? ' disabled' : '') + '>' + (hrsSum.feedback === 'Not helpful' ? 'Not helpful (noted)' : 'Not helpful') + '</button></div>' +
      '<div class="ai-note">AI-generated from the figures on this screen; every number was checked against them. The rule check and the dashboard are the source of truth.' + (hrsSum.note ? ' ' + escapeHtml(hrsSum.note) : '') + (d.model && d.model !== 'stub' ? ' Model: ' + escapeHtml(d.model) + '.' : '') + '</div></div>';
  }
  el.innerHTML = html;
  var b = document.getElementById('hrsSumBtn'); if(b) b.addEventListener('click', hrsSummarize);
  var tg = document.getElementById('hrsSumToggle'); if(tg) tg.addEventListener('click', function(){ hrsSum.collapsed = !hrsSum.collapsed; renderHoursSummary(); });
  var sv2 = document.getElementById('hrsSumSave'); if(sv2) sv2.addEventListener('click', hrsSumSave);
  var cp = document.getElementById('hrsSumCopy'); if(cp) cp.addEventListener('click', function(){
    var txt = hrsSumText(hrsSum.data, hrsSum.cands);
    try{ navigator.clipboard.writeText(txt).then(function(){ cp.textContent = 'Copied'; }, function(){ cp.textContent = 'Copy failed'; }); } catch(e){ cp.textContent = 'Copy failed'; }
  });
  Array.prototype.forEach.call(el.querySelectorAll('[data-sum-fb]'), function(x){ x.addEventListener('click', function(){ hrsSumFeedback(x.getAttribute('data-sum-fb')); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-sum-go]'), function(x){ x.addEventListener('click', function(){ hrsGoToFinding(x.getAttribute('data-sum-go')); }); });
}
function hrsGoToFinding(key){
  hrsFilter = 'all'; hrsStatusFilter = 'all'; hrsListHidden = false;
  var ix = visibleFindings().map(function(i){ return i.key; }).indexOf(key); if(ix >= 0) hrsPage = Math.floor(ix / hrsPageSize) + 1;
  renderHoursFindings();
  var box = null; Array.prototype.forEach.call(document.querySelectorAll('[data-hrs-sel]'), function(c){ if(c.getAttribute('data-hrs-sel') === key) box = c; });
  if(!box) return;
  var row = box.closest('.hrs-row') || box; row.scrollIntoView({ block: 'center', behavior: 'smooth' });
  row.style.outline = '2px solid var(--accent)'; setTimeout(function(){ row.style.outline = ''; }, 2500);
}
async function hrsSummarize(){
  if(!hrsFindings || (hrsSum && hrsSum.state === 'loading')) return;
  var st = hrsWeekStats(); if(!st) return;
  hrsSum = { weekStart: hrsWeek.start, state: 'loading' }; renderHoursSummary();
  try{
    if(typeof upaeAiSummary !== 'function') throw new Error('The AI service is not available in this view.');
    var data = await upaeAiSummary(st.payload);
    if(!hrsWeek || hrsWeek.start !== st.payload.week.start) return;   // the user moved to another week meanwhile
    hrsSum = { weekStart: hrsWeek.start, state: 'done', data: data, cands: st.cands, feedback: '', saved: false, note: '' };
    renderHoursSummary();
    if(typeof upaeCreateAILog === 'function' && !data.stub){
      upaeCreateAILog({ purpose: 'summary', findingKey: '', weekStart: hrsWeek.start, userEmail: (typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : ''), model: data.model, promptVersion: data.promptVersion, confidence: data.confidence, usage: data.usage,
        output: JSON.stringify({ headline: data.headline, keyChanges: data.keyChanges, riskRanked: data.riskRanked, suggestedFollowUps: data.suggestedFollowUps, dataLimits: data.dataLimits }) })
        .then(function(it){ if(hrsSum && hrsSum.state === 'done') hrsSum.logId = it && it.id; })
        .catch(function(err){ console.warn('UPAE: could not write AILog', err); });
    }
  } catch(err){
    console.error('UPAE: summary failed', err);
    if(typeof upaeLogFailedUsage === 'function') upaeLogFailedUsage('summary', err, hrsWeek && hrsWeek.start);
    hrsSum = { weekStart: hrsWeek ? hrsWeek.start : '', state: 'error', err: err.message || String(err) }; renderHoursSummary();
  }
}
async function hrsSumSave(){
  if(!hrsSum || hrsSum.state !== 'done' || hrsSum.saved) return;
  var d = hrsSum.data, btn = document.getElementById('hrsSumSave'); if(btn){ btn.disabled = true; btn.textContent = 'Saving…'; }
  try{
    if(typeof upaeSaveSummary !== 'function') throw new Error('SharePoint is not connected in this view.');
    var who = (typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : '');
    var item = await upaeSaveSummary({ weekStart: hrsWeek.start, weekNo: hrsWeek.weekNo, headline: d.headline, summaryJson: JSON.stringify({ data: { headline: d.headline, keyChanges: d.keyChanges, riskRanked: d.riskRanked, jobFocus: d.jobFocus || [], suggestedFollowUps: d.suggestedFollowUps, dataLimits: d.dataLimits, confidence: d.confidence }, cands: hrsSum.cands }),
      model: d.model, promptVersion: d.promptVersion, confidence: d.confidence, generatedBy: who });
    hrsSum.saved = true; hrsSum.id = item && item.id; hrsSum.note = 'Saved to SharePoint (WeeklySummaries).';
    hrsSums = hrsSums || []; hrsSums.push({ spItemId: hrsSum.id, weekStart: hrsWeek.start, weekNo: hrsWeek.weekNo, headline: d.headline, json: JSON.stringify({ data: d, cands: hrsSum.cands }), model: d.model, promptVersion: d.promptVersion, confidence: d.confidence, generatedBy: who, feedback: '', at: new Date().toISOString() });
  } catch(err){
    console.warn('UPAE: could not save summary', err);
    hrsSum.note = 'Not saved: ' + (err.message || String(err)) + ' (the WeeklySummaries list may not exist yet).';
  }
  renderHoursSummary();
}
async function hrsSumFeedback(value){
  if(!hrsSum || hrsSum.feedback) return;
  hrsSum.feedback = value; renderHoursSummary();
  try{
    if(hrsSum.id && typeof upaeUpdateSummaryFeedback === 'function') await upaeUpdateSummaryFeedback(hrsSum.id, value);
    if(hrsSum.logId && typeof upaeUpdateAILogFeedback === 'function') await upaeUpdateAILogFeedback(hrsSum.logId, value);
  } catch(err){ console.warn('UPAE: could not save summary feedback', err); }
}

// ---------------- save week ----------------
async function hrsSave(){
  if(hrsBusy || !hrsWeek || !hrsSaved || hrsSavedError) return;
  if(typeof upaeBatchWeeklyHours !== 'function'){ alert('SharePoint is not connected in this view.'); return; }
  if(!hrsFindings) hrsFindings = hrsRunCheck();
  var plan = hrsPlanSave(), ops = [];
  plan.fresh.forEach(function(r){ ops.push({ kind: 'create', row: r }); });
  plan.changed.forEach(function(c){ ops.push({ kind: 'update', row: c.next, spItemId: c.cur.spItemId }); });
  hrsBusy = true; document.getElementById('hrsSaveBtn').disabled = true;
  var st = document.getElementById('hrsStatus');
  try{
    var ok = 0, bad = [];
    if(ops.length){
      var res = await upaeBatchWeeklyHours(ops, function(done, total){ st.textContent = 'Saving hours ' + done + ' of ' + total + '…'; });
      res.forEach(function(r, i){
        if(r && r.ok){
          ok++;
          var row = ops[i].row;
          if(ops[i].kind === 'create') hrsSaved.push(Object.assign({}, row, { spItemId: r.spItemId }));
          else { var cur = hrsSaved.filter(function(x){ return x.spItemId === ops[i].spItemId; })[0]; if(cur) Object.assign(cur, row); }
        } else bad.push({ error: r && r.error });
      });
    }
    var fr = { ok: 0, bad: [] };
    if(!bad.length && hrsFindings && hrsFindings.items.length && !hrsFindError){
      // only findings that are new or whose saved copy differs need writing
      var by = {}; (hrsFindSaved || []).forEach(function(r){ by[r.key] = r; });
      var todo = hrsFindings.items.filter(function(i){ var s = by[i.key]; return !s || JSON.stringify(hrsFindingRow(i)) !== JSON.stringify(Object.assign({}, hrsFindingRow(i), { reported: s.reported, expected: s.expected, broken: s.broken, status: s.status, sentTo: s.sentTo, sentDate: s.sentDate, ruleName: s.ruleName, ruleId: s.ruleId, union: s.union, weekNo: s.weekNo })); });
      fr = await hrsPersistFindings(todo, function(d, t){ st.textContent = 'Saving findings ' + d + ' of ' + t + '…'; });
    }
    hrsBusy = false;
    renderHoursPanel(); renderHoursSaved();
    var st2 = document.getElementById('hrsStatus');
    if(bad.length || fr.bad.length){
      st2.className = 'emp-status err';
      st2.textContent = (bad.length ? bad.length + ' hour rows failed (' + (bad[0].error || 'unknown') + '). ' : '') + (fr.bad.length ? fr.bad.length + ' findings not saved (' + (fr.bad[0].error || 'unknown') + '). ' : '') + 'Press Save again to retry only what is missing.';
    } else st2.textContent = 'Saved week ' + hrsWeek.weekNo + ': ' + ok + ' hour row(s) and ' + fr.ok + ' finding(s)' + (hrsFindError ? ' (findings not saved: WeeklyFindings list missing)' : '') + '.';
  } catch(err){
    hrsBusy = false; console.error('UPAE: saving the week failed', err);
    renderHoursPanel();
    var st3 = document.getElementById('hrsStatus'); st3.className = 'emp-status err'; st3.textContent = 'Saving failed: ' + (err.message || err);
  }
}

// ---------------- email the project managers ----------------
window.hrsGetPMs = function(){ return hrsPMs; };
window.hrsOpenWeek = async function(start){ if(!hrsLoadedOnce) await hrsRefreshSaved(); hrsOpenSaved(start); };
var mailOverlay = document.getElementById('hrsMailOverlay');
var mailGroups = [], mailBusy = false, mailSkipped = [];
function pmOptions(selected){
  return '<option value="">— choose a project manager —</option>' + hrsPMs.slice().sort(function(a, b){ return a.name.localeCompare(b.name); })
    .map(function(p){ return '<option value="' + escapeHtml(p.email) + '"' + (selected && selected.toLowerCase() === p.email.toLowerCase() ? ' selected' : '') + '>' + escapeHtml(p.name) + ' &lt;' + escapeHtml(p.email) + '&gt;</option>'; }).join('');
}
function mailTable(items, withAi){
  var th = 'style="text-align:left;padding:5px 8px;border:1px solid #ccc;background:#f0f0f0;font-size:12px;"', td = 'style="padding:5px 8px;border:1px solid #ccc;font-size:12px;vertical-align:top;"';
  return '<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;"><tr><th ' + th + '>Employee</th><th ' + th + '>Day</th><th ' + th + '>Reported</th><th ' + th + '>Rule says</th><th ' + th + '>Rule broken</th></tr>' +
    items.map(function(i){
      return '<tr><td ' + td + '>' + escapeHtml(i.name) + ' (' + escapeHtml(i.id) + ')' + (i.union ? '<br><span style="color:#666">' + escapeHtml(i.union) + '</span>' : '') + '</td><td ' + td + '>' + (i.iso ? hrsPretty(i.iso) : '—') + (i.jobs ? '<br><span style="color:#666">Job: ' + escapeHtml(i.jobs) + '</span>' : '') + '</td><td ' + td + '>' + (i.rep ? hrsFmt(i.rep) : '—') +
        '</td><td ' + td + '>' + (i.exp ? hrsFmt(i.exp) : '—') + '</td><td ' + td + '><b>' + escapeHtml(i.ruleId === 'Data check' ? 'Data check' : i.ruleName) + '</b><br>' + escapeHtml(i.broken.join(' ')) +
        (withAi && i.ai && i.ai.explanation ? '<div style="margin-top:6px;padding:5px 7px;background:#f4f1fb;border-left:3px solid #7a5fc0;color:#333;"><b>AI note:</b> ' + escapeHtml(i.ai.explanation) + (i.ai.suggestedAction ? '<br><b>Suggested check:</b> ' + escapeHtml(i.ai.suggestedAction) : '') + '<br><span style="color:#777;font-size:11px;">AI-generated, ' + escapeHtml(i.ai.confidence || 'low') + ' confidence – please verify.</span></div>' : '') + '</td></tr>';
    }).join('') + '</table>';
}

function pmByEmail(email){ return hrsPMs.filter(function(p){ return p.email.toLowerCase() === String(email || '').toLowerCase(); })[0] || null; }
function empOf(id){ return employees.filter(function(e){ return empKey(e.id) === empKey(id); })[0] || null; }
function mailSubjectBase(){ return 'Payroll review needed – Week ' + (hrsWeek.weekNo || '') + ' (' + fmtMD(hrsWeek.start) + '–' + fmtMD(hrsWeek.end) + ')'; }
function mailIntro(){
  return 'The weekly payroll check flagged the records below for week ' + (hrsWeek.weekNo || '') + '. ' + (hrsWeek.source === 'detail' ? 'The job each day was logged on is shown under the day, so please look at the ones on your jobs and let me know' : 'Job numbers are not on this report, so please look at the ones for employees on your jobs and let me know') + ' what needs to be corrected in B2W.\n\nThank you.';
}
function itemJobBases(i){
  var out = [], seen = {};
  String(i.jobs || '').split('; ').forEach(function(x){ var j = x.replace(/ \(.*$/, '').trim(); if(!j || j === '(no job)') return; var b = hrsJobBase(j); if(!seen[b]){ seen[b] = 1; out.push(b); } });
  return out;
}
function jobTitleOf(b){ var a = hrsJobPMs[b]; return (a && a.jobName) || (hrsWeek && hrsWeek.jobTitles && hrsWeek.jobTitles[b]) || ''; }
function jobLabel(b){ return 'Job ' + b + (jobTitleOf(b) ? ' · ' + jobTitleOf(b) : ''); }
// A record goes to the PM who owns its job. A day on two jobs goes to both PMs. Jobs nobody owns yet get their own group to assign.
function buildMailGroups(items){
  var groups = {}, order = [];
  function put(id, o, i){
    var g = groups[id];
    if(!g){ g = groups[id] = Object.assign({ id: id, items: [], seen: {}, pm: '', suggested: '', jobKey: '', jobs: [] }, o); order.push(id); }
    if(!g.seen[i.key]){ g.seen[i.key] = 1; g.items.push(i); }
    return g;
  }
  items.forEach(function(i){
    var bases = itemJobBases(i), e = empOf(i.id), ePm = e && e.pmEmail ? e.pmEmail.toLowerCase() : '';
    if(!bases.length){ put('emp:' + ePm, { suggested: ePm, pm: ePm && pmByEmail(ePm) ? ePm : '' }, i); return; }
    bases.forEach(function(b){
      var jp = hrsJobPMs[b], pe = jp ? jp.pmEmail.toLowerCase() : '';
      if(pe){ var g = put('pm:' + pe, { suggested: pe, pm: pmByEmail(pe) ? pe : '', fromJob: true }, i); if(g.jobs.indexOf(b) < 0) g.jobs.push(b); }
      else put('job:' + b, { jobKey: b }, i);
    });
  });
  var rank = function(g){ return g.pm ? 0 : g.fromJob ? 1 : g.jobKey ? 2 : 3; };
  return order.map(function(k){ return groups[k]; }).sort(function(a, b){ return rank(a) - rank(b) || a.id.localeCompare(b.id); });
}
function mailPlan(){
  // merge groups that were pointed at the same person
  var byPm = {}, missing = 0, miss = {};
  mailGroups.forEach(function(g){
    if(!g.pm){ g.items.forEach(function(i){ miss[i.key] = 1; }); return; }
    var o = byPm[g.pm.toLowerCase()] = byPm[g.pm.toLowerCase()] || { pm: g.pm, items: [], seen: {}, jobs: [], learn: [] };
    g.items.forEach(function(i){ if(!o.seen[i.key]){ o.seen[i.key] = 1; o.items.push(i); } });
    (g.jobs || []).forEach(function(b){ if(o.jobs.indexOf(b) < 0) o.jobs.push(b); });
    if(g.jobKey){ if(o.jobs.indexOf(g.jobKey) < 0) o.jobs.push(g.jobKey); o.learn.push(g.jobKey); }
  });
  missing = Object.keys(miss).length;
  return { byPm: byPm, emails: Object.keys(byPm).map(function(k){ return byPm[k]; }), missing: missing };
}
function renderMailGroups(){
  var el = document.getElementById('hrsMailGroups');
  el.innerHTML = mailGroups.map(function(g, gi){
    var tag = g.jobKey ? '<span class="mg-tag warn">job has no PM yet</span>' : g.fromJob ? (g.pm ? '<span class="mg-tag">PM of the job</span>' : '<span class="mg-tag warn">PM of the job is not in the Project managers list</span>') : g.suggested ? '<span class="mg-tag">employee\'s usual PM</span>' : '<span class="mg-tag warn">no job or PM known</span>';
    var what = g.jobKey ? escapeHtml(jobLabel(g.jobKey)) : g.fromJob ? 'Jobs: ' + escapeHtml(g.jobs.sort().join(', ')) : '';
    return '<div class="mail-group"><div class="mg-head"><b>' + g.items.length + ' record' + (g.items.length === 1 ? '' : 's') + '</b>' + tag +
      (what ? '<span class="e-sub" style="display:inline;">' + what + '</span>' : '') +
      '<select data-mg-pm="' + gi + '">' + pmOptions(g.pm) + '</select></div>' +
      '<details><summary>Records in this email</summary>' + mailTable(g.items.slice(0, 15), document.getElementById('hrsMailAi').checked) + (g.items.length > 15 ? '<div class="emp-status">&hellip; and ' + (g.items.length - 15) + ' more rows in the email.</div>' : '') + '</details></div>';
  }).join('');
  Array.prototype.forEach.call(el.querySelectorAll('[data-mg-pm]'), function(s){
    s.addEventListener('change', function(){ mailGroups[parseInt(s.getAttribute('data-mg-pm'), 10)].pm = s.value; updateMailSummary(); });
  });
  updateMailSummary();
}
function updateMailSummary(){
  var plan = mailPlan(), uq = {}; plan.emails.forEach(function(e){ e.items.forEach(function(i){ uq[i.key] = 1; }); }); var n = Object.keys(uq).length;
  var copyMe = document.getElementById('hrsMailCopyMe').checked;
  var sum = document.getElementById('hrsMailSummary');
  sum.innerHTML = plan.emails.length
    ? '<b>Will send ' + plan.emails.length + ' email' + (plan.emails.length === 1 ? '' : 's') + '</b> covering ' + n + ' record' + (n === 1 ? '' : 's') + ': ' + plan.emails.map(function(e){ var p = pmByEmail(e.pm); return escapeHtml(p ? p.name : e.pm) + ' (' + e.items.length + ')'; }).join(', ') + (copyMe ? '. A copy goes to you.' : '.')
    : 'Choose a project manager for each group below.';
  if(plan.missing) sum.innerHTML += ' <span style="color:var(--warn);font-weight:600;">' + plan.missing + ' record(s) still need a project manager and will not be sent.</span>';
  document.getElementById('hrsMailSend').textContent = plan.emails.length ? 'Send ' + plan.emails.length + ' email' + (plan.emails.length === 1 ? '' : 's') : 'Send email';
  document.getElementById('hrsMailSend').disabled = mailBusy || !plan.emails.length;
  document.getElementById('hrsMailTest').disabled = mailBusy || !plan.emails.length;
}
function openMailModal(){
  var sel = selectedItems(); if(!sel.length) return;
  // safeguard: findings already sent to a PM are left out unless the user asks for them
  mailSkipped = sel.filter(function(i){ return i.status === STATUS_SENT; });
  var items = sel.filter(function(i){ return i.status !== STATUS_SENT; });
  document.getElementById('hrsMailIncludeSent').checked = false;
  var sk = document.getElementById('hrsMailSkipped');
  sk.hidden = !mailSkipped.length;
  sk.querySelector('span').textContent = mailSkipped.length + ' selected record' + (mailSkipped.length === 1 ? ' was' : 's were') + ' already sent to a project manager (' + mailSkipped.slice(0, 3).map(function(i){ return i.name + ' on ' + (i.sentDate || '?'); }).join('; ') + (mailSkipped.length > 3 ? '…' : '') + ') and ' + (mailSkipped.length === 1 ? 'is' : 'are') + ' left out.';
  mailGroups = buildMailGroups(items);
  document.getElementById('hrsMailCc').value = '';
  document.getElementById('hrsMailCopyMe').checked = false;
  document.getElementById('hrsMailRemember').checked = true;
  var nAi = items.filter(function(i){ return i.ai && i.ai.explanation; }).length;
  document.getElementById('hrsMailAi').checked = nAi > 0; document.getElementById('hrsMailAi').disabled = !nAi;
  document.getElementById('hrsMailAiLbl').textContent = 'Include the AI explanation (' + nAi + ' of ' + items.length + ' records have one)';
  document.getElementById('hrsMailSubject').value = mailSubjectBase();
  document.getElementById('hrsMailNote').value = mailIntro();
  document.getElementById('hrsMailCount').textContent = sel.length + ' record' + (sel.length === 1 ? '' : 's') + ' selected.';
  var s = document.getElementById('hrsMailStatus'); s.textContent = hrsPMs.length ? '' : (hrsPMError ? 'The ProjectManagers list could not be read (' + hrsPMError + ').' : 'No project managers yet — add them under "Project managers" first.'); s.className = 'emp-status' + (hrsPMs.length ? '' : ' err');
  renderMailGroups();
  mailOverlay.hidden = false;
}
function closeMailModal(){ if(mailBusy) return; mailOverlay.hidden = true; }
function reviewLinkHtml(){
  var on = document.getElementById('hrsMailLink');
  if(!on || !on.checked) return '';
  var url = location.origin + '/review.html', f = 'font-family:Segoe UI,Arial,sans-serif;';
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:14px 0 6px;"><tr><td bgcolor="#235870" style="background:#235870;border-radius:6px;padding:10px 18px;"><a href="' + url + '" style="' + f + 'color:#ffffff;font-size:14px;font-weight:600;text-decoration:none;display:inline-block;">Open the review page</a></td></tr></table>' +
    '<p style="' + f + 'color:#666;font-size:12px;margin:6px 0 14px;">Sign in with your company account. Mark each record Correct as reported, Needs correction or Question, and add a note. Payroll sees your answer straight away, so there is no need to reply by email.</p>' +
    '<hr style="border:0;border-top:1px solid #ddd;margin:0 0 4px;">';
}
function mailBody(pmName, e){
  var items = e.items;
  var note = document.getElementById('hrsMailNote').value.split('\n').map(function(l){ return l.trim() ? '<p style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:0 0 8px;">' + escapeHtml(l) + '</p>' : ''; }).join('');
  var first = pmName ? pmName.split(/[ ,]+/)[0] : '';
  var withAi = document.getElementById('hrsMailAi').checked, owned = (e.jobs || []).slice().sort(), body = '', used = {};
  owned.forEach(function(b){
    var its = items.filter(function(i){ return itemJobBases(i).indexOf(b) >= 0; });
    if(!its.length) return;
    its.forEach(function(i){ used[i.key] = 1; });
    body += '<p style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:12px 0 4px;"><b>' + escapeHtml(jobLabel(b)) + '</b> &middot; ' + its.length + ' record' + (its.length === 1 ? '' : 's') + '</p>' + mailTable(its, withAi);
  });
  var rest = items.filter(function(i){ return !used[i.key]; });
  if(rest.length) body += (body ? '<p style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:12px 0 4px;"><b>Other records</b> (no job on the report)</p>' : '') + mailTable(rest, withAi);
  return '<p style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:0 0 8px;">Hi ' + escapeHtml(first) + ',</p>' + note + reviewLinkHtml() + body + '<p style="font-family:Segoe UI,Arial,sans-serif;font-size:11px;color:#777;">Sent from the UPAE Exception Console.</p>';
}
async function sendMails(testOnly){
  if(mailBusy) return;
  var plan = mailPlan(), st = document.getElementById('hrsMailStatus');
  if(!plan.emails.length) return;
  if(typeof upaeSendMail !== 'function'){ st.className = 'emp-status err'; st.textContent = 'Sending is not available in this view.'; return; }
  var me = (typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : '');
  if(testOnly && !me){ st.className = 'emp-status err'; st.textContent = 'Your own email address is not known, so a test copy cannot be sent.'; return; }
  var cc = document.getElementById('hrsMailCc').value.split(/[;,\s]+/).filter(function(x){ return /@/.test(x); });
  if(!testOnly && document.getElementById('hrsMailCopyMe').checked && me) cc.push(me);
  var base = document.getElementById('hrsMailSubject').value || mailSubjectBase();
  mailBusy = true; updateMailSummary();
  st.className = 'emp-status'; var sent = [], failed = [], today = hrsIso(new Date());
  for(var k = 0; k < plan.emails.length; k++){
    var e = plan.emails[k], p = pmByEmail(e.pm), nm = p ? p.name : e.pm;
    st.textContent = (testOnly ? 'Sending test ' : 'Sending ') + (k + 1) + ' of ' + plan.emails.length + ' (' + nm + ')…';
    var subject = (testOnly ? '[TEST – not sent to the PM] ' : '') + base + ': ' + e.items.length + ' record' + (e.items.length === 1 ? '' : 's');
    try{
      await upaeSendMail(testOnly ? [me] : [e.pm], testOnly ? [] : cc, subject, (testOnly ? '<p style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;color:#a33a34;">This is a test copy. It would go to ' + escapeHtml(nm) + ' &lt;' + escapeHtml(e.pm) + '&gt;.</p>' : '') + mailBody(nm, e));
      if(testOnly) { sent.push(nm); break; }   // one sample is enough for a test
      sent.push(nm);
      e.items.forEach(function(i){ i.status = STATUS_SENT; i.sentTo = nm + ' <' + e.pm + '>'; i.sentDate = today; });
      e.ok = true;
    } catch(err){
      console.error('UPAE: send mail failed', err);
      failed.push(nm + ': ' + (err.message || String(err)));
      if(/Mail\.Send|permission|consent|401|403/i.test(String(err.message))) break;   // no point trying the rest
    }
  }
  mailBusy = false;
  if(testOnly){
    st.className = 'emp-status' + (failed.length ? ' err' : ''); st.textContent = failed.length ? 'Test failed: ' + failed[0] : 'Test copy sent to ' + me + '. Nothing was marked as sent.';
    updateMailSummary(); return;
  }
  var okEmails = plan.emails.filter(function(e){ return e.ok; });
  var okItems = [], okSeen = {};
  okEmails.forEach(function(e){ e.items.forEach(function(i){ if(!okSeen[i.key]){ okSeen[i.key] = 1; okItems.push(i); } }); });
  // remember who usually handles each employee
  var remembered = 0, jobsLearned = 0;
  if(document.getElementById('hrsMailRemember').checked){
    okEmails.forEach(function(e){
      (e.learn || []).forEach(function(b){
        var pmo = pmByEmail(e.pm); if(hrsJobPMs[b]) return;
        var rec = { job: b, jobName: jobTitleOf(b), pmEmail: e.pm, pmName: pmo ? pmo.name : '' };
        hrsJobPMs[b] = Object.assign({ spItemId: '' }, rec); jobsLearned++;
        if(typeof upaeCreateJobAssignment === 'function') upaeCreateJobAssignment(rec).then(function(r){ if(r && r.id) hrsJobPMs[b].spItemId = r.id; }).catch(function(err){ console.warn('UPAE: could not remember the job PM', err); });
      });
      if(e.jobs && e.jobs.length) return;   // routed by job: the employee's usual PM is not changed by it
      var seen = {};
      e.items.forEach(function(i){
        var emp = empOf(i.id); if(!emp || seen[empKey(i.id)] || String(emp.pmEmail || '').toLowerCase() === e.pm.toLowerCase()) return;
        seen[empKey(i.id)] = true; emp.pmEmail = e.pm; remembered++;
        if(typeof upaeSetEmployeePM === 'function' && emp.spItemId) upaeSetEmployeePM(emp.spItemId, e.pm).catch(function(err){ console.warn('UPAE: could not remember PM', err); });
      });
    });
  }
  if(okItems.length) hrsSel = {};
  if(!failed.length){ mailOverlay.hidden = true; }
  else { st.className = 'emp-status err'; st.textContent = (okEmails.length ? okEmails.length + ' sent. ' : '') + 'Not sent — ' + failed.join(' | ') + (/Mail\.Send|permission|consent/i.test(failed.join()) ? ' If this is a permission message, an administrator needs to add the Mail.Send permission to the app registration.' : ''); }
  renderHoursFindings(); renderHoursDash();
  var r = okItems.length ? await hrsPersistFindings(okItems, null, { force: true }) : { bad: [] };
  var st2 = document.getElementById('hrsActionStatus');
  if(st2 && okItems.length){ st2.className = 'emp-status' + (r.bad.length ? ' err' : ''); st2.textContent = okEmails.length + ' email' + (okEmails.length === 1 ? '' : 's') + ' sent for ' + okItems.length + ' record(s)' + (jobsLearned ? '; remembered the PM for ' + jobsLearned + ' job(s)' : '') + (remembered ? '; remembered the PM for ' + remembered + ' employee(s)' : '') + '.' + (r.bad.length ? ' Status could not be saved to SharePoint: ' + (r.bad[0].error || 'unknown error') : ' Status saved.'); }
  updateMailSummary();
}
document.getElementById('hrsMailSend').addEventListener('click', function(){ sendMails(false); });
document.getElementById('hrsMailTest').addEventListener('click', function(){ sendMails(true); });
document.getElementById('hrsMailCancel').addEventListener('click', closeMailModal);
document.getElementById('hrsMailClose').addEventListener('click', closeMailModal);
document.getElementById('hrsMailCopyMe').addEventListener('change', updateMailSummary);
document.getElementById('hrsMailAi').addEventListener('change', function(){ renderMailGroups(); });
document.getElementById('hrsMailIncludeSent').addEventListener('change', function(){
  var inc = this.checked, sel = selectedItems().filter(function(i){ return inc || i.status !== STATUS_SENT; });
  var keep = {}; mailGroups.forEach(function(g){ if(g.pm) keep[g.id] = g.pm; });
  mailGroups = buildMailGroups(sel);
  mailGroups.forEach(function(g){ if(keep[g.id]) g.pm = keep[g.id]; });
  renderMailGroups();
});

// ---------------- jobs and their project manager (Project managers page) ----------------
function jobPmName(a){ var p = pmByEmail(a.pmEmail); return p ? p.name : (a.pmName || ''); }
function jobMatches(a){
  var q = empClean(document.getElementById('jobSearch').value).toLowerCase(), f = document.getElementById('jobFilter').value;
  if(f === 'active' && !a.active) return false;
  if(f === 'inactive' && a.active) return false;
  if(f === 'nopm' && a.pmEmail) return false;
  if(!q) return true;
  return (a.job + ' ' + a.jobName + ' ' + jobPmName(a) + ' ' + a.pmEmail).toLowerCase().indexOf(q) >= 0;
}
function renderJobs(){
  var el = document.getElementById('jobList'); if(!el) return;
  document.getElementById('jobCount').textContent = hrsJobRows.length;
  var rows = hrsJobRows.filter(jobMatches).sort(function(a, b){ return a.job.localeCompare(b.job); });
  var pages = Math.max(1, Math.ceil(rows.length / jobPageSize)); if(jobPage > pages) jobPage = pages; if(jobPage < 1) jobPage = 1;
  var html = hrsJobsError ? '<div class="emp-status err">The JobAssignments list could not be read (' + escapeHtml(hrsJobsError) + ').</div>' : '';
  html += '<div class="hrs-row head" style="grid-template-columns:110px 1.4fr 1.4fr 70px 150px;"><div>Job</div><div>Job name</div><div>Project manager</div><div>Active</div><div></div></div>';
  html += rows.slice((jobPage - 1) * jobPageSize, jobPage * jobPageSize).map(function(a){
    var st = 'grid-template-columns:110px 1.4fr 1.4fr 70px 150px;';
    if(jobEdit === a.spItemId)
      return '<div class="hrs-row" style="' + st + '"><div><b>' + escapeHtml(a.job) + '</b></div><div><input type="text" id="jobEditName" value="' + escapeHtml(a.jobName) + '" style="width:100%;padding:5px 7px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink);font:inherit;font-size:12.5px;"></div>' +
        '<div><select id="jobEditPm" style="width:100%;padding:5px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink);font:inherit;font-size:12.5px;">' + pmOptions(a.pmEmail) + '</select></div>' +
        '<div><input type="checkbox" id="jobEditActive"' + (a.active ? ' checked' : '') + '></div><div><button class="btn primary" type="button" data-job-save="' + escapeHtml(a.spItemId) + '">Save</button> <button class="btn ghost" type="button" data-job-cancel="1">Cancel</button></div></div>';
    return '<div class="hrs-row' + (a.pmEmail ? '' : ' warn') + '" style="' + st + '"><div><b>' + escapeHtml(a.job) + '</b></div><div>' + escapeHtml(a.jobName || '—') + '</div>' +
      '<div>' + (a.pmEmail ? escapeHtml(jobPmName(a) || a.pmEmail) + '<span class="e-sub">' + escapeHtml(a.pmEmail) + (pmByEmail(a.pmEmail) ? '' : ' · not in the Project managers list') + '</span>' : '<i>No manager yet</i>') + '</div>' +
      '<div>' + (a.active ? 'Yes' : 'No') + '</div><div><button class="btn ghost" type="button" data-job-edit="' + escapeHtml(a.spItemId) + '">Edit</button></div></div>';
  }).join('') || '<div class="emp-status">No jobs match.</div>';
  html += rows.length > 10 ? pagerHtml('job', rows.length, jobPage, jobPageSize) : '';
  el.innerHTML = html;
  pagerWire(el, 'job', function(pg, size){ if(size){ jobPageSize = size; jobPage = 1; } else jobPage = pg; renderJobs(); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-job-edit]'), function(b){ b.addEventListener('click', function(){ jobEdit = b.getAttribute('data-job-edit'); renderJobs(); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-job-cancel]'), function(b){ b.addEventListener('click', function(){ jobEdit = null; renderJobs(); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-job-save]'), function(b){ b.addEventListener('click', function(){ saveJobEdit(b.getAttribute('data-job-save')); }); });
  document.getElementById('jobAddPm').innerHTML = pmOptions(document.getElementById('jobAddPm').value);
  // jobs in the loaded week that nobody owns yet
  var hint = document.getElementById('jobWeekHint'), miss = [];
  if(hrsWeek && hrsWeek.jobTitles){
    var have = {}; hrsJobRows.forEach(function(a){ have[hrsJobBase(a.job)] = 1; });
    miss = Object.keys(hrsWeek.jobTitles).filter(function(b){ return !have[b]; }).sort();
  }
  hint.innerHTML = miss.length ? '<div class="emp-status">' + miss.length + ' job' + (miss.length === 1 ? '' : 's') + ' in the loaded week ' + (miss.length === 1 ? 'has' : 'have') + ' no manager yet. Click one to fill the Add form: ' +
    miss.map(function(b){ return '<button class="chip-filter" type="button" data-job-fill="' + escapeHtml(b) + '">' + escapeHtml(b) + '</button>'; }).join(' ') + '</div>' : '';
  Array.prototype.forEach.call(hint.querySelectorAll('[data-job-fill]'), function(b){ b.addEventListener('click', function(){
    var k = b.getAttribute('data-job-fill'); document.getElementById('jobAddNo').value = k; document.getElementById('jobAddName').value = hrsWeek.jobTitles[k] || ''; document.getElementById('jobAddPm').focus();
  }); });
}
async function saveJobEdit(id){
  var a = hrsJobRows.filter(function(x){ return x.spItemId === id; })[0], st = document.getElementById('jobStatus'); if(!a) return;
  var email = document.getElementById('jobEditPm').value, nm = empClean(document.getElementById('jobEditName').value), act = document.getElementById('jobEditActive').checked;
  var p = pmByEmail(email), rec = { jobName: nm, pmEmail: email, pmName: p ? p.name : (email ? a.pmName : ''), active: act };
  try{
    await upaeUpdateJobAssignment(id, rec);
    a.jobName = rec.jobName; a.pmEmail = rec.pmEmail; a.pmName = rec.pmName; a.active = act; jobEdit = null;
    delete hrsJobPMs[hrsJobBase(a.job)]; hrsJobRows.forEach(function(x){ if(x.active && x.job && x.pmEmail) hrsJobPMs[hrsJobBase(x.job)] = x; });
    st.className = 'emp-status'; st.textContent = 'Saved ' + a.job + '.'; renderJobs();
  } catch(err){ st.className = 'emp-status err'; st.textContent = 'Could not save: ' + (err.message || err); }
}
async function addJob(){
  var st = document.getElementById('jobStatus'), no = hrsJobBase(document.getElementById('jobAddNo').value), nm = empClean(document.getElementById('jobAddName').value), email = document.getElementById('jobAddPm').value;
  if(!no){ st.className = 'emp-status err'; st.textContent = 'Enter a job number.'; return; }
  if(hrsJobRows.some(function(a){ return hrsJobBase(a.job) === no; })){ st.className = 'emp-status err'; st.textContent = 'Job ' + no + ' is already on the list. Edit it above instead.'; return; }
  var p = pmByEmail(email), rec = { job: no, jobName: nm, pmEmail: email, pmName: p ? p.name : '' };
  try{
    var r = await upaeCreateJobAssignment(rec);
    var row = Object.assign({ spItemId: r && r.id, active: true }, rec); hrsJobRows.push(row); if(email) hrsJobPMs[no] = row;
    document.getElementById('jobAddNo').value = ''; document.getElementById('jobAddName').value = '';
    st.className = 'emp-status'; st.textContent = 'Added job ' + no + '.'; renderJobs();
  } catch(err){ st.className = 'emp-status err'; st.textContent = 'Could not add: ' + (err.message || err); }
}
['jobSearch', 'jobFilter'].forEach(function(id){ var e = document.getElementById(id); if(e) e.addEventListener(id === 'jobSearch' ? 'input' : 'change', function(){ jobPage = 1; renderJobs(); }); });
if(document.getElementById('jobAddBtn')) document.getElementById('jobAddBtn').addEventListener('click', addJob);

// ---------------- project managers ----------------
function renderPMs(){
  var el = document.getElementById('hrsPmList'); if(!el) return;
  document.getElementById('hrsPmCount').textContent = hrsPMs.length;
  var html = hrsPMError ? '<div class="emp-status err">The ProjectManagers list could not be read (' + escapeHtml(hrsPMError) + ').</div>' : '';
  html += hrsPMs.slice().sort(function(a, b){ return a.name.localeCompare(b.name); }).map(function(p){
    if(hrsPmEdit === p.spItemId)
      return '<div class="pm-row"><span class="pm-edit"><input type="text" id="pmEditName" value="' + escapeHtml(p.name) + '" aria-label="Name"><input type="text" id="pmEditEmail" value="' + escapeHtml(p.email) + '" aria-label="Email"></span>' +
        '<span><button class="btn primary" type="button" data-pm-save="' + escapeHtml(p.spItemId) + '">Save</button> <button class="btn ghost" type="button" data-pm-cancel="1">Cancel</button></span></div>';
    return '<div class="pm-row"><span><b>' + escapeHtml(p.name) + '</b> &middot; ' + escapeHtml(p.email) + '</span><span><button class="btn ghost" type="button" data-pm-edit="' + escapeHtml(p.spItemId) + '">Edit</button> <button class="btn ghost" type="button" data-pm-del="' + escapeHtml(p.spItemId) + '">Remove</button></span></div>';
  }).join('') || '<div class="emp-status">No project managers yet.</div>';
  el.innerHTML = html;
  Array.prototype.forEach.call(el.querySelectorAll('[data-pm-edit]'), function(b){ b.addEventListener('click', function(){ hrsPmEdit = b.getAttribute('data-pm-edit'); renderPMs(); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-pm-cancel]'), function(b){ b.addEventListener('click', function(){ hrsPmEdit = null; renderPMs(); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-pm-save]'), function(b){
    b.addEventListener('click', async function(){
      var id = b.getAttribute('data-pm-save'), pm = hrsPMs.filter(function(p){ return p.spItemId === id; })[0];
      var nm = empClean(document.getElementById('pmEditName').value), em = empClean(document.getElementById('pmEditEmail').value);
      if(!pm || !nm || !/^[^\s@]+@[^\s@]+$/.test(em)){ alert('Enter a name and a valid email address.'); return; }
      try{ await upaeUpdateProjectManager(id, { name: nm, email: em }); pm.name = nm; pm.email = em; hrsPmEdit = null; renderPMs(); }
      catch(err){ alert('Could not save: ' + (err.message || err)); }
    });
  });
  Array.prototype.forEach.call(el.querySelectorAll('[data-pm-del]'), function(b){
    b.addEventListener('click', async function(){
      var id = b.getAttribute('data-pm-del'), pm = hrsPMs.filter(function(p){ return p.spItemId === id; })[0];
      if(!pm || !confirm('Remove ' + pm.name + ' from the project manager list?')) return;
      try{ await upaeDeleteProjectManager(id); hrsPMs = hrsPMs.filter(function(p){ return p.spItemId !== id; }); renderPMs(); }
      catch(err){ alert('Could not remove: ' + (err.message || err)); }
    });
  });
}
async function addPMs(){
  var ta = document.getElementById('hrsPmPaste'), st = document.getElementById('hrsPmStatus');
  var lines = ta.value.split(/\r?\n/).map(function(l){ return l.trim(); }).filter(Boolean), add = [], bad = [];
  lines.forEach(function(l){
    var m = l.match(/^(.*?)[,;\t]\s*([^\s,;]+@[^\s,;]+)\s*$/) || l.match(/^(.*?)\s*<([^>]+@[^>]+)>\s*$/);
    if(!m){ bad.push(l); return; }
    var email = m[2].trim(), name = empClean(m[1]) || email;
    if(hrsPMs.some(function(p){ return p.email.toLowerCase() === email.toLowerCase(); }) || add.some(function(p){ return p.email.toLowerCase() === email.toLowerCase(); })) return;
    add.push({ name: name, email: email });
  });
  if(typeof upaeCreateProjectManager !== 'function'){ st.className = 'emp-status err'; st.textContent = 'SharePoint is not connected in this view.'; return; }
  if(hrsPMError){ st.className = 'emp-status err'; st.textContent = 'The ProjectManagers list could not be read, so adding is switched off.'; return; }
  st.className = 'emp-status'; st.textContent = 'Adding…';
  var ok = 0, fail = 0;
  for(var k = 0; k < add.length; k++){
    try{ var it = await upaeCreateProjectManager(add[k]); hrsPMs.push({ spItemId: it && it.id, name: add[k].name, email: add[k].email }); ok++; }
    catch(err){ fail++; console.error(err); }
  }
  ta.value = bad.join('\n');
  renderPMs();
  st.className = 'emp-status' + (fail || bad.length ? ' err' : '');
  st.textContent = ok + ' added' + (fail ? ', ' + fail + ' failed' : '') + (bad.length ? '. ' + bad.length + ' line(s) were not understood and are left in the box (use: Name, email).' : '.');
}
document.getElementById('hrsPmAdd').addEventListener('click', addPMs);

var hrsFileInput = document.getElementById('hrsFileInput');
if(hrsFileInput){
  hrsFileInput.addEventListener('change', async function(){
    var file = hrsFileInput.files && hrsFileInput.files[0];
    if(!file) return;
    document.getElementById('hrsFileName').textContent = file.name;
    var el = document.getElementById('hoursWork');
    el.innerHTML = '<div class="emp-status">Reading ' + escapeHtml(file.name) + '…</div>';
    hrsFindings = null; hrsFilter = 'all'; hrsStatusFilter = 'all'; hrsSel = {};
    try{
      var buf = await file.arrayBuffer();
      hrsFileInput.value = '';
      var head = new Uint8Array(buf.slice(0, 2)), rows;
      if(head[0] === 0x50 && head[1] === 0x4B) rows = await empParseXlsx(buf);
      else if(/\.xls$/i.test(file.name)) throw new Error('Old .xls files are not supported. Please save it as .xlsx and try again.');
      else rows = empParseCsv(new TextDecoder().decode(buf));
      hrsWeek = hrsIsDetail(rows) ? hrsParseDetail(rows, file.name) : hrsParseSummary(rows, file.name);
      if(!hrsWeek.records.length) throw new Error('No employee rows were found under the header.');
      el.innerHTML = '<div class="emp-status">Checking what is already saved in SharePoint…</div>';
      await hrsRefreshSaved();
    } catch(err){
      console.error('UPAE: could not read the weekly summary', err);
      hrsWeek = null;
      el.innerHTML = '<div class="emp-status err">' + escapeHtml(err.message || String(err)) + '</div>';
      renderHoursKpis();
    }
  });
}
// load saved weeks / findings / project managers the first time the screen is opened
Array.prototype.forEach.call(document.querySelectorAll('a[data-view="hours"]'), function(a){
  a.addEventListener('click', function(){ if(!hrsLoadedOnce) hrsRefreshSaved().then(renderPMs); });
});
Array.prototype.forEach.call(document.querySelectorAll('a[data-view="pms"]'), function(a){
  a.addEventListener('click', function(){ (hrsLoadedOnce ? Promise.resolve() : hrsRefreshSaved()).then(function(){ renderPMs(); renderJobs(); }); });
});
function upaeUpdateTopbar(){
  var v = document.getElementById('topWeekVal'); if(!v) return;
  var wk = hrsWeek, txt = 'none loaded';
  if(wk) txt = 'Week ' + (wk.weekNo || '?') + ' \u00b7 ' + fmtMD(wk.start) + '\u2013' + fmtMD(wk.end);
  else if(typeof hrsSavedWeeks === 'function' && hrsSaved != null){ var ws = hrsSavedWeeks(); if(ws.length) txt = 'Latest saved: week ' + (ws[0].weekNo || '?') + ' \u00b7 ' + fmtMD(ws[0].start) + '\u2013' + fmtMD(hrsAddDays(ws[0].start, 6)); }
  v.textContent = txt;
}
function upaeGoWeekly(run){
  var nav = document.querySelector('.rail nav a[data-view="hours"]'); if(nav) nav.click();
  if(!run) return;
  setTimeout(function(){
    var b = document.getElementById('hrsCheckBtn');
    if(b) b.click(); else { var f = document.getElementById('hrsFileInput'); if(f) f.click(); }
  }, 50);
}
(function(){
  var a = document.getElementById('topWeek'), b = document.getElementById('runWeeklyBtn');
  if(a) a.addEventListener('click', function(){ upaeGoWeekly(false); });
  if(b) b.addEventListener('click', function(){ upaeGoWeekly(true); });
  ['renderHoursPanel','renderHoursSaved'].forEach(function(n){
    var orig = window[n]; if(typeof orig !== 'function') return;
    window[n] = function(){ var r = orig.apply(this, arguments); upaeUpdateTopbar(); return r; };
  });
})();
renderHoursKpis(); renderHoursSaved(); renderHoursDash(); renderPMs(); renderJobs(); upaeUpdateTopbar();
