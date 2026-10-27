// Phase 3 pattern detection, statistical pass. Pure functions: saved WeeklyHours rows in, outliers out. No AI and no randomness,
// so the same saved data always gives the same outliers. The AI only interprets what this finds.
var PAT_DEFAULTS = { repeatWeeks: 3, spikeRatio: 1.5, spikeHours: 10, unionPts: 5, longHours: 12, minUnionHours: 200 };
var PAT_WD = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
var PAT_TYPES = { repeat_dt_day: 'Repeat double time on one weekday', long_day: 'Repeated very long days', spike: 'Hours well above own normal', union_premium: 'Union overtime share jumped' };

function patWd(iso){ var p = String(iso).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]).getDay(); }
function patR1(x){ return Math.round(x * 10) / 10; }
function patMedian(a){ var s = a.slice().sort(function(x, y){ return x - y; }), n = s.length; if(!n) return 0; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; }
function patCfg(c){ var o = {}; Object.keys(PAT_DEFAULTS).forEach(function(k){ var v = c && c[k] != null ? Number(c[k]) : NaN; o[k] = isFinite(v) && v > 0 ? v : PAT_DEFAULTS[k]; }); o.repeatWeeks = Math.max(2, Math.round(o.repeatWeeks)); return o; }

// rows: [{employeeId, employeeName, workDate, weekStart, weekNo, rt, ot, dt}]; unionOf(id) -> union name or ''
function patDetect(rows, unionOf, weekStart, cfgIn){
  var cfg = patCfg(cfgIn), all = {}, emps = {}, unions = {};
  (rows || []).forEach(function(r){ if(r.weekStart && r.weekStart <= weekStart) all[r.weekStart] = true; });
  var weeks = Object.keys(all).sort(), prior = weeks.filter(function(w){ return w < weekStart; });
  var res = { weekStart: weekStart, weeksOfHistory: prior.length, priorWeeks: prior.slice(-12), outliers: [], notes: [], cfg: cfg };
  if(weeks.indexOf(weekStart) < 0){ res.notes.push('That week is not among the saved weeks.'); return res; }
  (rows || []).forEach(function(r){
    if(!r.weekStart || r.weekStart > weekStart) return;
    var id = String(r.employeeId || '').trim().toUpperCase(); if(!id) return;
    var e = emps[id] = emps[id] || { id: id, name: r.employeeName || id, weeks: {} };
    var w = e.weeks[r.weekStart] = e.weeks[r.weekStart] || { days: {}, total: 0 };
    var d = w.days[r.workDate] = w.days[r.workDate] || { rt: 0, ot: 0, dt: 0, wd: patWd(r.workDate) };
    d.rt += r.rt || 0; d.ot += r.ot || 0; d.dt += r.dt || 0; w.total += (r.rt || 0) + (r.ot || 0) + (r.dt || 0);
    var un = unionOf ? unionOf(id) : ''; if(un){
      var u = unions[un] = unions[un] || { name: un, weeks: {} }, uw = u.weeks[r.weekStart] = u.weeks[r.weekStart] || { total: 0, prem: 0 };
      uw.total += (r.rt || 0) + (r.ot || 0) + (r.dt || 0); uw.prem += (r.ot || 0) + (r.dt || 0);
    }
  });
  var out = res.outliers;
  function label(e){ return e.name + ' (' + e.id + ')'; }
  Object.keys(emps).forEach(function(id){
    var e = emps[id], tw = e.weeks[weekStart]; if(!tw) return;
    var worked = Object.keys(e.weeks).filter(function(w){ return e.weeks[w].total > 0; }), un = unionOf ? unionOf(id) : '';
    // 1. double time that keeps landing on the same weekday
    for(var wd = 0; wd < 7; wd++){
      var wk = [], hrs = 0, inTarget = false;
      worked.forEach(function(w){ var h = 0; Object.keys(e.weeks[w].days).forEach(function(iso){ var d = e.weeks[w].days[iso]; if(d.wd === wd) h += d.dt; }); if(h > 0){ wk.push(w); hrs += h; if(w === weekStart) inTarget = true; } });
      if(inTarget && wk.length >= cfg.repeatWeeks){
        out.push({ key: 'pat|repeat_dt_day|' + id + '|' + PAT_WD[wd].slice(0, 3) + '|' + weekStart, type: 'repeat_dt_day', subject: id, employee: label(e), union: un, weekStart: weekStart,
          score: wk.length * 10 + hrs, facts: ['Double time on ' + PAT_WD[wd] + ' in ' + wk.length + ' of ' + worked.length + ' saved weeks worked (' + patR1(hrs) + ' hours in total).'],
          metrics: { weekday: PAT_WD[wd], weeksWithPattern: wk.length, weeksWorked: worked.length, totalDtHours: patR1(hrs) } });
      }
    }
    // 2. very long days that repeat
    var lw = 0, longest = 0, lt = false;
    worked.forEach(function(w){ var mx = 0; Object.keys(e.weeks[w].days).forEach(function(iso){ var d = e.weeks[w].days[iso], t = d.rt + d.ot + d.dt; if(t > mx) mx = t; }); if(mx >= cfg.longHours){ lw++; longest = Math.max(longest, mx); if(w === weekStart) lt = true; } });
    if(lt && lw >= cfg.repeatWeeks){
      out.push({ key: 'pat|long_day|' + id + '|' + weekStart, type: 'long_day', subject: id, employee: label(e), union: un, weekStart: weekStart, score: lw * 10 + longest,
        facts: ['A day of ' + cfg.longHours + ' hours or more in ' + lw + ' of ' + worked.length + ' saved weeks worked (longest ' + patR1(longest) + ' hours).'],
        metrics: { longDayHours: cfg.longHours, weeksWithPattern: lw, weeksWorked: worked.length, longestDay: patR1(longest) } });
    }
    // 3. this week far above the employee's own normal
    var pw = prior.filter(function(w){ return e.weeks[w] && e.weeks[w].total > 0; });
    if(pw.length >= 3 && tw.total > 0){
      var med = patMedian(pw.map(function(w){ return e.weeks[w].total; })), delta = tw.total - med;
      if(med > 0 && tw.total / med >= cfg.spikeRatio && delta >= cfg.spikeHours){
        out.push({ key: 'pat|spike|' + id + '|' + weekStart, type: 'spike', subject: id, employee: label(e), union: un, weekStart: weekStart, score: delta,
          facts: ['Worked ' + patR1(tw.total) + ' hours this week, against a typical ' + patR1(med) + ' hours (median of ' + pw.length + ' earlier weeks), ' + patR1(delta) + ' hours more.'],
          metrics: { thisWeekHours: patR1(tw.total), typicalHours: patR1(med), earlierWeeks: pw.length, extraHours: patR1(delta) } });
      }
    }
  });
  // 4. a union whose overtime and double-time share rose sharply
  Object.keys(unions).forEach(function(n){
    var u = unions[n], t = u.weeks[weekStart]; if(!t || t.total < cfg.minUnionHours) return;
    var pw = prior.filter(function(w){ return u.weeks[w] && u.weeks[w].total >= cfg.minUnionHours; });
    if(pw.length < 3) return;
    var med = patMedian(pw.map(function(w){ return u.weeks[w].prem / u.weeks[w].total * 100; })), cur = t.prem / t.total * 100, pts = cur - med;
    if(pts >= cfg.unionPts && (med === 0 || cur / med >= 1.25)){
      out.push({ key: 'pat|union_premium|' + n + '|' + weekStart, type: 'union_premium', subject: n, employee: '', union: n, weekStart: weekStart, score: pts * 3,
        facts: ['Overtime and double time were ' + patR1(cur) + ' percent of ' + n + ' hours this week, against a typical ' + patR1(med) + ' percent (median of ' + pw.length + ' earlier weeks), up ' + patR1(pts) + ' points.'],
        metrics: { thisWeekSharePct: patR1(cur), typicalSharePct: patR1(med), earlierWeeks: pw.length, pointsUp: patR1(pts), unionHours: patR1(t.total) } });
    }
  });
  out.sort(function(a, b){ return b.score - a.score || (a.key < b.key ? -1 : 1); });
  if(prior.length < 3) res.notes.push('Only ' + prior.length + ' earlier saved week' + (prior.length === 1 ? '' : 's') + ': the "hours above normal" and "union overtime share" checks need at least 3 earlier weeks, so they were skipped.');
  if(weeks.length < cfg.repeatWeeks) res.notes.push('The repeat checks need at least ' + cfg.repeatWeeks + ' saved weeks including this one; ' + weeks.length + ' available.');
  return res;
}
