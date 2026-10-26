(function(){
  var exceptions = [], liveFindings = [], liveHistory = null, liveHistError = null, liveError = null, liveLoaded = false, liveStamp = 0, liveStale = false, liveBusy = false;

  var sevRank = {crit:0, warn:1, good:2};
  var queueEl = document.getElementById('queue');
  var detailBody = document.getElementById('detailBody');
  var detailSub = document.getElementById('detailSub');
  var selectedId = null;
  var activeFilter = 'all';

  function statusLabel(s){
    return s === 'open' ? 'Open' : s === 'review' ? 'Sent to PM' : 'Resolved';
  }
  var KIND_LABEL = { mismatch: 'Hours differ from rule', notlisted: 'Not in Employees list', nolabor: 'No default labor type', excessive: 'Very long day', negative: 'Negative hours' };
  function dEsc(x){ return String(x == null ? '' : x).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function dDay(iso){ if(!iso) return ''; var p = String(iso).split('-'); return p.length === 3 ? (+p[1]) + '/' + (+p[2]) : iso; }
  function buildExceptions(list){
    return list.map(function(f){
      var st = f.status === 'Resolved' ? 'resolved' : f.status === 'Sent to PM' ? 'review' : 'open';
      var sev = st === 'resolved' ? 'good' : (f.kind === 'mismatch' ? 'crit' : 'warn');
      var pmName = f.sentTo ? String(f.sentTo).replace(/\s*<.*>$/, '') : '';
      return { id: 'W' + (f.weekNo || '?') + ' · ' + f.id + ' · ' + dDay(f.iso), key: f.key, severity: sev, source: f.aiExplanation ? 'ai' : 'rule', status: st,
        title: (KIND_LABEL[f.kind] || f.kind) + ' — ' + (f.name || f.id),
        meta: f.id + ' · ' + (f.union || 'no union') + ' · week ' + (f.weekNo || '?'),
        reviewer: pmName || 'Not sent yet', f: f };
    });
  }
  function matchesFilter(ex){
    switch(activeFilter){
      case 'crit': return ex.severity === 'crit';
      case 'needs-review': return ex.status !== 'resolved';
      case 'rule': return ex.status === 'review';
      case 'ai': return ex.source === 'ai';
      case 'pm': return !!ex.f.pmResponse && ex.status !== 'resolved';
      default: return true;
    }
  }
  var SHOW_MAX = 100;
  function renderChips(){
    var c = { all: exceptions.length, crit: 0, nr: 0, sent: 0, ai: 0, pm: 0 };
    exceptions.forEach(function(e){ if(e.severity === 'crit') c.crit++; if(e.status !== 'resolved') c.nr++; if(e.status === 'review') c.sent++; if(e.source === 'ai') c.ai++; if(e.f.pmResponse && e.status !== 'resolved') c.pm++; });
    var lab = { all: 'All · ' + c.all, crit: 'Critical · ' + c.crit, 'needs-review': 'Not resolved · ' + c.nr, rule: 'Sent to PM · ' + c.sent, ai: 'Has AI note · ' + c.ai, pm: 'PM answered · ' + c.pm };
    Array.prototype.forEach.call(document.querySelectorAll('#filters .chip-filter'), function(ch){ ch.textContent = lab[ch.getAttribute('data-filter')]; });
  }
  function renderQueue(){
    queueEl.innerHTML = '';
    renderChips();
    if(!liveLoaded){ queueEl.innerHTML = '<div class="empty-note">' + (liveError ? 'Could not read the WeeklyFindings list: ' + dEsc(liveError) : 'Sign in to load the live exception queue from SharePoint.') + '</div>'; return; }
    var visible = exceptions.filter(matchesFilter);
    if(visible.length === 0){
      queueEl.innerHTML = '<div class="empty-note">' + (exceptions.length ? 'No exceptions match this filter.' : 'No findings saved yet. Load a week on the Weekly Hours screen, run the check and save it; the findings show up here.') + '</div>';
      return;
    }
    var rank = { open: 0, review: 1, resolved: 2 };
    visible = visible.slice().sort(function(a, b){ return (rank[a.status] - rank[b.status]) || (sevRank[a.severity] - sevRank[b.severity]) || String(b.f.weekStart).localeCompare(String(a.f.weekStart)) || String(a.f.name).localeCompare(String(b.f.name)); });
    visible.slice(0, SHOW_MAX).forEach(function(ex){
      var row = document.createElement('div');
      row.className = 'qrow' + (ex.key === selectedId ? ' selected' : '');
      row.setAttribute('role','button');
      row.setAttribute('tabindex','0');
      row.innerHTML =
        '<span class="sev-stripe ' + ex.severity + '" aria-hidden="true"></span>' +
        '<span class="q-main">' +
          '<span class="q-title">' + dEsc(ex.title) + '</span>' +
          '<span class="q-meta"><span class="mono">' + dEsc(ex.id) + '</span> · ' + dEsc(ex.meta) + '</span>' +
        '</span>' +
        '<span class="src-tag ' + ex.source + '">' + (ex.source === 'ai' ? 'AI note' : 'Rule') + '</span>' +
        '<span class="q-reviewer">' + dEsc(ex.reviewer) + '</span>' +
        '<span class="status-pill ' + ex.status + '">' + statusLabel(ex.status) + '</span>';
      var pick = function(){ selectedId = ex.key; renderQueue(); renderDetail(); };
      row.addEventListener('click', pick);
      row.addEventListener('keydown', function(e){ if(e.key === 'Enter' || e.key === ' '){ e.preventDefault(); pick(); }});
      queueEl.appendChild(row);
    });
    if(visible.length > SHOW_MAX){ var more = document.createElement('div'); more.className = 'empty-note'; more.textContent = 'Showing the first ' + SHOW_MAX + ' of ' + visible.length + '. Use the filters to narrow the list.'; queueEl.appendChild(more); }
  }

  function renderDetail(){
    var ex = exceptions.filter(function(e){ return e.key === selectedId; })[0];
    if(!ex){ detailBody.innerHTML = '<div class="empty-note">Select an exception from the queue.</div>'; detailSub.textContent = ''; return; }
    var f = ex.f;
    detailSub.textContent = ex.id + ' · ' + statusLabel(ex.status);
    function field(label, value){ return '<div class="detail-field"><div class="f-label">' + label + '</div><div class="f-value">' + dEsc(value || '—') + '</div></div>'; }
    var fieldsHtml = field('Employee', (f.name || '') + ' (' + f.id + ')') + field('Union', f.union) + field('Day', f.iso) + field('Week', f.weekNo ? 'Week ' + f.weekNo + ' (starts ' + f.weekStart + ')' : f.weekStart) +
      field('Rule triggered', f.ruleId === 'Data check' ? 'Data check' : (f.ruleId + ' · ' + f.ruleName)) + field('Reported', f.reported) + field('Rule says', f.expected) + field('Sent to', f.sentTo ? f.sentTo + (f.sentDate ? ' on ' + f.sentDate : '') : '') + field('Jobs', f.jobs) + field('PM answer', f.pmResponse ? f.pmResponse + (f.pmNote ? ' — ' + f.pmNote : '') + (f.pmBy ? ' (' + f.pmBy + (f.pmDate ? ', ' + f.pmDate : '') + ')' : '') : '');
    var noteBlock = '<div class="rationale" style="background:var(--accent-soft); border-color:transparent; color:var(--accent-soft-ink);"><div class="r-head" style="color:var(--accent-soft-ink);">Rule broken</div>' + dEsc(f.broken ? String(f.broken).split(' | ').join(' ') : '—') + '</div>';
    var aiBlock = '';
    if(f.aiExplanation){
      aiBlock = '<div class="rationale"><div class="r-head">AI explanation · ' + dEsc(f.aiConfidence || 'low') + ' confidence</div>' + dEsc(f.aiExplanation) +
        (f.aiSuggestedAction ? '<br><b>Suggested check:</b> ' + dEsc(f.aiSuggestedAction) : '') + '</div>';
    }
    var resBlock = '';
    if(ex.status === 'resolved'){
      resBlock = '<div class="rationale" style="background:#eaf5ee; border-color:transparent;"><div class="r-head">Resolution</div>' + dEsc(f.resReason || 'Resolved') + (f.note ? ' — ' + dEsc(f.note) : '') + '<br><span style="color:var(--ink-soft)">' + dEsc(f.resolvedBy || '') + (f.resolvedDate ? ' on ' + dEsc(f.resolvedDate) : '') + '</span></div>';
    } else if(f.note){
      resBlock = '<div class="rationale"><div class="r-head">Note</div>' + dEsc(f.note) + '</div>';
    }
    var hist = '';
    if(liveHistory){
      var h = liveHistory.filter(function(x){ return x.key === f.key; }).sort(function(a, b){ return String(b.at).localeCompare(String(a.at)); });
      hist = '<div class="rationale"><div class="r-head">History</div>' + (h.length ? h.map(function(x){
        return '<div style="margin-bottom:4px;">' + dEsc(x.at ? new Date(x.at).toLocaleString() : '') + ' — <b>' + dEsc(x.from || '—') + ' → ' + dEsc(x.to) + '</b> by ' + dEsc(x.by || 'unknown') + (x.detail ? '<br><span style="color:var(--ink-soft)">' + dEsc(x.detail) + '</span>' : '') + '</div>';
      }).join('') : 'No status changes recorded yet.') + '</div>';
    } else if(liveHistError){
      hist = '<div class="rationale"><div class="r-head">History</div>The FindingHistory list could not be read (' + dEsc(liveHistError) + '). Create it to start recording who changed what.</div>';
    }
    var stamp = f.modified ? '<div class="emp-status" style="padding:8px 0 0;">Last updated ' + dEsc(new Date(f.modified).toLocaleString()) + (f.modifiedBy ? ' by ' + dEsc(f.modifiedBy) : '') + '</div>' : '';
    detailBody.innerHTML =
      '<div class="detail-title">' + dEsc(ex.title) + '</div>' +
      '<div class="detail-grid">' + fieldsHtml + '</div>' + noteBlock + aiBlock + resBlock + hist + stamp +
      '<div class="actions"><button class="btn primary" type="button" id="dashOpenWeek">Open this week in Weekly Hours</button></div>';
    document.getElementById('dashOpenWeek').addEventListener('click', async function(){
      var nav = document.querySelector('.rail nav a[data-view="hours"]'); if(nav) nav.click();
      if(typeof window.hrsOpenWeek === 'function'){ try{ await window.hrsOpenWeek(f.weekStart); }catch(err){ console.warn('UPAE: open week', err); } }
    });
  }

  // ---- live KPIs, routing and trend ----
  function medianDays(list){
    var d = [];
    list.forEach(function(f){ if(f.status === 'Resolved' && f.sentDate && f.resolvedDate){ var x = (Date.parse(f.resolvedDate) - Date.parse(f.sentDate)) / 86400000; if(isFinite(x) && x >= 0) d.push(x); } });
    if(!d.length) return null; d.sort(function(a, b){ return a - b; }); var m = Math.floor(d.length / 2);
    return d.length % 2 ? d[m] : (d[m - 1] + d[m]) / 2;
  }
  function renderLiveKpis(){
    var el = document.getElementById('dashKpis'); if(!el) return;
    function kpi(label, val, sub, cls){ return '<div class="kpi"><div class="k-label">' + label + '</div><div class="k-value">' + val + '</div><div class="k-delta ' + (cls || 'flat') + '">' + sub + '</div></div>'; }
    if(!liveLoaded){ el.innerHTML = kpi('Open exceptions', '—', 'sign in to load') + kpi('Resolved', '—', '') + kpi('With an AI note', '—', '') + kpi('Median days to resolve', '—', ''); return; }
    var total = exceptions.length, res = exceptions.filter(function(e){ return e.status === 'resolved'; }).length, sent = exceptions.filter(function(e){ return e.status === 'review'; }).length;
    var open = total - res, ai = exceptions.filter(function(e){ return e.source === 'ai' && e.status !== 'resolved'; }).length, md = medianDays(liveFindings);
    el.innerHTML = kpi('Open exceptions', open, sent + ' waiting on a PM · ' + (open - sent) + ' not sent yet', open ? 'up' : 'down') +
      kpi('Resolved', res + '<span class="unit">of ' + total + '</span>', (total ? Math.round(res * 100 / total) : 0) + '% of flagged items', 'flat') +
      kpi('Open with an AI note', ai, 'explained with AI, still open', 'flat') +
      kpi('Median days to resolve', md == null ? '—' : (Math.round(md * 10) / 10) + '<span class="unit">days</span>', md == null ? 'no sent-then-resolved items yet' : 'from sent to PM until resolved', 'flat');
  }
  function renderLiveRouting(){
    var tb = document.getElementById('dashRouting'); if(!tb) return;
    if(!liveLoaded){ tb.innerHTML = '<tr><td colspan="3" class="empty-note">Loads after sign-in.</td></tr>'; return; }
    var by = {}, notSent = 0;
    exceptions.forEach(function(e){
      if(e.status === 'open'){ notSent++; return; }
      var k = e.reviewer; var o = by[k] = by[k] || { name: k, waiting: 0, done: 0 };
      if(e.status === 'review') o.waiting++; else o.done++;
    });
    var rows = Object.keys(by).map(function(k){ return by[k]; }).sort(function(a, b){ return (b.waiting - a.waiting) || (b.done - a.done); });
    var max = Math.max.apply(null, [notSent].concat(rows.map(function(r){ return r.waiting; })).concat([1]));
    tb.innerHTML = rows.slice(0, 12).map(function(r){
      return '<tr><td><div class="route-name">' + dEsc(r.name) + '</div><div class="route-role">Project manager</div></td><td>' + r.waiting + ' waiting · ' + r.done + ' resolved</td><td class="num"><span class="load-bar"><span style="width:' + Math.round(r.waiting * 100 / max) + '%"></span></span>' + r.waiting + '</td></tr>';
    }).join('') + '<tr><td><div class="route-name">Not sent yet</div><div class="route-role">Needs a PM</div></td><td>Open findings with no PM</td><td class="num"><span class="load-bar"><span style="width:' + Math.round(notSent * 100 / max) + '%"></span></span>' + notSent + '</td></tr>';
  }
  function renderLiveTrend(){
    var by = {};
    liveFindings.forEach(function(f){ var w = f.weekStart || '?'; var o = by[w] = by[w] || { w: w, flagged: 0, resolved: 0 }; o.flagged++; if(f.status === 'Resolved') o.resolved++; });
    var weeks = Object.keys(by).sort().slice(-6).map(function(k){ return by[k]; });
    if(!weeks.length){ drawTrend([], [], []); return; }
    drawTrend(weeks.map(function(w){ return dDay(w.w); }), weeks.map(function(w){ return w.flagged; }), weeks.map(function(w){ return w.resolved; }));
  }
  function renderAllLive(){ renderLiveKpis(); renderQueue(); renderDetail(); renderLiveRouting(); renderLiveTrend();
    var sub = document.getElementById('dashQueueSub'); if(sub) sub.textContent = liveLoaded ? (exceptions.length + ' findings across ' + Object.keys(liveFindings.reduce(function(a, f){ a[f.weekStart] = 1; return a; }, {})).length + ' saved week(s) · live from SharePoint') : 'not loaded'; }
  async function dashRefresh(){
    if(liveBusy || typeof upaeFetchFindings !== 'function') return;
    liveBusy = true; liveStale = false;
    var sub = document.getElementById('dashQueueSub'); if(sub) sub.textContent = 'refreshing…';
    try{ liveFindings = await upaeFetchFindings(); liveError = null; liveLoaded = true; exceptions = buildExceptions(liveFindings); }
    catch(err){ console.warn('UPAE: dashboard findings', err); liveError = err.message || String(err); liveLoaded = false; }
    liveHistory = null; liveHistError = null;
    if(liveLoaded && typeof upaeFetchHistory === 'function'){ try{ liveHistory = await upaeFetchHistory(); } catch(err){ liveHistError = 'not set up yet'; } }
    liveStamp = Date.now(); liveBusy = false;
    renderAllLive();
  }
  window.upaeDashRefresh = dashRefresh;
  window.upaeDashInvalidate = function(){ liveStale = true; window.__dashStale = true; };
  renderAllLive();
  var refreshBtn = document.getElementById('dashRefreshBtn'); if(refreshBtn) refreshBtn.addEventListener('click', dashRefresh);

  // ---- filter chips ----
  (function initFilters(){
    var chips = document.querySelectorAll('#filters .chip-filter');
    chips.forEach(function(chip){
      chip.addEventListener('click', function(){
        activeFilter = chip.getAttribute('data-filter');
        chips.forEach(function(c){ c.classList.toggle('active', c === chip); });
        renderQueue();
      });
    });
  })();

  // ---- view switching (Dashboard / Help) ----
  (function initViews(){
    var navLinks = document.querySelectorAll('.rail nav a[data-view]');
    var allTriggers = document.querySelectorAll('a[data-view]');
    var views = { dashboard: document.getElementById('view-dashboard'), insights: document.getElementById('view-insights'), rules: document.getElementById('view-rules'), admin: document.getElementById('view-admin'), employees: document.getElementById('view-employees'), hours: document.getElementById('view-hours'), help: document.getElementById('view-help') };
    allTriggers.forEach(function(link){
      link.addEventListener('click', function(e){
        e.preventDefault();
        var target = link.getAttribute('data-view');
        navLinks.forEach(function(l){ l.classList.toggle('active', l.getAttribute('data-view') === target); });
        Object.keys(views).forEach(function(key){ views[key].hidden = (key !== target); });
        if(target === 'dashboard' && typeof window.upaeDashRefresh === 'function' && (Date.now() - (window.__dashStamp || 0) > 120000 || window.__dashStale)){ window.__dashStamp = Date.now(); window.__dashStale = false; window.upaeDashRefresh(); }
        if(views[target]) views[target].scrollIntoView({block:'start'});
      });
    });
  })();

  // ---- run validation modal ----
  (function initRunModal(){
    var jobs = [
      { code:'24-118', name:'Girard Ave. Overpass', union:'IUOE 542' },
      { code:'24-142', name:'Delancey St. Housing', union:'IBEW 98' },
      { code:'24-095', name:'Roosevelt Blvd. Interchange', union:'Laborers 413' },
      { code:'23-871', name:'Kensington Ave. Streetscape', union:'IUOE 542' }
    ];

    var overlay = document.getElementById('runModalOverlay');
    var openBtn = document.getElementById('openRunModal');
    var closeBtn = document.getElementById('closeRunModal');
    var cancelBtn = document.getElementById('cancelRun');
    var startBtn = document.getElementById('startRun');
    var viewResultsBtn = document.getElementById('viewResultsBtn');
    var jobListEl = document.getElementById('jobList');
    var jobListLabel = document.getElementById('jobListLabel');
    var multiCount = document.getElementById('multiCount');
    var runSummary = document.getElementById('runSummary');
    var weekSelect = document.getElementById('weekEndingSelect');
    var runForm = document.getElementById('runForm');
    var runProgress = document.getElementById('runProgress');
    var runActions = document.getElementById('runActions');
    var runResult = document.getElementById('runResult');

    function currentScope(){
      var checked = document.querySelector('input[name="scope"]:checked');
      return checked ? checked.value : 'single';
    }

    function selectedJobs(){
      var scope = currentScope();
      if(scope === 'all') return jobs.slice();
      if(scope === 'single'){
        var r = document.querySelector('input[name="job-single"]:checked');
        return r ? [jobs[parseInt(r.value,10)]] : [];
      }
      var checks = document.querySelectorAll('input[name="job-multi"]:checked');
      return Array.prototype.map.call(checks, function(c){ return jobs[parseInt(c.value,10)]; });
    }

    function renderJobList(){
      var scope = currentScope();
      jobListEl.innerHTML = '';
      if(scope === 'all'){
        jobListLabel.textContent = 'Jobs included';
        jobs.forEach(function(j){
          jobListEl.insertAdjacentHTML('beforeend', jobRowHtml(j, false, true, true));
        });
      } else if(scope === 'single'){
        jobListLabel.textContent = 'Job';
        jobs.forEach(function(j, i){
          jobListEl.insertAdjacentHTML('beforeend', jobRowHtml(j, false, i === 0, false, 'radio', 'job-single', i));
        });
      } else {
        jobListLabel.textContent = 'Jobs (select one or more)';
        jobs.forEach(function(j, i){
          jobListEl.insertAdjacentHTML('beforeend', jobRowHtml(j, false, i === 0, false, 'checkbox', 'job-multi', i));
        });
      }
      updateSummary();
    }

    function jobRowHtml(j, disabled, checked, forceChecked, type, name, value){
      type = type || 'checkbox';
      name = name || 'job-all';
      var dis = forceChecked ? 'disabled' : '';
      return '<label class="job-row">' +
        '<input type="' + type + '" name="' + name + '" value="' + (value!==undefined?value:'') + '" ' + ((checked||forceChecked)?'checked':'') + ' ' + dis + '>' +
        '<span class="job-code">' + j.code + '</span>' +
        '<span class="job-name">' + j.name + '</span>' +
        '<span class="job-union">' + j.union + '</span>' +
      '</label>';
    }

    function updateSummary(){
      var jobsSel = selectedJobs();
      var week = weekSelect.value.split(' ')[0];
      var scope = currentScope();
      var label;
      if(scope === 'all'){
        label = 'all <b>' + jobs.length + ' active jobs</b>';
      } else if(jobsSel.length === 0){
        label = '<b>no jobs selected</b>';
      } else if(jobsSel.length === 1){
        label = '<b>Job ' + jobsSel[0].code + ' &middot; ' + jobsSel[0].name + '</b>';
      } else {
        label = '<b>' + jobsSel.length + ' selected jobs</b>';
      }
      runSummary.innerHTML = 'Will run against ' + label + ' for week ending <b>' + week + '</b>.';
      multiCount.textContent = (scope === 'multiple' ? jobsSel.length : document.querySelectorAll('input[name="job-multi"]:checked').length) + ' selected';
      startBtn.disabled = jobsSel.length === 0;
      startBtn.style.opacity = startBtn.disabled ? '0.5' : '1';
    }

    function resetModal(){
      runForm.hidden = false;
      runResult.classList.remove('active');
      runProgress.classList.remove('active');
      runActions.style.display = 'flex';
      startBtn.textContent = 'Run validation';
      startBtn.disabled = false;
      document.querySelectorAll('input[name="scope"]').forEach(function(r){ r.checked = (r.value === 'single'); });
      renderJobList();
    }

    function openModal(){
      overlay.hidden = false;
      resetModal();
    }
    function closeModal(){ overlay.hidden = true; }

    openBtn.addEventListener('click', openModal);
    closeBtn.addEventListener('click', closeModal);
    cancelBtn.addEventListener('click', closeModal);
    overlay.addEventListener('click', function(e){ if(e.target === overlay) closeModal(); });

    document.querySelectorAll('input[name="scope"]').forEach(function(r){
      r.addEventListener('change', renderJobList);
    });
    jobListEl.addEventListener('change', updateSummary);
    weekSelect.addEventListener('change', updateSummary);

    startBtn.addEventListener('click', function(){
      var jobsSel = selectedJobs();
      if(jobsSel.length === 0) return;
      runProgress.classList.add('active');
      startBtn.disabled = true;
      startBtn.textContent = 'Running…';
      var steps = [
        'Pulling CMiC timesheet and pay-rate data…',
        'Running rules engine against ' + jobsSel.length + ' job' + (jobsSel.length>1?'s':'') + '…',
        'Scoring AI anomaly pass…',
        'Generating exception reports…'
      ];
      var stepEl = document.getElementById('runProgressText');
      var i = 0;
      stepEl.textContent = steps[0];
      var interval = setInterval(function(){
        i++;
        if(i < steps.length){ stepEl.textContent = steps[i]; }
      }, 550);
      setTimeout(function(){
        clearInterval(interval);
        var week = weekSelect.value.split(' ')[0];
        var newCount = Math.max(1, Math.min(9, jobsSel.length * 2 + 1));
        var autoCount = Math.max(0, jobsSel.length);
        var cleanCount = jobsSel.length * 11 + 3;
        document.getElementById('statNew').textContent = newCount;
        document.getElementById('statAuto').textContent = autoCount;
        document.getElementById('statClean').textContent = cleanCount;
        document.getElementById('runResultSummary').textContent =
          'Ran against ' + jobsSel.length + ' job' + (jobsSel.length>1?'s':'') + ' for week ending ' + week + '.';
        runForm.hidden = true;
        runResult.classList.add('active');
      }, 550 * steps.length + 200);
    });

    viewResultsBtn.addEventListener('click', function(){
      closeModal();
      var dashLink = document.querySelector('.rail nav a[data-view="dashboard"]');
      if(dashLink) dashLink.click();
      var queuePanel = document.getElementById('queue');
      if(queuePanel) queuePanel.scrollIntoView({behavior:'smooth', block:'center'});
    });

    renderJobList();
  })();

  // ---- trend chart ----
  function drawTrend(periods, ruleData, aiData){
    var svg = document.getElementById('trendChart');
    while(svg.firstChild) svg.removeChild(svg.firstChild);
    if(!periods.length){ var t = document.createElementNS('http://www.w3.org/2000/svg', 'text'); t.setAttribute('x', 280); t.setAttribute('y', 100); t.setAttribute('text-anchor', 'middle'); t.setAttribute('font-size', 12); t.setAttribute('fill', 'var(--ink-faint)'); t.textContent = 'No saved weeks yet'; svg.appendChild(t); return; }

    var W = 560, H = 200;
    var padL = 30, padR = 12, padT = 14, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var maxV = Math.max(4, Math.ceil(Math.max.apply(null, ruleData.concat(aiData)) / 4) * 4);
    var n = periods.length;

    function x(i){ return n === 1 ? padL + plotW / 2 : padL + (plotW / (n - 1)) * i; }
    function y(v){ return padT + plotH - (v / maxV) * plotH; }

    var ns = 'http://www.w3.org/2000/svg';
    function el(tag, attrs){
      var e = document.createElementNS(ns, tag);
      for(var k in attrs){ e.setAttribute(k, attrs[k]); }
      return e;
    }

    // gridlines + y labels
    [0, maxV / 4, maxV / 2, maxV * 3 / 4, maxV].forEach(function(v){
      var gl = el('line', {x1:padL, x2:W-padR, y1:y(v), y2:y(v), stroke:'var(--line-soft)', 'stroke-width':1});
      svg.appendChild(gl);
      var lab = el('text', {x:padL - 8, y:y(v)+3, 'text-anchor':'end', 'font-size':10, fill:'var(--ink-faint)', 'font-family':'IBM Plex Mono, monospace'});
      lab.textContent = Math.round(v);
      svg.appendChild(lab);
    });

    // x labels
    periods.forEach(function(p, i){
      var lab = el('text', {x:x(i), y:H-6, 'text-anchor':'middle', 'font-size':10, fill:'var(--ink-faint)', 'font-family':'IBM Plex Mono, monospace'});
      lab.textContent = p;
      svg.appendChild(lab);
    });

    function path(data){
      return data.map(function(v,i){ return (i===0?'M':'L') + x(i).toFixed(1) + ' ' + y(v).toFixed(1); }).join(' ');
    }
    function areaPath(data){
      return 'M' + x(0) + ' ' + y(0) + ' ' + data.map(function(v,i){ return 'L' + x(i).toFixed(1) + ' ' + y(v).toFixed(1); }).join(' ') + ' L' + x(n-1) + ' ' + y(0) + ' Z';
    }

    var ruleArea = el('path', {d:areaPath(ruleData), fill:'var(--accent)', opacity:'0.10', stroke:'none'});
    svg.appendChild(ruleArea);
    var ruleLine = el('path', {d:path(ruleData), fill:'none', stroke:'var(--accent)', 'stroke-width':2.5, 'stroke-linecap':'round', 'stroke-linejoin':'round'});
    svg.appendChild(ruleLine);
    var aiLine = el('path', {d:path(aiData), fill:'none', stroke:'var(--ai-tag)', 'stroke-width':2.5, 'stroke-linecap':'round', 'stroke-linejoin':'round', 'stroke-dasharray':'0'});
    svg.appendChild(aiLine);

    ruleData.forEach(function(v,i){
      svg.appendChild(el('circle', {cx:x(i), cy:y(v), r:i===n-1?3.5:2.5, fill:'var(--accent)'}));
    });
    aiData.forEach(function(v,i){
      svg.appendChild(el('circle', {cx:x(i), cy:y(v), r:i===n-1?3.5:2.5, fill:'var(--ai-tag)'}));
    });

    // endpoint labels
    var lastRule = el('text', {x:x(n-1)+7, y:y(ruleData[n-1])+4, 'font-size':11, fill:'var(--accent)', 'font-family':'IBM Plex Mono, monospace', 'font-weight':'600'});
    lastRule.textContent = ruleData[n-1];
    svg.appendChild(lastRule);
    var lastAi = el('text', {x:x(n-1)+7, y:y(aiData[n-1])+4, 'font-size':11, fill:'var(--ai-tag)', 'font-family':'IBM Plex Mono, monospace', 'font-weight':'600'});
    lastAi.textContent = aiData[n-1];
    svg.appendChild(lastAi);
  }

  // ---- insights: category / dollar-impact bar lists ----
  (function drawInsightBars(){
    function renderBars(containerId, rows, formatValue){
      var el = document.getElementById(containerId);
      if(!el) return;
      var max = Math.max.apply(null, rows.map(function(r){ return r.value; }));
      el.innerHTML = rows.map(function(r){
        var pct = Math.round((r.value / max) * 100);
        return '<div class="ibar-row">' +
          '<span class="ibar-label">' + r.label + '</span>' +
          '<span class="ibar-track"><span style="width:' + pct + '%"></span></span>' +
          '<span class="ibar-value">' + formatValue(r.value) + '</span>' +
        '</div>';
      }).join('');
    }

    renderBars('catBarList', [
      { label: 'Overtime miscalculation', value: 42 },
      { label: 'Union / classification mismatch', value: 31 },
      { label: 'Missing hours', value: 27 },
      { label: 'Job cost allocation error', value: 19 },
      { label: 'Duplicate entries', value: 14 },
      { label: 'Union work-rule violation', value: 11 }
    ], function(v){ return v; });

    renderBars('dollarBarList', [
      { label: 'Overtime miscalculation', value: 58400 },
      { label: 'Job cost allocation error', value: 34200 },
      { label: 'Union / classification mismatch', value: 27900 },
      { label: 'Missing hours', value: 14600 },
      { label: 'Duplicate entries', value: 7300 },
      { label: 'Union work-rule violation', value: 4400 }
    ], function(v){ return '$' + v.toLocaleString('en-US'); });
  })();

  // ---- insights: detection trend chart ----
  (function drawInsightsTrend(){
    var svg = document.getElementById('insightsTrendChart');
    if(!svg) return;
    var periods = ['Cyc 1','Cyc 2','Cyc 3','Cyc 4','Cyc 5','Cyc 6','Cyc 7','Cyc 8'];
    var ruleData = [9, 11, 10, 13, 12, 15, 14, 16];
    var aiData   = [2, 3, 4, 3, 5, 6, 6, 8];

    var W = 560, H = 200;
    var padL = 24, padR = 12, padT = 14, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var maxV = 18;
    var n = periods.length;

    function x(i){ return padL + (plotW / (n - 1)) * i; }
    function y(v){ return padT + plotH - (v / maxV) * plotH; }

    var ns = 'http://www.w3.org/2000/svg';
    function el(tag, attrs){
      var e = document.createElementNS(ns, tag);
      for(var k in attrs){ e.setAttribute(k, attrs[k]); }
      return e;
    }

    [0, 6, 12, 18].forEach(function(v){
      var gl = el('line', {x1:padL, x2:W-padR, y1:y(v), y2:y(v), stroke:'var(--line-soft)', 'stroke-width':1});
      svg.appendChild(gl);
      var lab = el('text', {x:padL - 6, y:y(v)+3, 'text-anchor':'end', 'font-size':9, fill:'var(--ink-faint)', 'font-family':'IBM Plex Mono, monospace'});
      lab.textContent = v;
      svg.appendChild(lab);
    });

    periods.forEach(function(p, i){
      var lab = el('text', {x:x(i), y:H-6, 'text-anchor':'middle', 'font-size':9.5, fill:'var(--ink-faint)', 'font-family':'IBM Plex Mono, monospace'});
      lab.textContent = p;
      svg.appendChild(lab);
    });

    function path(data){
      return data.map(function(v,i){ return (i===0?'M':'L') + x(i).toFixed(1) + ' ' + y(v).toFixed(1); }).join(' ');
    }
    function areaPath(data){
      return 'M' + x(0) + ' ' + y(0) + ' ' + data.map(function(v,i){ return 'L' + x(i).toFixed(1) + ' ' + y(v).toFixed(1); }).join(' ') + ' L' + x(n-1) + ' ' + y(0) + ' Z';
    }

    svg.appendChild(el('path', {d:areaPath(ruleData), fill:'var(--accent)', opacity:'0.10', stroke:'none'}));
    svg.appendChild(el('path', {d:path(ruleData), fill:'none', stroke:'var(--accent)', 'stroke-width':2.5, 'stroke-linecap':'round', 'stroke-linejoin':'round'}));
    svg.appendChild(el('path', {d:path(aiData), fill:'none', stroke:'var(--ai-tag)', 'stroke-width':2.5, 'stroke-linecap':'round', 'stroke-linejoin':'round'}));

    ruleData.forEach(function(v,i){
      var c = el('circle', {cx:x(i), cy:y(v), r:i===n-1?3.5:2.5, fill:'var(--accent)'});
      var t = el('title', {}); t.textContent = periods[i] + ' — Rule-based: ' + v;
      c.appendChild(t);
      svg.appendChild(c);
    });
    aiData.forEach(function(v,i){
      var c = el('circle', {cx:x(i), cy:y(v), r:i===n-1?3.5:2.5, fill:'var(--ai-tag)'});
      var t = el('title', {}); t.textContent = periods[i] + ' — AI-flagged: ' + v;
      c.appendChild(t);
      svg.appendChild(c);
    });

    var lastRule = el('text', {x:x(n-1)+7, y:y(ruleData[n-1])+4, 'font-size':11, fill:'var(--accent)', 'font-family':'IBM Plex Mono, monospace', 'font-weight':'600'});
    lastRule.textContent = ruleData[n-1];
    svg.appendChild(lastRule);
    var lastAi2 = el('text', {x:x(n-1)+7, y:y(aiData[n-1])+4, 'font-size':11, fill:'var(--ai-tag)', 'font-family':'IBM Plex Mono, monospace', 'font-weight':'600'});
    lastAi2.textContent = aiData[n-1];
    svg.appendChild(lastAi2);
  })();
})();
