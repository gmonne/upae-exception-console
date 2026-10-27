// Patterns screen (Phase 3): statistical outliers from saved weeks, optional AI interpretation, advisory only.
var patSaved = null, patSavedErr = '', patRes = null, patItems = [], patAi = { state: 'idle', err: '', overview: '', dataLimits: '' }, patWeek = '', patFilter = 'active', patBusy = false, patMsg = '';
var PAT_CFG_KEY = 'upae.patterns.cfg';
function patLoadCfg(){ try{ return JSON.parse(localStorage.getItem(PAT_CFG_KEY) || '{}'); }catch(e){ return {}; } }
function patSaveCfg(c){ try{ localStorage.setItem(PAT_CFG_KEY, JSON.stringify(c)); }catch(e){} }
function patUnionOf(id){ var k = empKey(id), e = null; for(var i = 0; i < employees.length; i++){ if(empKey(employees[i].id) === k){ e = employees[i]; break; } } return e ? empUnionName(e) : ''; }
function patWho(){ return typeof upaeCurrentUserEmail === 'function' ? upaeCurrentUserEmail() : ''; }
var PAT_LEAN = { worth_checking: ['Worth a look', 'draft'], likely_routine: ['Likely routine', 'ok'], unclear: ['Unclear', 'muted'] };

function patMerge(){
  var by = {}; (patSaved || []).forEach(function(r){ by[r.key] = r; });
  patItems = (patRes ? patRes.outliers : []).map(function(o){
    var s = by[o.key] || null, it = Object.assign({ status: 'New', interpretation: '', leaning: '', confidence: '', suggestedCheck: '', evidence: [], spItemId: '' }, o);
    if(s){ ['spItemId', 'status', 'interpretation', 'leaning', 'confidence', 'suggestedCheck', 'evidence', 'reviewedBy', 'reviewedDate', 'model', 'promptVersion'].forEach(function(k){ if(s[k] != null && s[k] !== '') it[k] = s[k]; }); }
    return it;
  });
}
function patRun(){
  var sel = document.getElementById('patWeekSel'); patWeek = sel ? sel.value : patWeek;
  var cfg = {}; ['repeatWeeks', 'spikeRatio', 'spikeHours', 'unionPts', 'longHours'].forEach(function(k){ var f = document.getElementById('patCfg_' + k); if(f && f.value !== '') cfg[k] = f.value; });
  patSaveCfg(cfg);
  patRes = patDetect(hrsSaved || [], patUnionOf, patWeek, cfg); patAi = { state: 'idle', err: '', overview: '', dataLimits: '' }; patMsg = '';
  patMerge(); renderPatterns();
}
function patPayload(){
  var wk = (hrsSavedWeeks().filter(function(w){ return w.start === patWeek; })[0]) || {};
  return { context: { weekStart: patWeek, weekNo: wk.weekNo == null ? null : wk.weekNo, weeksOfHistory: patRes.weeksOfHistory, priorWeeks: patRes.priorWeeks },
    outliers: patItems.slice(0, 15).map(function(i){ return { key: i.key, type: i.type, employee: i.employee, union: i.union, facts: i.facts, metrics: i.metrics }; }) };
}
async function patInterpret(){
  if(!patItems.length || patBusy) return;
  patBusy = true; patAi = { state: 'loading', err: '', overview: '', dataLimits: '' }; renderPatterns();
  try{
    if(typeof upaeAiPatterns !== 'function') throw new Error('The AI service is not available in this view.');
    var d = await upaeAiPatterns(patPayload()), by = {}; d.items.forEach(function(x){ by[x.key] = x; });
    patItems.forEach(function(i){ var x = by[i.key]; if(x){ i.interpretation = x.interpretation; i.leaning = x.leaning; i.confidence = x.confidence; i.suggestedCheck = x.suggestedCheck; i.evidence = x.evidence; i.model = d.model; i.promptVersion = d.promptVersion; i.fresh = true; } });
    patAi = { state: 'done', err: '', overview: d.overview, dataLimits: d.dataLimits, stub: !!d.stub, usage: d.usage };
    if(typeof upaeCreateAILog === 'function' && !d.stub) upaeCreateAILog({ purpose: 'patterns', findingKey: '', weekStart: patWeek, userEmail: patWho(), model: d.model, promptVersion: d.promptVersion, confidence: '', usage: d.usage, output: JSON.stringify({ overview: d.overview, items: d.items.length }) }).catch(function(e){ console.warn('UPAE: could not write AILog', e); });
  }catch(err){
    console.error('UPAE: patterns failed', err);
    if(typeof upaeLogFailedUsage === 'function') upaeLogFailedUsage('patterns', err, patWeek);
    patAi = { state: 'error', err: err.message || String(err), overview: '', dataLimits: '' };
  }
  patBusy = false; renderPatterns();
}
async function patPersist(items){
  if(typeof upaeSavePattern !== 'function'){ patMsg = 'Not connected to SharePoint.'; renderPatterns(); return; }
  patBusy = true; renderPatterns(); var ok = 0, fail = '';
  for(var n = 0; n < items.length; n++){
    try{ var r = await upaeSavePattern(items[n]); if(r && r.id) items[n].spItemId = r.id; ok++; items[n].fresh = false; }
    catch(err){ console.warn('UPAE: could not save pattern', err); fail = err.message || String(err); }
  }
  patSaved = patSaved || []; items.forEach(function(i){ if(!i.spItemId) return; var e = patSaved.filter(function(s){ return s.key === i.key; })[0]; if(e) Object.assign(e, i); else patSaved.push(Object.assign({}, i)); });
  patMsg = fail ? 'Saved ' + ok + ', but some failed: ' + fail + ' Check that the PatternFindings list exists with all its columns.' : 'Saved ' + ok + ' to the PatternFindings list.';
  patBusy = false; renderPatterns();
}
function patSetStatus(key, status){
  var it = patItems.filter(function(i){ return i.key === key; })[0]; if(!it) return;
  it.status = status; it.reviewedBy = status === 'New' ? '' : patWho(); it.reviewedDate = status === 'New' ? '' : new Date().toISOString().slice(0, 10);
  patPersist([it]);
}
function patCard(i){
  var lean = PAT_LEAN[i.leaning], hasAi = !!i.interpretation;
  return '<div class="hrs-row ' + (i.status === 'Dismissed' ? '' : 'crit') + '" style="display:block;padding:12px 14px;' + (i.status === 'Dismissed' ? 'opacity:.6;' : '') + '">' +
    '<div><span class="cov-pill draft">' + escapeHtml(PAT_TYPES[i.type] || i.type) + '</span> <b>' + escapeHtml(i.employee || i.union) + '</b>' + (i.employee && i.union ? ' <span class="e-sub" style="display:inline;">' + escapeHtml(i.union) + '</span>' : '') +
      ' <span class="h-msg" style="font-size:12px;">' + escapeHtml(i.status) + (i.reviewedBy ? ' · ' + escapeHtml(i.reviewedBy) : '') + (i.reviewedDate ? ' · ' + escapeHtml(i.reviewedDate) : '') + '</span></div>' +
    i.facts.map(function(f){ return '<div style="margin-top:4px;">' + escapeHtml(f) + '</div>'; }).join('') +
    (hasAi ? '<div class="ai-card" style="margin-top:8px;"><div class="ai-head"><span class="ai-tag">AI</span>' + (lean ? '<span class="cov-pill ' + lean[1] + '">' + lean[0] + '</span>' : '') + '<span class="ai-conf">Confidence: ' + escapeHtml(i.confidence || 'low') + '</span></div>' +
      '<p style="margin:6px 0;">' + escapeHtml(i.interpretation) + '</p>' + (i.evidence && i.evidence.length ? '<ul class="sum-list">' + i.evidence.map(function(e){ return '<li>' + escapeHtml(e) + '</li>'; }).join('') + '</ul>' : '') +
      (i.suggestedCheck ? '<p><b>Check:</b> ' + escapeHtml(i.suggestedCheck) + '</p>' : '') + '<div class="ai-note">Advisory. Not a rule violation. Please verify before acting.</div></div>' : '') +
    '<div class="emp-actions" style="padding:8px 0 0;">' +
      (i.status === 'New' ? '<button class="btn" type="button" data-pat-st="Reviewed" data-key="' + escapeHtml(i.key) + '">Mark reviewed</button><button class="btn ghost" type="button" data-pat-st="Dismissed" data-key="' + escapeHtml(i.key) + '">Dismiss</button>' : '<button class="btn ghost" type="button" data-pat-st="New" data-key="' + escapeHtml(i.key) + '">Reopen</button>') +
      (!i.spItemId || i.fresh ? '<button class="btn ghost" type="button" data-pat-save="' + escapeHtml(i.key) + '">Save</button>' : '') + '</div></div>';
}
function renderPatterns(){
  var el = document.getElementById('patBody'); if(!el) return;
  if(hrsSavedError || !hrsLoadedOnce || !hrsSaved){ el.innerHTML = '<div class="emp-status' + (hrsSavedError ? ' err' : '') + '">' + (hrsSavedError ? 'Saved weeks could not be read (' + escapeHtml(hrsSavedError) + ').' : hrsLoadedOnce ? 'No saved weeks yet. Save weeks on the Weekly Hours screen first.' : 'Saved weeks load after you sign in.') + '</div>'; return; }
  var weeks = hrsSavedWeeks(); if(!patWeek && weeks.length) patWeek = weeks[0].start;
  var cfg = patCfg(patLoadCfg());
  var html = '<div class="panel"><div class="panel-head"><h2>Find patterns</h2><span class="panel-sub">' + weeks.length + ' saved week' + (weeks.length === 1 ? '' : 's') + ' available</span></div><div class="emp-import" style="gap:12px;flex-wrap:wrap;">' +
    '<label>Week <select id="patWeekSel">' + weeks.map(function(w){ return '<option value="' + w.start + '"' + (w.start === patWeek ? ' selected' : '') + '>Week ' + (w.weekNo || '?') + ' · ' + fmtMD(w.start) + '–' + fmtMD(hrsAddDays(w.start, 6)) + '</option>'; }).join('') + '</select></label>' +
    '<button class="btn primary" type="button" id="patRunBtn"' + (patBusy || !weeks.length ? ' disabled' : '') + '>Find patterns</button>' +
    '<button class="btn" type="button" id="patAiBtn"' + (patBusy || !patItems.length ? ' disabled' : '') + '>Interpret with AI</button>' +
    '<button class="btn ghost" type="button" id="patSaveAll"' + (patBusy || !patItems.length ? ' disabled' : '') + '>Save all</button></div>' +
    '<details class="emp-details" style="margin:6px 14px 10px;"><summary>Thresholds</summary><div class="emp-import" style="gap:12px;flex-wrap:wrap;padding-top:8px;">' +
    [['repeatWeeks', 'Repeat in at least (weeks)', 1], ['longHours', 'Long day (hours)', 1], ['spikeRatio', 'Hours above own normal (times)', 0.1], ['spikeHours', '... and at least (extra hours)', 1], ['unionPts', 'Union overtime share rise (points)', 1]].map(function(c){ return '<label>' + c[1] + ' <input type="number" step="' + c[2] + '" min="1" id="patCfg_' + c[0] + '" value="' + cfg[c[0]] + '" style="width:70px;"></label>'; }).join('') + '</div></details></div>';
  if(patRes){
    var act = patItems.filter(function(i){ return i.status !== 'Dismissed'; }).length, dis = patItems.length - act;
    html += '<div class="emp-status">The statistical pass looked at ' + (patRes.weeksOfHistory + 1) + ' saved week' + (patRes.weeksOfHistory ? 's' : '') + ' and found <b>' + patItems.length + '</b> outlier' + (patItems.length === 1 ? '' : 's') + ' for this week. The same saved data always gives the same result.' + (patRes.notes.length ? '<br>' + patRes.notes.map(escapeHtml).join('<br>') : '') + '</div>';
    if(patAi.state === 'loading') html += '<div class="ai-card"><span class="ai-tag">AI</span> Interpreting the outliers&hellip;</div>';
    if(patAi.state === 'error') html += '<div class="ai-card err"><span class="ai-tag">AI</span> ' + escapeHtml(patAi.err) + '</div>';
    if(patAi.state === 'done') html += '<div class="ai-card"><div class="ai-head"><span class="ai-tag">AI</span>' + (patAi.stub ? '<span class="ai-stub">Sample answer &mdash; the AI key is not set up yet</span>' : '') + '</div><p><b>' + escapeHtml(patAi.overview) + '</b></p>' + (patAi.dataLimits ? '<p class="ai-warn"><b>Limits.</b> ' + escapeHtml(patAi.dataLimits) + '</p>' : '') + '</div>';
    if(patMsg) html += '<div class="emp-status">' + escapeHtml(patMsg) + '</div>';
    html += '<div class="filters" style="padding:8px 0;">' + [['active', 'Active (' + act + ')'], ['dismissed', 'Dismissed (' + dis + ')'], ['all', 'All']].map(function(f){ return '<button class="chip' + (patFilter === f[0] ? ' active' : '') + '" type="button" data-pat-filter="' + f[0] + '">' + f[1] + '</button>'; }).join('') + '</div>';
    var vis = patItems.filter(function(i){ return patFilter === 'all' || (patFilter === 'dismissed' ? i.status === 'Dismissed' : i.status !== 'Dismissed'); });
    html += vis.length ? '<div class="panel">' + vis.map(patCard).join('') + '</div>' : '<div class="emp-status">Nothing to show here.</div>';
  } else html += '<div class="emp-status">Choose a week and press <b>Find patterns</b>. The statistical pass is free and instant; the AI step is optional and uses one call for all outliers.</div>';
  html += '<div class="emp-status">' + (patSavedErr ? 'Saved pattern findings could not be read (' + escapeHtml(patSavedErr) + '). They can still be found and shown, but not remembered until the PatternFindings list exists.' : '') + '</div>';
  el.innerHTML = html;
  document.getElementById('patRunBtn').addEventListener('click', patRun);
  document.getElementById('patAiBtn').addEventListener('click', patInterpret);
  document.getElementById('patSaveAll').addEventListener('click', function(){ patPersist(patItems.slice()); });
  var ws = document.getElementById('patWeekSel'); ws.addEventListener('change', function(){ patWeek = ws.value; patRes = null; patItems = []; patAi = { state: 'idle', err: '', overview: '', dataLimits: '' }; patMsg = ''; renderPatterns(); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-pat-st]'), function(b){ b.addEventListener('click', function(){ patSetStatus(b.getAttribute('data-key'), b.getAttribute('data-pat-st')); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-pat-save]'), function(b){ b.addEventListener('click', function(){ var it = patItems.filter(function(i){ return i.key === b.getAttribute('data-pat-save'); })[0]; if(it) patPersist([it]); }); });
  Array.prototype.forEach.call(el.querySelectorAll('[data-pat-filter]'), function(b){ b.addEventListener('click', function(){ patFilter = b.getAttribute('data-pat-filter'); renderPatterns(); }); });
}
async function patOpen(){
  if(!hrsLoadedOnce && typeof hrsRefreshSaved === 'function') await hrsRefreshSaved();
  if(patSaved == null){
    patSavedErr = '';
    try{ patSaved = typeof upaeFetchPatterns === 'function' ? await upaeFetchPatterns() : []; }catch(err){ console.warn('UPAE: PatternFindings', err); patSavedErr = err.message || String(err); patSaved = []; }
  }
  if(patRes) patMerge(); renderPatterns();
}
(function(){
  Array.prototype.forEach.call(document.querySelectorAll('a[data-view="patterns"]'), function(a){ a.addEventListener('click', function(){ renderPatterns(); patOpen(); }); });
  renderPatterns();
})();
