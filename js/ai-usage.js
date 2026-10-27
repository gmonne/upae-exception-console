// AI usage screen: reads the AILog list and shows calls, tokens and an estimated cost. Prices and budget are entered by the user.
var aiuRows = null, aiuLoaded = false, aiuBusy = false, aiuErr = '';
var AIU_KEY = 'upae.aiusage.settings';
function aiuSettings(){
  var d = { inPrice: '', outPrice: '', budget: '' };
  try{ var o = JSON.parse(localStorage.getItem(AIU_KEY) || '{}'); ['inPrice','outPrice','budget'].forEach(function(k){ if(o[k] != null) d[k] = o[k]; }); }catch(e){}
  return d;
}
function aiuSave(d){ try{ localStorage.setItem(AIU_KEY, JSON.stringify(d)); }catch(e){} }
function aiuNum(v){ var x = parseFloat(v); return isFinite(x) && x >= 0 ? x : null; }
function aiuFmt(n){ return Math.round(n).toLocaleString('en-US'); }
function aiuMoney(n){ return '$' + (n < 10 ? n.toFixed(2) : Math.round(n).toLocaleString('en-US')); }
function aiuDay(d){ return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
function aiuCost(inT, outT, s){ var a = aiuNum(s.inPrice), b = aiuNum(s.outPrice); return (a == null || b == null) ? null : inT / 1e6 * a + outT / 1e6 * b; }

// Pure summary of the rows, so it can be tested without the screen.
function aiuSummarize(rows, now){
  now = now || new Date();
  var today = aiuDay(now), monthKey = today.slice(0, 7), weekAgo = new Date(now.getTime() - 6 * 86400000), weekStart = aiuDay(weekAgo);
  var S = { today: { calls: 0, tokens: 0, inT: 0, outT: 0 }, week: { calls: 0, tokens: 0, inT: 0, outT: 0 }, month: { calls: 0, tokens: 0, inT: 0, outT: 0 }, all: { calls: 0, tokens: 0 },
    noTokens: 0, failed: 0, days: {}, byUser: {}, byPurpose: {}, byModel: {} };
  (rows || []).forEach(function(r){
    var d = r.at ? aiuDay(new Date(r.at)) : '';
    var has = r.inputTokens != null || r.outputTokens != null, inT = r.inputTokens || 0, outT = r.outputTokens || 0, calls = has ? (r.calls || 1) : 1, tk = inT + outT;
    if(String(r.purpose).indexOf('(') === 0) return;
    S.all.calls += calls; S.all.tokens += tk;
    if(!has) S.noTokens++;
    if(r.outcome === 'failed') S.failed++;
    function add(o){ o.calls += calls; o.tokens += tk; o.inT += inT; o.outT += outT; }
    if(d === today) add(S.today);
    if(d >= weekStart && d <= today) add(S.week);
    if(d.slice(0, 7) === monthKey){
      add(S.month);
      var u = S.byUser[r.user || '(unknown)'] = S.byUser[r.user || '(unknown)'] || { calls: 0, tokens: 0, inT: 0, outT: 0 }; add(u);
      var p = S.byPurpose[r.purpose || '?'] = S.byPurpose[r.purpose || '?'] || { calls: 0, tokens: 0, inT: 0, outT: 0 }; add(p);
      var m = S.byModel[r.model || '?'] = S.byModel[r.model || '?'] || { calls: 0, tokens: 0, inT: 0, outT: 0 }; add(m);
    }
    if(d){ var dd = S.days[d] = S.days[d] || { explain: 0, summary: 0, other: 0 }; var k = r.purpose === 'explain' ? 'explain' : r.purpose === 'summary' ? 'summary' : 'other'; dd[k] += tk; }
  });
  return S;
}
function aiuChart(S, now){
  now = now || new Date();
  var days = [], i; for(i = 29; i >= 0; i--) days.push(aiuDay(new Date(now.getTime() - i * 86400000)));
  var max = 1; days.forEach(function(d){ var x = S.days[d]; if(x) max = Math.max(max, x.explain + x.summary + x.other); });
  var W = 640, H = 170, bw = W / 30, svg = '<svg viewBox="0 0 ' + W + ' ' + (H + 22) + '" width="100%" role="img" aria-label="Tokens used per day over the last 30 days">';
  days.forEach(function(d, n){
    var x = S.days[d] || { explain: 0, summary: 0, other: 0 }, y = H, parts = [['explain', '#1F6F6B'], ['summary', '#7A5FC0'], ['other', '#9a9a9a']];
    parts.forEach(function(p){ var v = x[p[0]]; if(!v) return; var h = Math.max(1, v / max * (H - 8)); y -= h; svg += '<rect x="' + (n * bw + 2) + '" y="' + y.toFixed(1) + '" width="' + (bw - 4) + '" height="' + h.toFixed(1) + '" fill="' + p[1] + '"><title>' + d + ' · ' + p[0] + ': ' + aiuFmt(v) + ' tokens</title></rect>'; });
    if(n % 5 === 0 || n === 29) svg += '<text x="' + (n * bw + bw / 2) + '" y="' + (H + 15) + '" font-size="10" text-anchor="middle" fill="#777">' + d.slice(5) + '</text>';
  });
  return svg + '<line x1="0" y1="' + H + '" x2="' + W + '" y2="' + H + '" stroke="#ccc"/></svg>' +
    '<div class="h-msg" style="font-size:12px;"><span style="color:#1F6F6B;">&#9632;</span> Explain &nbsp; <span style="color:#7A5FC0;">&#9632;</span> Summary &nbsp; Tokens per day, busiest day ' + aiuFmt(max) + '.</div>';
}
function aiuTable(title, obj, s){
  var keys = Object.keys(obj).sort(function(a, b){ return obj[b].tokens - obj[a].tokens; });
  if(!keys.length) return '';
  return '<h3 style="margin:14px 0 6px;font-size:14px;">' + title + '</h3><table class="rules-table"><thead><tr><th>' + (title.indexOf('person') >= 0 ? 'Person' : title.indexOf('purpose') >= 0 ? 'Purpose' : 'Model') + '</th><th>Calls</th><th>Input tokens</th><th>Output tokens</th><th>Avg per call</th><th>Est. cost</th></tr></thead><tbody>' +
    keys.map(function(k){ var o = obj[k], c = aiuCost(o.inT, o.outT, s); return '<tr><td>' + escapeHtml(k) + '</td><td class="num">' + aiuFmt(o.calls) + '</td><td class="num">' + aiuFmt(o.inT) + '</td><td class="num">' + aiuFmt(o.outT) + '</td><td class="num">' + (o.calls ? aiuFmt(o.tokens / o.calls) : '—') + '</td><td class="num">' + (c == null ? '—' : aiuMoney(c)) + '</td></tr>'; }).join('') + '</tbody></table>';
}
function renderAiUsage(){
  var el = document.getElementById('aiuBody'); if(!el) return;
  var s = aiuSettings(), st = document.getElementById('aiuStatus');
  ['aiuInPrice','aiuOutPrice','aiuBudget'].forEach(function(id){ var f = document.getElementById(id); if(f && document.activeElement !== f){ f.value = s[id === 'aiuInPrice' ? 'inPrice' : id === 'aiuOutPrice' ? 'outPrice' : 'budget']; } });
  if(aiuErr){ el.innerHTML = '<div class="emp-status err">' + escapeHtml(aiuErr) + '</div>'; return; }
  if(!aiuLoaded){ el.innerHTML = '<div class="emp-status">' + (aiuBusy ? 'Loading the AI log&hellip;' : 'Sign in to load the AI log.') + '</div>'; return; }
  var S = aiuSummarize(aiuRows), mc = aiuCost(S.month.inT, S.month.outT, s), b = aiuNum(s.budget);
  function kpi(l, v, sub, cls){ return '<div class="kpi"><div class="k-label">' + l + '</div><div class="k-value">' + v + '</div><div class="k-delta ' + (cls || 'flat') + '">' + sub + '</div></div>'; }
  var bud = (mc != null && b) ? Math.round(mc / b * 100) : null;
  var html = '<section class="kpis" style="margin-bottom:12px;">' +
    kpi('Calls today', aiuFmt(S.today.calls), aiuFmt(S.today.tokens) + ' tokens') +
    kpi('Last 7 days', aiuFmt(S.week.calls), aiuFmt(S.week.tokens) + ' tokens') +
    kpi('This month', aiuFmt(S.month.calls), aiuFmt(S.month.tokens) + ' tokens' + (S.month.calls ? ' · ' + aiuFmt(S.month.tokens / S.month.calls) + ' per call' : '')) +
    kpi('Estimated cost this month', mc == null ? '—' : aiuMoney(mc), mc == null ? 'enter the two prices below' : (bud == null ? 'no budget set' : bud + '% of the ' + aiuMoney(b) + ' budget'), bud != null && bud >= 80 ? 'down' : 'flat') + '</section>';
  if(bud != null && bud >= 80) html += '<div class="emp-status err">' + (bud >= 100 ? 'The estimated cost has passed the monthly budget.' : 'The estimated cost is at ' + bud + '% of the monthly budget.') + ' Check the Anthropic Console for the real figure.</div>';
  html += '<div class="panel" style="margin-top:12px;"><div class="panel-head"><h2>Tokens per day</h2><span class="panel-sub">Last 30 days</span></div><div style="padding:8px 14px;">' + aiuChart(S) + '</div></div>';
  html += '<div class="panel"><div class="panel-head"><h2>This month in detail</h2><span class="panel-sub">Estimates use the prices you entered</span></div><div style="padding:0 14px 12px;">' +
    aiuTable('By person', S.byUser, s) + aiuTable('By purpose', S.byPurpose, s) + aiuTable('By model', S.byModel, s) + (S.month.calls ? '' : '<div class="emp-status">No AI calls logged this month.</div>') + '</div></div>';
  var notes = [];
  if(S.noTokens) notes.push(aiuFmt(S.noTokens) + ' logged call' + (S.noTokens === 1 ? '' : 's') + ' have no token count (made before token logging was added, or the token columns are missing from the AILog list). They are counted as one call each and add no tokens.');
  if(S.failed) notes.push(aiuFmt(S.failed) + ' call' + (S.failed === 1 ? '' : 's') + ' failed after using tokens; they are included.');
  notes.push('This screen only sees calls made through this console and logged to the AILog list. The Anthropic Console usage page is the source of truth for the bill, and is where the monthly spend limit is set.');
  html += '<div class="emp-status">' + notes.map(escapeHtml).join('<br>') + '</div>';
  el.innerHTML = html;
}
async function aiuLoad(){
  if(aiuBusy) return; aiuBusy = true; aiuErr = ''; renderAiUsage();
  try{ if(typeof upaeFetchAILog !== 'function') throw new Error('The log is not available in this view.'); aiuRows = await upaeFetchAILog(); aiuLoaded = true; }
  catch(err){ aiuErr = 'The AILog list could not be read (' + (err.message || err) + ').'; }
  aiuBusy = false; renderAiUsage();
}
(function(){
  Array.prototype.forEach.call(document.querySelectorAll('a[data-view="usage"]'), function(a){ a.addEventListener('click', function(){ if(!aiuLoaded) aiuLoad(); else renderAiUsage(); }); });
  var r = document.getElementById('aiuRefresh'); if(r) r.addEventListener('click', aiuLoad);
  ['aiuInPrice','aiuOutPrice','aiuBudget'].forEach(function(id){
    var f = document.getElementById(id); if(!f) return;
    f.addEventListener('input', function(){ aiuSave({ inPrice: document.getElementById('aiuInPrice').value, outPrice: document.getElementById('aiuOutPrice').value, budget: document.getElementById('aiuBudget').value }); if(aiuLoaded) renderAiUsage(); });
  });
  renderAiUsage();
})();
