// ======================================================================
// Employees: import from the B2W Employee Listing + add / edit / delete
// ======================================================================
var employees = [];            // { spItemId, id, last, first, businessUnit, laborCode, laborName, role }
var employeesLoadError = null; // set by upaeApplyEmployees when the SharePoint list could not be read
var empFilter = 'all';
var empSearch = '';
var empShown = 200;
var empPending = null;         // the current import preview, or null
var empBusy = false;

function empKey(id){ return String(id == null ? '' : id).trim().toUpperCase(); }
function empClean(s){ return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function empName(e){ return [e.last, e.first].filter(Boolean).join(', ') || e.id; }

// An employee's union is derived, never stored: default labor code -> LaborTypes -> union.
// Built-in default labor type "Employee": regular staff, no union. Anyone whose Role is one of
// these and who has no default labor type in the listing is treated as (and imported as) Employee.
var EMP_REGULAR_ROLES = ['field employee', 'project manager', 'field log', 'supervisor'];
function empRegularRole(role){
  var r = empClean(role).toLowerCase();
  if(!r) return false;
  return EMP_REGULAR_ROLES.some(function(x){ return r.indexOf(x) !== -1; });
}
function empUnionName(e){
  if(e.unionOverride) return e.unionOverride;
  if(!e.laborCode) return '';
  var lt = laborTypes.filter(function(x){ return String(x.code).toUpperCase() === String(e.laborCode).toUpperCase(); })[0];
  return lt && lt.unionName ? lt.unionName : '';
}
function employeeUnionInfo(e){
  if(e.unionOverride) return { label: e.unionOverride + ' (override)', cls: 'ok', kind: 'union' };
  if(String(e.laborCode || '').toUpperCase() === 'EMPLOYEE' || (!e.laborCode && !e.laborName && empRegularRole(e.role)))
    return { label: 'Employee (regular, no union)', cls: 'muted', kind: 'regular' };
  if(!e.laborCode) return { label: 'No default labor type', cls: 'warn', kind: 'nolabor' };
  var code = String(e.laborCode).toUpperCase();
  var lt = laborTypes.filter(function(x){ return String(x.code).toUpperCase() === code; })[0];
  if(!lt) return { label: 'Labor type not in list', cls: 'warn', kind: 'nounion' };
  if(!lt.unionName) return { label: 'No union assigned', cls: 'muted', kind: 'nounion' };
  return { label: lt.unionName, cls: 'ok', kind: 'union' };
}

function empWritesBlocked(){
  if(employeesLoadError) return 'The Employees list could not be read from SharePoint, so changes are switched off to avoid creating duplicates. See the notice at the top of the Employees screen.';
  return '';
}

function renderEmployeeKpis(){
  var el = document.getElementById('employeeKpis');
  if(!el) return;
  var withUnion = 0, noUnion = 0, noLabor = 0, regular = 0;
  employees.forEach(function(e){
    var k = employeeUnionInfo(e).kind;
    if(k === 'union') withUnion++; else if(k === 'regular') regular++; else if(k === 'nolabor') noLabor++; else noUnion++;
  });
  var bus = {};
  employees.forEach(function(e){ if(e.businessUnit) bus[e.businessUnit] = true; });
  el.innerHTML =
    '<div class="kpi"><div class="k-label">Employees on file</div><div class="k-value">' + employees.length + '</div>' +
      '<div class="k-delta flat">across ' + Object.keys(bus).length + ' business unit' + (Object.keys(bus).length === 1 ? '' : 's') + '</div></div>' +
    '<div class="kpi"><div class="k-label">With a union</div><div class="k-value">' + withUnion + '<span class="unit">/' + employees.length + '</span></div>' +
      '<div class="k-delta flat">from their default labor type</div></div>' +
    '<div class="kpi"><div class="k-label">Regular employees</div><div class="k-value">' + regular + '</div>' +
      '<div class="k-delta flat">default labor type Employee, no union</div></div>' +
    '<div class="kpi"><div class="k-label">No union / unknown</div><div class="k-value">' + noUnion + '</div>' +
      '<div class="k-delta ' + (noUnion ? 'up' : 'flat') + '">labor type has no union, or is not in the list</div></div>' +
    '<div class="kpi"><div class="k-label">No default labor type</div><div class="k-value">' + noLabor + '</div>' +
      '<div class="k-delta ' + (noLabor ? 'up' : 'flat') + '">blank in the Employee Listing</div></div>';
}

function renderEmployees(){
  renderEmployeeKpis();
  var banner = document.getElementById('employeesBanner');
  if(banner){
    if(employeesLoadError){
      banner.hidden = false;
      banner.innerHTML = '<b>The Employees list could not be read from SharePoint.</b> ' + escapeHtml(employeesLoadError) +
        '<br>If you have not created the list yet, follow "One-time setup" at the bottom of this screen, then reload and sign in again. Importing and editing are switched off until it loads.';
    } else { banner.hidden = true; banner.innerHTML = ''; }
  }
  var listEl = document.getElementById('employeeList');
  var countEl = document.getElementById('empCount');
  if(!listEl) return;

  var q = empSearch.trim().toLowerCase();
  var rows = employees.filter(function(e){
    var info = employeeUnionInfo(e);
    if(empFilter === 'union' && info.kind !== 'union') return false;
    if(empFilter === 'nounion' && info.kind !== 'nounion') return false;
    if(empFilter === 'regular' && info.kind !== 'regular') return false;
    if(empFilter === 'nolabor' && info.kind !== 'nolabor') return false;
    if(!q) return true;
    return (e.id + ' ' + empName(e) + ' ' + e.businessUnit + ' ' + e.laborName + ' ' + e.laborCode + ' ' + info.label).toLowerCase().indexOf(q) !== -1;
  });
  rows.sort(function(a, b){ return empName(a).localeCompare(empName(b)); });
  countEl.textContent = rows.length === employees.length ? (employees.length + ' employee' + (employees.length === 1 ? '' : 's'))
    : (rows.length + ' of ' + employees.length + ' shown');

  if(!employees.length){
    listEl.innerHTML = '<div class="admin-empty">No employees yet &mdash; choose the Employee Listing file above to load them.</div>';
    return;
  }
  if(!rows.length){ listEl.innerHTML = '<div class="admin-empty">No employees match.</div>'; return; }

  var html = '<div class="emp-row head"><div>ID</div><div>Name</div><div class="e-lab">Default labor type</div><div class="e-union">Union</div><div></div></div>';
  rows.slice(0, empShown).forEach(function(e){
    var info = employeeUnionInfo(e);
    var idx = employees.indexOf(e);
    html += '<div class="emp-row"><div class="e-id">' + escapeHtml(e.id) + '</div>' +
      '<div class="e-name">' + escapeHtml(empName(e)) + '<span class="e-sub">' + escapeHtml(e.businessUnit || '—') + '</span></div>' +
      '<div class="e-lab">' + escapeHtml(e.laborName ? e.laborName + (e.laborCode ? ' [' + e.laborCode + ']' : '') : (e.laborCode || '—')) + '</div>' +
      '<div class="e-union ' + info.cls + '">' + escapeHtml(info.label) + '</div>' +
      '<div class="e-btns"><button class="btn ghost" type="button" data-emp-edit="' + idx + '" style="padding:4px 10px; font-size:12px;">Edit</button>' +
      '<button class="a-del" type="button" data-emp-del="' + idx + '" title="Delete employee">&times;</button></div></div>';
  });
  if(rows.length > empShown){
    html += '<div class="emp-row" style="justify-content:center; grid-template-columns:1fr;"><button class="btn ghost" type="button" id="empShowMore" style="justify-self:center;">Show ' +
      Math.min(200, rows.length - empShown) + ' more (' + (rows.length - empShown) + ' not shown)</button></div>';
  }
  listEl.innerHTML = html;
  var more = document.getElementById('empShowMore');
  if(more) more.addEventListener('click', function(){ empShown += 200; renderEmployees(); });
  listEl.querySelectorAll('[data-emp-edit]').forEach(function(b){
    b.addEventListener('click', function(){ openEmployeeModal(parseInt(b.getAttribute('data-emp-edit'), 10)); });
  });
  listEl.querySelectorAll('[data-emp-del]').forEach(function(b){
    b.addEventListener('click', function(){ deleteEmployee(parseInt(b.getAttribute('data-emp-del'), 10)); });
  });
}

document.querySelectorAll('[data-emp-filter]').forEach(function(b){
  b.addEventListener('click', function(){
    empFilter = b.getAttribute('data-emp-filter'); empShown = 200;
    document.querySelectorAll('[data-emp-filter]').forEach(function(x){ x.classList.toggle('active', x === b); });
    renderEmployees();
  });
});
var empSearchEl = document.getElementById('empSearch');
if(empSearchEl) empSearchEl.addEventListener('input', function(){ empSearch = empSearchEl.value; empShown = 200; renderEmployees(); });

// ---- minimal .xlsx reader (no library): zip + shared strings + first sheet ----
async function empInflate(data){
  var ds = new DecompressionStream('deflate-raw');
  var stream = new Blob([data]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function empReadZip(buf){
  var u8 = new Uint8Array(buf), dv = new DataView(buf);
  var i = u8.length - 22;
  while(i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
  if(i < 0) throw new Error('This does not look like an .xlsx file. Please export the Employee Listing as .xlsx or .csv.');
  var count = dv.getUint16(i + 10, true), p = dv.getUint32(i + 16, true), entries = {};
  for(var n = 0; n < count; n++){
    if(dv.getUint32(p, true) !== 0x02014b50) break;
    var method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    var nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    var lho = dv.getUint32(p + 42, true);
    var name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen));
    entries[name] = { method: method, csize: csize, lho: lho };
    p += 46 + nlen + elen + clen;
  }
  async function read(name){
    var e = entries[name];
    if(!e) return null;
    var start = e.lho + 30 + dv.getUint16(e.lho + 26, true) + dv.getUint16(e.lho + 28, true);
    var data = u8.subarray(start, start + e.csize);
    var out = e.method === 0 ? data : await empInflate(data);
    return new TextDecoder().decode(out);
  }
  return { names: Object.keys(entries), read: read };
}
function empColIndex(ref){
  var m = /^([A-Z]+)/.exec(ref || ''), n = 0;
  if(!m) return 0;
  for(var i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
  return n - 1;
}
async function empParseXlsx(buf){
  var zip = await empReadZip(buf);
  var parser = new DOMParser();
  function xml(text){ return parser.parseFromString(text, 'application/xml'); }

  // which file is the first sheet?
  var sheetPath = 'xl/worksheets/sheet1.xml';
  var wbText = await zip.read('xl/workbook.xml');
  var relText = await zip.read('xl/_rels/workbook.xml.rels');
  if(wbText && relText){
    var sheetEl = xml(wbText).getElementsByTagName('sheet')[0];
    var rid = sheetEl && (sheetEl.getAttribute('r:id') || sheetEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
    var rels = xml(relText).getElementsByTagName('Relationship');
    for(var r = 0; r < rels.length; r++){
      if(rels[r].getAttribute('Id') === rid){
        var t = rels[r].getAttribute('Target') || '';
        sheetPath = t.charAt(0) === '/' ? t.slice(1) : 'xl/' + t;
        break;
      }
    }
  }
  var sheetText = await zip.read(sheetPath);
  if(!sheetText) throw new Error('Could not find the first worksheet in this file.');

  var shared = [];
  var ssText = await zip.read('xl/sharedStrings.xml');
  if(ssText){
    var sis = xml(ssText).getElementsByTagName('si');
    for(var k = 0; k < sis.length; k++){
      var out = '';
      for(var node = sis[k].firstChild; node; node = node.nextSibling){
        if(node.nodeName === 't') out += node.textContent;
        else if(node.nodeName === 'r'){
          var ts = node.getElementsByTagName('t');
          for(var m = 0; m < ts.length; m++) out += ts[m].textContent;
        }
      }
      shared.push(out);
    }
  }
  var rows = [];
  var rowEls = xml(sheetText).getElementsByTagName('row');
  for(var a = 0; a < rowEls.length; a++){
    var rn = parseInt(rowEls[a].getAttribute('r'), 10);
    var ri = isNaN(rn) ? rows.length : rn - 1;
    var cells = rowEls[a].getElementsByTagName('c'), arr = [];
    for(var b = 0; b < cells.length; b++){
      var c = cells[b], ci = empColIndex(c.getAttribute('r')), t2 = c.getAttribute('t');
      var vEl = c.getElementsByTagName('v')[0], val = '';
      if(t2 === 'inlineStr'){
        var isEl = c.getElementsByTagName('is')[0];
        val = isEl ? isEl.textContent : '';
      } else if(vEl){
        val = vEl.textContent;
        if(t2 === 's') val = shared[parseInt(val, 10)] || '';
        else if(t2 === 'e') val = '';
      }
      while(arr.length < ci) arr.push('');
      arr[ci] = val;
    }
    while(rows.length < ri) rows.push([]);
    rows[ri] = arr;
  }
  return rows;
}
function empParseCsv(text){
  var rows = [], row = [], cur = '', q = false;
  text = text.replace(/^﻿/, '');
  for(var i = 0; i < text.length; i++){
    var ch = text[i];
    if(q){
      if(ch === '"'){ if(text[i + 1] === '"'){ cur += '"'; i++; } else q = false; } else cur += ch;
    } else if(ch === '"') q = true;
    else if(ch === ','){ row.push(cur); cur = ''; }
    else if(ch === '\n' || ch === '\r'){
      if(ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur); rows.push(row); row = []; cur = '';
    } else cur += ch;
  }
  if(cur !== '' || row.length){ row.push(cur); rows.push(row); }
  return rows;
}

// ---- turn spreadsheet rows into employee records ----
function empSplitLabor(s){
  s = empClean(s);
  if(!s) return { code: '', name: '' };
  var m = s.match(/^(.*?)\s*\[([^\[\]]+)\]$/);
  return m ? { name: empClean(m[1]), code: empClean(m[2]) } : { name: s, code: '' };
}
function empParseListing(rows){
  var hi = -1, cols = null;
  for(var r = 0; r < Math.min(rows.length, 25); r++){
    var m = {};
    (rows[r] || []).forEach(function(c, i){ var h = empClean(c).toLowerCase(); if(h && m[h] == null) m[h] = i; });
    var idCol = m['id'] != null ? m['id'] : m['employee id'];
    if(idCol != null && m['last name'] != null){
      cols = { id: idCol, last: m['last name'], first: m['first name'], bu: m['business unit'], dl: m['default labor type'], role: m['role'] };
      hi = r; break;
    }
  }
  if(!cols) throw new Error('Could not find the header row. The file needs at least "ID" and "Last Name" columns (the B2W Employee Listing has Last Name, First Name, ID, Business Unit, Default Labor Type, Labor Type, Role).');
  function cell(row, i){ return i == null ? '' : empClean(row[i]); }
  var all = [];
  for(var j = hi + 1; j < rows.length; j++){
    var row = rows[j] || [];
    var id = cell(row, cols.id);
    if(!id) continue;
    var lab = empSplitLabor(cell(row, cols.dl));
    if(!lab.code && !lab.name && empRegularRole(cell(row, cols.role))) lab = { code: 'EMPLOYEE', name: 'Employee' };
    all.push({ id: id, last: cell(row, cols.last), first: cell(row, cols.first), businessUnit: cell(row, cols.bu),
               laborCode: lab.code, laborName: lab.name, role: cell(row, cols.role) });
  }
  // The same Employee ID can appear more than once (for example under two business units).
  // Keep one row per ID: prefer the row that has a default labor type, otherwise the first.
  var byKey = {}, order = [], skipped = [];
  all.forEach(function(rec){
    var k = empKey(rec.id);
    if(!byKey[k]){ byKey[k] = rec; order.push(k); return; }
    var keep = byKey[k];
    if(!keep.laborName && !keep.laborCode && (rec.laborName || rec.laborCode)){ skipped.push(keep); byKey[k] = rec; }
    else skipped.push(rec);
  });
  return { records: order.map(function(k){ return byKey[k]; }), skipped: skipped, rowsRead: all.length };
}

var EMP_FIELDS = [['last', 'Last name'], ['first', 'First name'], ['businessUnit', 'Business unit'],
                  ['laborCode', 'Labor code'], ['laborName', 'Labor type'], ['role', 'Role']];

function empBuildPreview(parsed, fileName){
  var existing = {};
  employees.forEach(function(e){ existing[empKey(e.id)] = e; });
  var fresh = [], changed = [], same = 0, seen = {};
  parsed.records.forEach(function(rec){
    var k = empKey(rec.id);
    seen[k] = true;
    var cur = existing[k];
    if(!cur){ fresh.push(rec); return; }
    var diffs = [];
    EMP_FIELDS.forEach(function(f){
      if(empClean(cur[f[0]]) !== empClean(rec[f[0]])) diffs.push({ label: f[1], from: cur[f[0]], to: rec[f[0]] });
    });
    if(diffs.length) changed.push({ cur: cur, next: rec, diffs: diffs }); else same++;
  });
  var notInFile = employees.filter(function(e){ return !seen[empKey(e.id)]; }).length;
  return { fileName: fileName, fresh: fresh, changed: changed, same: same, skipped: parsed.skipped, notInFile: notInFile, rowsRead: parsed.rowsRead };
}

function renderEmpPreview(){
  var el = document.getElementById('empImportPreview');
  if(!el) return;
  if(!empPending){ el.innerHTML = ''; return; }
  var p = empPending, total = p.fresh.length + p.changed.length;
  var html = '<div class="emp-sum">' +
    '<span class="emp-pill new"><b>' + p.fresh.length + '</b>new</span>' +
    '<span class="emp-pill chg"><b>' + p.changed.length + '</b>changed</span>' +
    '<span class="emp-pill"><b>' + p.same + '</b>unchanged</span>' +
    '<span class="emp-pill"><b>' + p.skipped.length + '</b>repeated IDs skipped</span>' +
    '<span class="emp-pill"><b>' + p.notInFile + '</b>on file, not in this listing (left alone)</span></div>';

  if(p.fresh.length){
    html += '<details class="emp-details"' + (p.fresh.length <= 25 ? ' open' : '') + '><summary>New employees (' + p.fresh.length + ')</summary><div class="emp-detail-body">';
    p.fresh.slice(0, 100).forEach(function(r){
      html += '<div><span class="mono">' + escapeHtml(r.id) + '</span> &nbsp;' + escapeHtml(empName(r)) + ' &middot; ' + escapeHtml(r.laborName || 'no default labor type') + '</div>';
    });
    if(p.fresh.length > 100) html += '<div>&hellip; and ' + (p.fresh.length - 100) + ' more</div>';
    html += '</div></details>';
  }
  if(p.changed.length){
    html += '<details class="emp-details" open><summary>Changed employees (' + p.changed.length + ')</summary><div class="emp-detail-body">';
    p.changed.slice(0, 150).forEach(function(c){
      html += '<div><span class="mono">' + escapeHtml(c.next.id) + '</span> &nbsp;' + escapeHtml(empName(c.cur)) + ': ' +
        c.diffs.map(function(d){ return escapeHtml(d.label) + ' <i>' + escapeHtml(d.from || '(blank)') + '</i> &rarr; <b>' + escapeHtml(d.to || '(blank)') + '</b>'; }).join('; ') + '</div>';
    });
    if(p.changed.length > 150) html += '<div>&hellip; and ' + (p.changed.length - 150) + ' more</div>';
    html += '</div></details>';
  }
  if(p.skipped.length){
    html += '<details class="emp-details"><summary>Repeated IDs skipped (' + p.skipped.length + ')</summary><div class="emp-detail-body">';
    p.skipped.forEach(function(r){
      html += '<div><span class="mono">' + escapeHtml(r.id) + '</span> &nbsp;' + escapeHtml(empName(r)) + ' &middot; ' + escapeHtml(r.businessUnit || 'no business unit') + ' &mdash; the other row for this ID was kept</div>';
    });
    html += '</div></details>';
  }
  var blocked = empWritesBlocked();
  html += '<div class="emp-actions"><button class="btn primary" type="button" id="empApplyBtn"' + ((!total || blocked || empBusy) ? ' disabled' : '') + '>' +
    (total ? 'Apply ' + total + ' change' + (total === 1 ? '' : 's') + ' to SharePoint' : 'Nothing to change') + '</button>' +
    '<button class="btn ghost" type="button" id="empCancelBtn"' + (empBusy ? ' disabled' : '') + '>Discard</button></div>';
  if(!total) html += '<div class="emp-status">The listing matches what is already on file (' + p.rowsRead + ' rows read).</div>';
  if(blocked) html += '<div class="emp-status err">' + escapeHtml(blocked) + '</div>';
  html += '<div class="emp-status" id="empApplyStatus"></div>';
  el.innerHTML = html;
  var apply = document.getElementById('empApplyBtn'), cancel = document.getElementById('empCancelBtn');
  if(apply) apply.addEventListener('click', empApplyPending);
  if(cancel) cancel.addEventListener('click', function(){
    empPending = null; document.getElementById('empFileName').textContent = 'No file chosen';
    document.getElementById('empFileInput').value = ''; renderEmpPreview();
  });
}

async function empApplyPending(){
  if(!empPending || empBusy) return;
  var blocked = empWritesBlocked();
  if(blocked){ alert(blocked); return; }
  var p = empPending, statusEl = document.getElementById('empApplyStatus');
  var ops = [];
  p.fresh.forEach(function(rec){ ops.push({ kind: 'create', employee: rec }); });
  p.changed.forEach(function(c){ ops.push({ kind: 'update', employee: Object.assign({}, c.next, { id: c.cur.id }), spItemId: c.cur.spItemId, cur: c.cur }); });
  if(!ops.length) return;
  empBusy = true;
  document.getElementById('empApplyBtn').disabled = true;
  document.getElementById('empCancelBtn').disabled = true;
  if(statusEl) statusEl.textContent = 'Saving 0 of ' + ops.length + '… keep this tab open.';
  var results;
  try{
    if(typeof upaeBatchEmployees === 'function'){
      results = await upaeBatchEmployees(ops, function(done, total){
        if(statusEl) statusEl.textContent = 'Saving ' + done + ' of ' + total + '… keep this tab open.';
      });
    } else {
      results = ops.map(function(){ return { ok: true }; }); // running without the SharePoint layer
    }
  } catch(err){
    console.error('UPAE: employee import failed', err);
    empBusy = false;
    if(statusEl){ statusEl.className = 'emp-status err'; statusEl.textContent = 'The import stopped: ' + err.message + '. Some rows may already have been saved; load the file again to see what is left.'; }
    var a1 = document.getElementById('empApplyBtn'), c1 = document.getElementById('empCancelBtn');
    if(a1) a1.disabled = false; if(c1) c1.disabled = false;
    return;
  }
  var added = 0, updated = 0, failed = [];
  ops.forEach(function(op, i){
    var res = results[i];
    if(!res || !res.ok){ failed.push({ id: op.employee.id, error: res && res.error ? res.error : 'Unknown error' }); return; }
    if(op.kind === 'create'){ employees.push(Object.assign({}, op.employee, { spItemId: res.spItemId })); added++; }
    else { EMP_FIELDS.forEach(function(f){ op.cur[f[0]] = op.employee[f[0]]; }); updated++; }
  });
  empBusy = false;
  empPending = null;
  document.getElementById('empFileInput').value = '';
  document.getElementById('empFileName').textContent = 'No file chosen';
  var el = document.getElementById('empImportPreview');
  el.innerHTML = '<div class="emp-sum"><span class="emp-pill new"><b>' + added + '</b>added</span><span class="emp-pill chg"><b>' + updated + '</b>updated</span>' +
    '<span class="emp-pill"' + (failed.length ? ' style="background:var(--crit-soft); color:var(--crit);"' : '') + '><b>' + failed.length + '</b>failed</span></div>' +
    (failed.length ? '<details class="emp-details" open><summary>Rows that could not be saved (' + failed.length + ')</summary><div class="emp-detail-body">' +
      failed.slice(0, 50).map(function(f){ return '<div><span class="mono">' + escapeHtml(f.id) + '</span> &mdash; ' + escapeHtml(f.error) + '</div>'; }).join('') +
      '</div></details><div class="emp-status">Load the same file again to retry just the rows that are still missing.</div>'
      : '<div class="emp-status">Done. Everything in the listing is now saved to SharePoint.</div>');
  empShown = 200;
  renderEmployees();
}

var empFileInput = document.getElementById('empFileInput');
if(empFileInput){
  empFileInput.addEventListener('change', async function(){
    var file = empFileInput.files && empFileInput.files[0];
    if(!file) return;
    document.getElementById('empFileName').textContent = file.name;
    var el = document.getElementById('empImportPreview');
    el.innerHTML = '<div class="emp-status">Reading ' + escapeHtml(file.name) + '…</div>';
    try{
      var rows;
      var buf = await file.arrayBuffer();
      empFileInput.value = ''; // so choosing the same file again (after fixing something) still fires 'change'
      var head = new Uint8Array(buf.slice(0, 2));
      if(head[0] === 0x50 && head[1] === 0x4B) rows = await empParseXlsx(buf);
      else if(/\.xls$/i.test(file.name)) throw new Error('Old .xls files are not supported. Please save it as .xlsx or .csv and try again.');
      else rows = empParseCsv(new TextDecoder().decode(buf));
      var parsed = empParseListing(rows);
      if(!parsed.records.length) throw new Error('No employee rows were found under the header.');
      empPending = empBuildPreview(parsed, file.name);
      renderEmpPreview();
    } catch(err){
      console.error('UPAE: could not read the Employee Listing', err);
      empPending = null;
      el.innerHTML = '<div class="emp-status err">' + escapeHtml(err.message || String(err)) + '</div>';
    }
  });
}

// ---- add / edit / delete one employee ----
var emOverlay = document.getElementById('employeeModalOverlay');
var emEditIdx = null;
function emFill(selectedCode, selectedOverride){
  var sel = document.getElementById('em_labor');
  var sorted = laborTypes.slice().sort(function(a, b){ return String(a.name).localeCompare(String(b.name)); });
  var html = '<option value="">— None —</option>';
  var found = false;
  var hasEmp = sorted.some(function(lt){ return String(lt.code).toUpperCase() === 'EMPLOYEE'; });
  if(!hasEmp){
    var empSel = selectedCode && String(selectedCode).toUpperCase() === 'EMPLOYEE';
    if(empSel) found = true;
    html += '<option value="EMPLOYEE"' + (empSel ? ' selected' : '') + '>Employee [EMPLOYEE] (regular, no union)</option>';
  }
  sorted.forEach(function(lt){
    var isSel = selectedCode && String(lt.code).toUpperCase() === String(selectedCode).toUpperCase();
    if(isSel) found = true;
    html += '<option value="' + escapeHtml(lt.code) + '"' + (isSel ? ' selected' : '') + '>' + escapeHtml(lt.name) + ' [' + escapeHtml(lt.code) + ']</option>';
  });
  if(selectedCode && !found) html += '<option value="' + escapeHtml(selectedCode) + '" selected>' + escapeHtml(selectedCode) + ' (not in labor type list)</option>';
  sel.innerHTML = html;
  var bus = {};
  employees.forEach(function(e){ if(e.businessUnit) bus[e.businessUnit] = true; });
  document.getElementById('em_buOptions').innerHTML = Object.keys(bus).sort().map(function(b){ return '<option value="' + escapeHtml(b) + '">'; }).join('');
  var uo = document.getElementById('em_unionOverride');
  uo.innerHTML = '<option value="">— Use the labor type —</option>' + unions.slice().sort(function(a, b){ return String(a.name).localeCompare(String(b.name)); })
    .map(function(u){ return '<option value="' + escapeHtml(u.name) + '"' + (selectedOverride && selectedOverride === u.name ? ' selected' : '') + '>' + escapeHtml(u.name) + '</option>'; }).join('') +
    (selectedOverride && !unions.some(function(u){ return u.name === selectedOverride; }) ? '<option value="' + escapeHtml(selectedOverride) + '" selected>' + escapeHtml(selectedOverride) + ' (not in union list)</option>' : '');
  emUnionHint();
}
async function emFillPM(cur){
  var sel = document.getElementById('em_pm'), list = (typeof window.hrsGetPMs === 'function') ? window.hrsGetPMs() : [];
  if(!list.length && typeof upaeFetchProjectManagers === 'function'){ try{ list = await upaeFetchProjectManagers(); }catch(e){ list = []; } }
  var h = '<option value="">— None —</option>' + list.slice().sort(function(a, b){ return a.name.localeCompare(b.name); }).map(function(p){ return '<option value="' + escapeHtml(p.email) + '"' + (cur && cur.toLowerCase() === p.email.toLowerCase() ? ' selected' : '') + '>' + escapeHtml(p.name) + '</option>'; }).join('');
  if(cur && !list.some(function(p){ return p.email.toLowerCase() === cur.toLowerCase(); })) h += '<option value="' + escapeHtml(cur) + '" selected>' + escapeHtml(cur) + '</option>';
  sel.innerHTML = h;
}
function emUnionHint(){
  var code = document.getElementById('em_labor').value;
  var info = employeeUnionInfo({ laborCode: code });
  document.getElementById('em_unionHint').textContent = 'Union: ' + info.label;
}
document.getElementById('em_labor').addEventListener('change', emUnionHint);

function openEmployeeModal(idx){
  var blocked = empWritesBlocked();
  if(blocked){ alert(blocked); return; }
  var e = idx == null ? null : employees[idx];
  emEditIdx = e ? idx : null;
  document.getElementById('employeeModalTitle').textContent = e ? 'Edit employee' : 'Add employee';
  var idInput = document.getElementById('em_id');
  idInput.value = e ? e.id : ''; idInput.disabled = !!e;
  document.getElementById('em_last').value = e ? e.last : '';
  document.getElementById('em_first').value = e ? e.first : '';
  document.getElementById('em_bu').value = e ? e.businessUnit : '';
  document.getElementById('em_role').value = e ? e.role : '';
  emFill(e ? e.laborCode : '', e ? e.unionOverride : '');
  emFillPM(e ? e.pmEmail : '');
  emOverlay.hidden = false;
  (e ? document.getElementById('em_last') : idInput).focus();
}
function closeEmployeeModal(){ emOverlay.hidden = true; emEditIdx = null; }
document.getElementById('closeEmployeeModal').addEventListener('click', closeEmployeeModal);
document.getElementById('cancelEmployee').addEventListener('click', closeEmployeeModal);
// No click-outside-to-close, same as the other modals: a stray click must not lose typed data.
document.getElementById('addEmployeeBtn').addEventListener('click', function(){ openEmployeeModal(null); });

document.getElementById('saveEmployee').addEventListener('click', async function(){
  var btn = document.getElementById('saveEmployee');
  var id = empClean(document.getElementById('em_id').value);
  if(!id){ document.getElementById('em_id').focus(); return; }
  var code = document.getElementById('em_labor').value;
  var lt = laborTypes.filter(function(x){ return String(x.code).toUpperCase() === String(code).toUpperCase(); })[0];
  var existing = emEditIdx == null ? null : employees[emEditIdx];
  var rec = {
    id: id,
    last: empClean(document.getElementById('em_last').value),
    first: empClean(document.getElementById('em_first').value),
    businessUnit: empClean(document.getElementById('em_bu').value),
    laborCode: code,
    laborName: lt ? lt.name : (String(code).toUpperCase() === 'EMPLOYEE' ? 'Employee' : (existing && existing.laborCode === code ? existing.laborName : '')),
    role: empClean(document.getElementById('em_role').value),
    unionOverride: document.getElementById('em_unionOverride').value,
    pmEmail: document.getElementById('em_pm').value
  };
  if(existing && existing.pmEmail && !rec.pmEmail) rec.clearPm = true;
  if(existing && existing.unionOverride && !rec.unionOverride) rec.clearOverride = true;
  if(!existing){
    var dup = employees.filter(function(x){ return empKey(x.id) === empKey(id); })[0];
    if(dup){ alert('Employee ID ' + id + ' is already on file (' + empName(dup) + '). Use Edit on that row instead.'); return; }
  }
  btn.disabled = true; btn.textContent = 'Saving...';
  try{
    if(existing){
      if(typeof upaeUpdateEmployee === 'function' && existing.spItemId) await upaeUpdateEmployee(existing.spItemId, rec);
      EMP_FIELDS.forEach(function(f){ existing[f[0]] = rec[f[0]]; });
      existing.unionOverride = rec.unionOverride; existing.pmEmail = rec.pmEmail;
    } else {
      var created = (typeof upaeCreateEmployee === 'function') ? await upaeCreateEmployee(rec) : null;
      employees.push(Object.assign({}, rec, { spItemId: created && created.id ? created.id : null }));
    }
    renderEmployees();
    closeEmployeeModal();
  } catch(err){
    console.error('UPAE: failed to save employee', err);
    alert('Could not save this employee to SharePoint: ' + err.message + '. Nothing was changed.');
  }
  btn.disabled = false; btn.textContent = 'Save';
});

async function deleteEmployee(idx){
  var e = employees[idx];
  if(!e) return;
  var blocked = empWritesBlocked();
  if(blocked){ alert(blocked); return; }
  if(!window.confirm('Delete ' + empName(e) + ' (' + e.id + ') from the Employees list? This cannot be undone.')) return;
  try{
    if(typeof upaeDeleteEmployee === 'function' && e.spItemId) await upaeDeleteEmployee(e.spItemId);
    employees.splice(idx, 1);
    renderEmployees();
  } catch(err){
    console.error('UPAE: failed to delete employee', err);
    alert('Could not delete this employee from SharePoint: ' + err.message + '. It has not been removed.');
  }
}

// Bridge called by upaeLoadAllData() once the Employees list has been read.
// list = array, or null if the read failed; err = message when it failed.
window.upaeApplyEmployees = function(list, err){
  employees.splice(0, employees.length);
  if(list) Array.prototype.push.apply(employees, list);
  employeesLoadError = err || null;
  empPending = null; // any earlier preview was built against the old list
  renderEmpPreview();
  renderEmployees();
};

