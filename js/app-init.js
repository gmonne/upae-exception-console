renderUnionAdminList();
renderLaborAdminList();
populateLaborUnionSelect();

// Expose so upaeLoadAllData() (sharepoint-integration.js) can re-render the
// Admin screen once live Unions/LaborTypes data has been fetched.
window.upaeRenderAdminScreen = function(){
  renderUnionAdminList();
  renderLaborAdminList();
  populateLaborUnionSelect();
};

// rules/unions/laborTypes above are local to this IIFE, not on window — so
// sharepoint-integration.js can't reassign them directly. This bridge lets
// it swap in live SharePoint data by mutating the SAME array objects the
// render functions already close over (splice + push, not reassignment),
// then re-renders both screens.
window.upaeApplyLiveData = function(liveUnions, liveLaborTypes, liveRules){
  unions.splice(0, unions.length);
  Array.prototype.push.apply(unions, liveUnions);
  laborTypes.splice(0, laborTypes.length);
  Array.prototype.push.apply(laborTypes, liveLaborTypes);
  rules.splice(0, rules.length);
  Array.prototype.push.apply(rules, liveRules);
  if (window.upaeRenderAdminScreen) window.upaeRenderAdminScreen();
  if (window.upaeRenderRuleLibrary) window.upaeRenderRuleLibrary();
};

function populateUnionOptions(){
  var dl = document.getElementById('rf_unionOptions');
  if(!dl) return;
  dl.innerHTML = unions.map(function(u){ return '<option value="' + u.name + '">'; }).join('');
}
populateUnionOptions();

var listEl = document.getElementById('ruleList');
var detailEl = document.getElementById('ruleDetail');
var examplesEl = document.getElementById('ruleExamples');
if(!listEl) console.warn('UPAE: rule list element missing');

function renderList(selectedId){
  renderCoverageInsights();
  renderRuleLibraryKpis();
  listEl.innerHTML = '';
  rules.forEach(function(r){
    var row = document.createElement('div');
    row.className = 'rule-row' + (r.id === selectedId ? ' selected' : '');
    row.innerHTML =
      '<div class="r-main"><div class="r-name">' + r.name + '</div><div class="r-meta">' + r.id + ' &middot; ' + r.codes + '</div></div>' +
      '<span class="rule-status ' + r.status + '">' + r.statusLabel + '</span>';
    row.addEventListener('click', function(){ selectRule(r.id); });
    listEl.appendChild(row);
  });
}

var currentId = null;

function renderDetail(r){
  currentId = r.id;
  var html = '<div class="rule-detail-head"><div><div class="rule-detail-id">' + r.id + '</div>' +
    '<div class="rule-detail-title">' + r.name + '</div>' +
    '<div class="rule-detail-cba">' + r.cba + ' &middot; Last verified: ' + r.verified + '</div></div>' +
    '<div style="display:flex; flex-direction:column; align-items:flex-end; gap:8px;">' +
    '<span class="rule-status ' + r.status + '">' + r.statusLabel + '</span>' +
    '<div style="display:flex; gap:6px;">' +
    '<button class="btn ghost" type="button" id="duplicateRuleBtn" style="padding:6px 11px; font-size:12px;" title="Copy this rule\'s fields into a new rule for a different union">Duplicate</button>' +
    '<button class="btn ghost" type="button" id="editRuleBtn" style="padding:6px 11px; font-size:12px;">Edit rule</button>' +
    '</div>' +
    '</div></div>';

  html += '<div class="field-block"><div class="field-block-title">Labor type mapping</div>' +
    '<div style="font-size:13px; line-height:1.55;">' + r.codes + '</div>';
  if(r.status === 'ambiguous'){
    html += '<button class="btn primary" type="button" id="resolveMappingBtn" style="margin-top:10px; padding:7px 12px; font-size:12.5px;">Resolve mapping &mdash; confirm this labor type assignment</button>';
  }
  html += '</div>';

  if(!r.fields){
    html += '<div class="form-note" style="margin-top:16px;">Not yet structured against the Rule Definition Schema. Currently exists as a prose CBA summary; being rewritten after Payroll confirms the employee&rarr;union crosswalk and related open items. Click "Edit rule" to fill it in.</div>';
    detailEl.innerHTML = html;
    renderExamples(r);
    wireEditButton(r);
    wireDuplicateButton(r);
    wireResolveButton(r);
    return;
  }

  var groups = [
    ['Work week & regular time', ['Work week start day','Standard workday (hrs)','Regular threshold','Regular applicable days']],
    ['Overtime 1.5x', ['OT 1.5 trigger type','OT 1.5 threshold','OT 1.5 applicable days','OT 1.5 rate mult.','OT 1.5 benefit mult.']],
    ['Overtime 2x', ['OT 2x trigger type','OT 2x threshold','OT 2x applicable days','OT 2x rate mult.','OT 2x benefit mult.']],
    ['Night differential', ['Night trigger type','Night trigger condition','Night window start','Night window end','Night applicability',
      'Night tier 1 threshold','Night tier 1 rate mult.','Night tier 1 benefit mult.','Night tier 2 threshold','Night tier 2 rate mult.','Night tier 2 benefit mult.']]
  ];
  groups.forEach(function(g){
    html += '<div class="field-block"><div class="field-block-title">' + g[0] + '</div><div class="field-block-grid">';
    g[1].forEach(function(k){
      html += '<div class="detail-field"><div class="f-label">' + k + '</div><div class="f-value">' + (r.fields[k] || '—') + '</div></div>';
    });
    html += '</div></div>';
  });
  html += '<div class="field-block"><div class="field-block-title">Precedence rule</div><div style="font-size:13px; line-height:1.55;">' + (r.fields['Precedence rule'] || '—') + '</div></div>';
  detailEl.innerHTML = html;
  renderExamples(r);
  wireEditButton(r);
  wireDuplicateButton(r);
  wireResolveButton(r);
}

function wireEditButton(r){
  var btn = document.getElementById('editRuleBtn');
  if(btn) btn.addEventListener('click', function(){ openRuleModal(r); });
}

function wireDuplicateButton(r){
  var btn = document.getElementById('duplicateRuleBtn');
  if(btn) btn.addEventListener('click', function(){ openDuplicateModal(r); });
}

function wireResolveButton(r){
  var btn = document.getElementById('resolveMappingBtn');
  if(!btn) return;
  btn.addEventListener('click', function(){
    r.codes = r.codes.replace(/\s*—\s*(needs confirm|needs Payroll confirmation|pending confirm)\.?$/i, '') + ' — confirmed';
    r.status = r.fields ? 'structured' : 'draft';
    r.statusLabel = r.fields ? 'Structured' : 'Draft — prose only';
    renderList(r.id);
    renderDetail(r);
  });
}


function renderExamples(r){
  var exHtml = '<div style="display:flex; justify-content:flex-end; margin-bottom:4px;"><button class="btn ghost" type="button" id="addExampleBtn" style="padding:6px 11px; font-size:12px;">+ Add example</button></div>';
  if(r.examples && Object.keys(r.examples).length){
    Object.keys(r.examples).forEach(function(cat){
      exHtml += '<div class="example-category">' + escapeHtml(cat) + '</div><div class="example-grid">';
      r.examples[cat].forEach(function(ex, idx){
        exHtml += '<div class="example-card ' + (ex.ok ? 'good' : 'bad') + '">' +
          '<div class="ex-head" style="display:flex; align-items:center; justify-content:space-between; gap:8px;">' +
            '<span>' + (ex.ok ? '✓ Correct' : '✕ Incorrect') + '</span>' +
            '<span style="display:flex; gap:4px;">' +
              '<button type="button" class="btn ghost ex-edit-btn" data-cat="' + escapeHtml(cat) + '" data-idx="' + idx + '" style="padding:2px 8px; font-size:11px;" title="Edit this example">Edit</button>' +
              '<button type="button" class="btn ghost ex-delete-btn" data-cat="' + escapeHtml(cat) + '" data-idx="' + idx + '" style="padding:2px 8px; font-size:11px;" title="Delete this example">Delete</button>' +
            '</span>' +
          '</div><p>' + escapeHtml(ex.text) + '</p></div>';
      });
      exHtml += '</div>';
    });
  } else {
    exHtml += '<div class="form-note">No examples on file yet for this union.</div>';
  }
  examplesEl.innerHTML = exHtml;
  var addBtn = document.getElementById('addExampleBtn');
  if(addBtn) addBtn.addEventListener('click', function(){ openExampleModal(r); });
  examplesEl.querySelectorAll('.ex-edit-btn').forEach(function(btn){
    btn.addEventListener('click', function(){
      var cat = btn.getAttribute('data-cat');
      var idx = parseInt(btn.getAttribute('data-idx'), 10);
      openExampleModal(r, { category: cat, index: idx });
    });
  });
  examplesEl.querySelectorAll('.ex-delete-btn').forEach(function(btn){
    btn.addEventListener('click', function(){
      var cat = btn.getAttribute('data-cat');
      var idx = parseInt(btn.getAttribute('data-idx'), 10);
      deleteExample(r, cat, idx);
    });
  });
}

async function deleteExample(r, category, index){
  var ex = (r.examples && r.examples[category]) ? r.examples[category][index] : null;
  if(!ex) return;
  if(!window.confirm('Delete this ' + (ex.ok ? 'correct' : 'incorrect') + ' example? This cannot be undone.')) return;
  if(ex.spItemId && typeof upaeDeleteRuleExample === 'function'){
    try{ await upaeDeleteRuleExample(ex.spItemId); }
    catch(err){
      console.error('UPAE: failed to delete example from SharePoint', err);
      alert('Could not delete this example from SharePoint: ' + err.message + '. It has not been removed.');
      return;
    }
  }
  r.examples[category].splice(index, 1);
  if(!r.examples[category].length) delete r.examples[category];
  if(r.id === currentId) renderDetail(r);
}

function selectRule(id){
  var r = rules.filter(function(x){ return x.id === id; })[0];
  if(!r) return;
  renderList(id);
  renderDetail(r);
}

if (rules.length) {
  renderList(rules[0].id);
  renderDetail(rules[0]);
} else {
  renderList();
  detailEl.innerHTML = '<div class="admin-empty">Sign in to load rules from SharePoint.</div>';
}

// Expose so upaeLoadAllData() (see sharepoint-integration.js wiring below)
// can re-render this view once live data has been fetched.
window.upaeRenderRuleLibrary = function(){
  renderList(rules.length ? rules[0].id : undefined);
  if (rules.length) renderDetail(rules[0]);
  populateUnionOptions();
};

// ---- Add / Edit rule modal ----
var addOverlay = document.getElementById('addRuleModalOverlay');
var addOpenBtn = document.getElementById('openAddRuleModal');
var addCloseBtn = document.getElementById('closeAddRuleModal');
var addCancelBtn = document.getElementById('cancelAddRule');
var addSaveBtn = document.getElementById('saveAddRule');
var addModalTitle = document.getElementById('addRuleModalTitle');
var addModalNote = document.getElementById('ruleFormNote');
var editingId = null; // null = adding a new rule; otherwise the rule id being edited
var duplicateSourceExamples = null; // set by openDuplicateModal; examples to copy onto the new rule on save

var formIds = {
  name:'rf_name', cba:'rf_cba', effdate:'rf_effdate', verified:'rf_verified',
  weekstart:'rf_weekstart', workday:'rf_workday', regthreshold:'rf_regthreshold', regdays:'rf_regdays',
  ot15trigger:'rf_ot15trigger', ot15threshold:'rf_ot15threshold', ot15days:'rf_ot15days', ot15rate:'rf_ot15rate', ot15benefit:'rf_ot15benefit',
  ot2trigger:'rf_ot2trigger', ot2threshold:'rf_ot2threshold', ot2days:'rf_ot2days', ot2rate:'rf_ot2rate', ot2benefit:'rf_ot2benefit',
  nighttrigger:'rf_nighttrigger', nightcondition:'rf_nightcondition', nightstart:'rf_nightstart', nightend:'rf_nightend', nightapplic:'rf_nightapplic',
  nighttier1threshold:'rf_nighttier1threshold', nighttier1rate:'rf_nighttier1rate', nighttier1benefit:'rf_nighttier1benefit',
  nighttier2threshold:'rf_nighttier2threshold', nighttier2rate:'rf_nighttier2rate', nighttier2benefit:'rf_nighttier2benefit',
  precedence:'rf_precedence', prose:'rf_prose'
};
function fv(id){ var el = document.getElementById(id); return el ? el.value : ''; }
function setfv(id, val){ var el = document.getElementById(id); if(el) el.value = val || ''; }

function applyNightTriggerVisibility(){
  var sel = document.getElementById('rf_nighttrigger');
  var startRow = document.getElementById('rf_nightwindow_start_row');
  var endRow = document.getElementById('rf_nightwindow_end_row');
  var condRow = document.getElementById('rf_nightcondition_row');
  if(!sel) return;
  var isClock = sel.value === 'Clock-time window';
  var isRecall = sel.value === 'Send-home / Recall event';
  if(startRow) startRow.style.display = isClock ? '' : 'none';
  if(endRow) endRow.style.display = isClock ? '' : 'none';
  if(condRow) condRow.style.display = isRecall ? '' : 'none';
}
var rf_nighttriggerSel = document.getElementById('rf_nighttrigger');
if(rf_nighttriggerSel) rf_nighttriggerSel.addEventListener('change', applyNightTriggerVisibility);
applyNightTriggerVisibility();

var naSwitchIds = ['rf_ot15_na', 'rf_ot2_na', 'rf_night_na'];
function wireNAtoggle(switchId, blockId, triggerId){
  var cb = document.getElementById(switchId);
  var block = document.getElementById(blockId);
  if(!cb || !block) return;
  var apply = function(){
    var na = cb.checked;
    block.classList.toggle('is-na', na);
    var inputs = block.querySelectorAll('.rule-form-grid input, .rule-form-grid select, .rule-form-grid textarea');
    inputs.forEach(function(el){ el.disabled = na; });
    if(triggerId){
      var sel = document.getElementById(triggerId);
      if(sel){
        if(na){ sel.value = 'Not Applicable'; }
        else if(sel.value === 'Not Applicable'){ sel.selectedIndex = 0; }
      }
    }
    if(blockId === 'fb_night') applyNightTriggerVisibility();
  };
  cb.addEventListener('change', apply);
  cb._applyNA = apply;
}
wireNAtoggle('rf_ot15_na', 'fb_ot15', 'rf_ot15trigger');
wireNAtoggle('rf_ot2_na', 'fb_ot2', 'rf_ot2trigger');
wireNAtoggle('rf_night_na', 'fb_night', 'rf_nighttrigger');

function clearForm(){
  Object.keys(formIds).forEach(function(k){ setfv(formIds[k], ''); });
  naSwitchIds.forEach(function(id){
    var cb = document.getElementById(id);
    if(cb){ cb.checked = false; if(cb._applyNA) cb._applyNA(); }
  });
}

function applyCodesUnionFilter(){
  var filterSel = document.getElementById('rf_codesUnionFilter');
  var listEl = document.getElementById('rf_codesList');
  if(!filterSel || !listEl) return;
  var val = filterSel.value; // '' = all, 'UNASSIGNED' = unassigned, else a union name
  var anyVisible = false;
  listEl.querySelectorAll('.code-row').forEach(function(row){
    var u = row.getAttribute('data-union') || '';
    var match = (val === '') || (val === 'UNASSIGNED' ? u === '' : u === val);
    row.style.display = match ? 'flex' : 'none';
    if(match) anyVisible = true;
  });
  var emptyEl = document.getElementById('rf_codesEmpty');
  if(emptyEl) emptyEl.style.display = anyVisible ? 'none' : 'block';
}

function renderCodesMultiselect(r){
  var listEl = document.getElementById('rf_codesList');
  var summaryEl = document.getElementById('rf_codesSummary');
  var filterSel = document.getElementById('rf_codesUnionFilter');
  if(!listEl) return;

  // Union filter dropdown: defaults to the rule's own union so you see its codes first,
  // but "All unions" is one click away.
  if(filterSel){
    var defaultVal = r ? r.name : '';
    filterSel.innerHTML = '<option value="">All unions</option>' +
      unions.map(function(u){ return '<option value="' + u.name + '"' + (u.name === defaultVal ? ' selected' : '') + '>' + u.name + '</option>'; }).join('') +
      '<option value="UNASSIGNED">Unassigned codes</option>';
    filterSel.onchange = applyCodesUnionFilter;
  }

  if(!laborTypes.length){
    listEl.innerHTML = '<div class="admin-empty">No labor type codes yet — add some on the Admin screen.</div>';
    summaryEl.textContent = '';
    return;
  }
  var picked = (r && r.laborTypeIds) || [];
  listEl.innerHTML = laborTypes.map(function(lt){
    var checked = picked.indexOf(lt.code) > -1 ? ' checked' : '';
    return '<label class="code-row" data-union="' + (lt.unionName || '') + '"><input type="checkbox" value="' + lt.code + '"' + checked + '>' +
      '<span class="code-id">' + lt.code + '</span><span class="code-name">' + lt.name + '</span>' +
      '<span class="code-union">' + (lt.unionName || 'Unassigned') + '</span></label>';
  }).join('') + '<div class="admin-empty" id="rf_codesEmpty" style="display:none;">No codes assigned to this union yet &mdash; switch to "All unions" or assign some on the Admin screen.</div>';
  listEl.querySelectorAll('input[type="checkbox"]').forEach(function(cb){
    cb.addEventListener('change', updateCodesSummary);
  });
  applyCodesUnionFilter();
  updateCodesSummary();
}
function getSelectedCodes(){
  var boxes = document.querySelectorAll('#rf_codesList input[type="checkbox"]:checked');
  var out = [];
  boxes.forEach ? boxes.forEach(function(b){ out.push(b.value); }) : (function(){ for(var i=0;i<boxes.length;i++) out.push(boxes[i].value); })();
  return out;
}
function updateCodesSummary(){
  var summaryEl = document.getElementById('rf_codesSummary');
  if(!summaryEl) return;
  var codes = getSelectedCodes();
  summaryEl.textContent = codes.length ? (codes.length + ' code' + (codes.length>1?'s':'') + ' selected: ' + codes.join(', ')) : 'No codes selected — managed on the Admin screen.';
}

function fillFormFromRule(r){
  clearForm();
  setfv(formIds.name, r.name === 'New union' ? '' : r.name);
  setfv(formIds.cba, r.cba === '—' ? '' : r.cba);
  setfv(formIds.verified, r.verified === 'Not yet verified' ? '' : r.verified);
  renderCodesMultiselect(r);
  var f = r.fields || {};
  setfv(formIds.weekstart, f['Work week start day']);
  setfv(formIds.workday, f['Standard workday (hrs)']);
  setfv(formIds.regthreshold, f['Regular threshold']);
  setfv(formIds.regdays, f['Regular applicable days']);
  setfv(formIds.ot15trigger, f['OT 1.5 trigger type']);
  setfv(formIds.ot15threshold, f['OT 1.5 threshold']);
  setfv(formIds.ot15days, f['OT 1.5 applicable days']);
  setfv(formIds.ot15rate, f['OT 1.5 rate mult.']);
  setfv(formIds.ot15benefit, f['OT 1.5 benefit mult.']);
  setfv(formIds.ot2trigger, f['OT 2x trigger type']);
  setfv(formIds.ot2threshold, f['OT 2x threshold']);
  setfv(formIds.ot2days, f['OT 2x applicable days']);
  setfv(formIds.ot2rate, f['OT 2x rate mult.']);
  setfv(formIds.ot2benefit, f['OT 2x benefit mult.']);
  setfv(formIds.nighttrigger, f['Night trigger type'] || 'Clock-time window');
  setfv(formIds.nightcondition, f['Night trigger condition']);
  setfv(formIds.nightstart, f['Night window start']);
  setfv(formIds.nightend, f['Night window end']);
  setfv(formIds.nightapplic, f['Night applicability']);
  setfv(formIds.nighttier1threshold, f['Night tier 1 threshold']);
  setfv(formIds.nighttier1rate, f['Night tier 1 rate mult.']);
  setfv(formIds.nighttier1benefit, f['Night tier 1 benefit mult.']);
  setfv(formIds.nighttier2threshold, f['Night tier 2 threshold']);
  setfv(formIds.nighttier2rate, f['Night tier 2 rate mult.']);
  setfv(formIds.nighttier2benefit, f['Night tier 2 benefit mult.']);
  setfv(formIds.precedence, f['Precedence rule']);
  setfv(formIds.prose, r.prose);
  applyNightTriggerVisibility();

  var naMap = {
    rf_ot15_na: f['OT 1.5 trigger type'] === 'Not Applicable',
    rf_ot2_na: f['OT 2x trigger type'] === 'Not Applicable',
    rf_night_na: f['Night trigger type'] === 'Not Applicable'
  };
  Object.keys(naMap).forEach(function(id){
    var cb = document.getElementById(id);
    if(cb){ cb.checked = naMap[id]; if(cb._applyNA) cb._applyNA(); }
  });
}

function openRuleModal(r){
  editingId = r ? r.id : null;
  duplicateSourceExamples = null;
  var liveNote = (typeof upaeCreateRule === 'function') ? 'Saving writes this rule to the shared SharePoint Rules list.' : 'Nothing leaves your browser — this updates the in-memory list for the demo; refreshing the page resets it.';
  if(r){
    addModalTitle.textContent = 'Edit union rule — ' + r.name;
    addModalNote.textContent = 'Editing ' + r.id + '. ' + liveNote;
    fillFormFromRule(r);
  } else {
    addModalTitle.textContent = 'Add union rule';
    addModalNote.textContent = liveNote;
    clearForm();
    renderCodesMultiselect(null);
  }
  addOverlay.hidden = false;
}

function openDuplicateModal(sourceRule){
  editingId = null; // duplicating always creates a new rule, never overwrites the source
  // Deep-copy the source rule's examples so they can be re-saved under the new
  // rule's name/id on Save — copied here rather than referenced, since the
  // Add Example modal mutates r.examples in place for whichever rule is open.
  duplicateSourceExamples = JSON.parse(JSON.stringify(sourceRule.examples || {}));
  var exCount = Object.keys(duplicateSourceExamples).reduce(function(n, cat){ return n + duplicateSourceExamples[cat].length; }, 0);
  var liveNote = (typeof upaeCreateRule === 'function') ? 'Saving writes this rule to the shared SharePoint Rules list.' : 'Nothing leaves your browser — this updates the in-memory list for the demo; refreshing the page resets it.';
  addModalTitle.textContent = 'Add union rule — duplicated from ' + sourceRule.name;
  addModalNote.textContent = 'All fields copied from ' + sourceRule.name + ' except the union name and labor type codes — set those two, double-check the numbers still apply, then save.' +
    (exCount ? ' ' + exCount + ' worked example' + (exCount > 1 ? 's' : '') + ' will be copied too — review them for the new union before saving.' : '') + ' ' + liveNote;
  fillFormFromRule(sourceRule);
  setfv(formIds.name, '');
  setfv(formIds.cba, '');
  setfv(formIds.verified, '');
  renderCodesMultiselect(null); // codes are union-specific — start with none picked
  addOverlay.hidden = false;
  var nameEl = document.getElementById(formIds.name);
  if(nameEl) nameEl.focus();
}
function closeAddModal(){ addOverlay.hidden = true; duplicateSourceExamples = null; }

if(addOpenBtn){
  addOpenBtn.addEventListener('click', function(){ openRuleModal(null); });
  addCloseBtn.addEventListener('click', closeAddModal);
  addCancelBtn.addEventListener('click', closeAddModal);
  // Deliberately no click-outside-to-close on this overlay — an accidental
  // click while filling out a rule used to silently discard everything
  // that had been entered. Closing now only happens via Save or Cancel/X.

  function laborTypeNameForCode(code){
    var lt = laborTypes.filter(function(x){ return x.code === code; })[0];
    return lt ? lt.name : code;
  }

  addSaveBtn.addEventListener('click', async function(){
    var name = fv(formIds.name).trim();
    if(!name){ document.getElementById(formIds.name).focus(); return; }

    var ot15NA = !!(document.getElementById('rf_ot15_na') && document.getElementById('rf_ot15_na').checked);
    var ot2NA = !!(document.getElementById('rf_ot2_na') && document.getElementById('rf_ot2_na').checked);
    var nightNA = !!(document.getElementById('rf_night_na') && document.getElementById('rf_night_na').checked);

    var fields = {
      'Work week start day': fv(formIds.weekstart), 'Standard workday (hrs)': fv(formIds.workday),
      'Regular threshold': fv(formIds.regthreshold), 'Regular applicable days': fv(formIds.regdays),
      'OT 1.5 trigger type': ot15NA ? 'Not Applicable' : fv(formIds.ot15trigger),
      'OT 1.5 threshold': ot15NA ? '' : fv(formIds.ot15threshold),
      'OT 1.5 applicable days': ot15NA ? '' : fv(formIds.ot15days),
      'OT 1.5 rate mult.': ot15NA ? '' : fv(formIds.ot15rate),
      'OT 1.5 benefit mult.': ot15NA ? '' : fv(formIds.ot15benefit),
      'OT 2x trigger type': ot2NA ? 'Not Applicable' : fv(formIds.ot2trigger),
      'OT 2x threshold': ot2NA ? '' : fv(formIds.ot2threshold),
      'OT 2x applicable days': ot2NA ? '' : fv(formIds.ot2days),
      'OT 2x rate mult.': ot2NA ? '' : fv(formIds.ot2rate),
      'OT 2x benefit mult.': ot2NA ? '' : fv(formIds.ot2benefit),
      'Night trigger type': nightNA ? 'Not Applicable' : fv(formIds.nighttrigger),
      'Night trigger condition': nightNA ? '' : fv(formIds.nightcondition),
      'Night window start': nightNA ? '' : fv(formIds.nightstart),
      'Night window end': nightNA ? '' : fv(formIds.nightend),
      'Night applicability': nightNA ? '' : fv(formIds.nightapplic),
      'Night tier 1 threshold': nightNA ? '' : fv(formIds.nighttier1threshold),
      'Night tier 1 rate mult.': nightNA ? '' : fv(formIds.nighttier1rate),
      'Night tier 1 benefit mult.': nightNA ? '' : fv(formIds.nighttier1benefit),
      'Night tier 2 threshold': nightNA ? '' : fv(formIds.nighttier2threshold),
      'Night tier 2 rate mult.': nightNA ? '' : fv(formIds.nighttier2rate),
      'Night tier 2 benefit mult.': nightNA ? '' : fv(formIds.nighttier2benefit),
      'Precedence rule': fv(formIds.precedence)
    };
    var hasStructure = fields['Regular threshold'] || fields['OT 1.5 threshold'] || fields['OT 2x threshold'];
    var pickedCodes = getSelectedCodes();
    var codesSummary = pickedCodes.length
      ? pickedCodes.join(', ') + ' (' + pickedCodes.length + ' code' + (pickedCodes.length>1?'s':'') + ')'
      : 'Not yet mapped to labor type codes';

    var canWriteLive = (typeof upaeCreateRule === 'function' && typeof upaeUpdateRule === 'function');
    var statusLabel = hasStructure ? 'Structured' : 'Draft - Prose Only';
    // Union is a SharePoint Lookup column — Graph only accepts it as
    // UnionLookupId (the target Unions item's numeric id), never as plain
    // text on the "Union" key itself (that gets silently dropped). Resolve
    // it by Title if we can write live; if the union doesn't exist yet in
    // the Unions list, the rule still saves, just without that link.
    var unionLookupId = null;
    if(canWriteLive){
      try{ unionLookupId = await upaeGetListItemId('unions', name); }
      catch(err){ console.error('UPAE: could not resolve union lookup id for', name, err); }
    }
    var fieldsPayload = {
      Title: name,
      Status: statusLabel,
      WorkWeekStartDay: fields['Work week start day'],
      StandardWorkdayHrs: parseFloat(fields['Standard workday (hrs)']) || null,
      RegularThreshold: fields['Regular threshold'],
      RegularApplicableDays: fields['Regular applicable days'],
      OT15TriggerType: fields['OT 1.5 trigger type'] || 'Not Applicable',
      OT15Threshold: fields['OT 1.5 threshold'],
      OT15ApplicableDays: fields['OT 1.5 applicable days'],
      OT15RateMult: parseFloat(fields['OT 1.5 rate mult.']) || null,
      OT15BenefitMult: parseFloat(fields['OT 1.5 benefit mult.']) || null,
      OT2TriggerType: fields['OT 2x trigger type'] || 'Not Applicable',
      OT2Threshold: fields['OT 2x threshold'],
      OT2ApplicableDays: fields['OT 2x applicable days'],
      OT2RateMult: parseFloat(fields['OT 2x rate mult.']) || null,
      OT2BenefitMult: parseFloat(fields['OT 2x benefit mult.']) || null,
      NightTriggerType: fields['Night trigger type'] || 'Not Applicable',
      NightTriggerCondition: fields['Night trigger condition'],
      NightWindowStart: fields['Night window start'],
      NightWindowEnd: fields['Night window end'],
      NightApplicability: fields['Night applicability'],
      NightTier1Threshold: fields['Night tier 1 threshold'],
      NightTier1RateMult: parseFloat(fields['Night tier 1 rate mult.']) || null,
      NightTier1BenefitMult: parseFloat(fields['Night tier 1 benefit mult.']) || null,
      NightTier2Threshold: fields['Night tier 2 threshold'],
      NightTier2RateMult: parseFloat(fields['Night tier 2 rate mult.']) || null,
      NightTier2BenefitMult: parseFloat(fields['Night tier 2 benefit mult.']) || null,
      PrecedenceRule: fields['Precedence rule']
    };
    if(unionLookupId) fieldsPayload.UnionLookupId = unionLookupId;

    addSaveBtn.disabled = true;
    addSaveBtn.textContent = 'Saving to SharePoint...';
    addModalNote.textContent = 'Saving to SharePoint...';
    var writeError = null;

    if(editingId){
      var idx = rules.map(function(x){return x.id;}).indexOf(editingId);
      if(idx > -1){
        var existing = rules[idx];
        if(canWriteLive){
          try{
            if(existing.spItemId){
              await upaeUpdateRule(existing.spItemId, fieldsPayload);
            } else {
              var createdEdit = await upaeCreateRule(Object.assign({ RuleID: existing.id }, fieldsPayload));
              existing.spItemId = createdEdit.id;
            }
            for(var i=0;i<pickedCodes.length;i++){
              if((existing.laborTypeIds || []).indexOf(pickedCodes[i]) === -1){
                await upaeCreateRuleLaborType(name, laborTypeNameForCode(pickedCodes[i]));
              }
            }
          } catch(err){
            console.error('UPAE: failed to save rule to SharePoint', err);
            writeError = err;
          }
        }
        existing.name = name;
        existing.cba = fv(formIds.cba) || '—';
        existing.verified = fv(formIds.verified) || 'Not yet verified';
        existing.codes = codesSummary;
        existing.laborTypeIds = pickedCodes;
        existing.status = (existing.status === 'ambiguous') ? existing.status : (hasStructure ? 'structured' : existing.status);
        existing.prose = fv(formIds.prose);
        existing.fields = hasStructure ? fields : existing.fields;
        if(hasStructure && existing.status !== 'ambiguous'){ existing.status = 'structured'; existing.statusLabel = 'Structured'; }
        renderList(existing.id);
        renderDetail(existing);
      }
    } else {
      var newId = 'R-' + String(100 + rules.length).slice(1);
      var newRule = {
        id: newId, name: name, cba: fv(formIds.cba) || '—', verified: fv(formIds.verified) || 'Not yet verified',
        codes: codesSummary,
        laborTypeIds: pickedCodes,
        prose: fv(formIds.prose),
        status: hasStructure ? 'structured' : 'draft',
        statusLabel: hasStructure ? 'Structured' : 'Draft — prose only',
        fields: hasStructure ? fields : null,
        examples: duplicateSourceExamples ? JSON.parse(JSON.stringify(duplicateSourceExamples)) : {}
      };
      if(canWriteLive){
        try{
          var created = await upaeCreateRule(Object.assign({ RuleID: newId }, fieldsPayload));
          newRule.spItemId = created.id;
          for(var j=0;j<pickedCodes.length;j++){
            await upaeCreateRuleLaborType(name, laborTypeNameForCode(pickedCodes[j]));
          }
          if(duplicateSourceExamples){
            var exCats = Object.keys(duplicateSourceExamples);
            for(var ci=0;ci<exCats.length;ci++){
              var cat = exCats[ci];
              var exList = duplicateSourceExamples[cat];
              for(var ei=0;ei<exList.length;ei++){
                // Each copy needs its OWN new SharePoint item id — newRule.examples
                // still holds spItemIds copied from the source rule's examples at
                // this point, which would point Edit/Delete at the wrong row.
                var createdEx = await upaeCreateRuleExample(name, cat, exList[ei].ok, exList[ei].text);
                if(newRule.examples[cat] && newRule.examples[cat][ei]){
                  newRule.examples[cat][ei].spItemId = createdEx && createdEx.id;
                }
              }
            }
          }
        } catch(err){
          console.error('UPAE: failed to create rule in SharePoint', err);
          writeError = err;
        }
      }
      duplicateSourceExamples = null;
      rules.push(newRule);
      if(!unions.some(function(u){ return u.name === newRule.name; })){
        unions.push({ id: newRule.id, name: newRule.name });
        renderUnionAdminList();
        populateUnionOptions();
        populateLaborUnionSelect();
      }
      renderList(newRule.id);
      renderDetail(newRule);
    }

    addSaveBtn.disabled = false;
    if(writeError){
      addSaveBtn.textContent = 'Save rule';
      addModalNote.textContent = 'Saved to this screen, but the SharePoint write failed: ' + writeError.message + '. Check your connection/sign-in and try again.';
    } else {
      addSaveBtn.textContent = canWriteLive ? 'Saved to SharePoint' : 'Saved';
      setTimeout(function(){ closeAddModal(); addSaveBtn.textContent = 'Save rule'; }, 700);
    }
  });
}

// ---- Add example modal ----
var exOverlay = document.getElementById('addExampleModalOverlay');
var exCloseBtn = document.getElementById('closeAddExampleModal');
var exCancelBtn = document.getElementById('cancelAddExample');
var exSaveBtn = document.getElementById('saveAddExample');
var exCategorySel = document.getElementById('ex_category');
var exTextArea = document.getElementById('ex_text');
var exSub = document.getElementById('addExampleModalSub');
var exTargetRuleId = null;
var exEditRef = null; // {category, index} when editing an existing example; null when adding

function openExampleModal(r, editRef){
  exTargetRuleId = r.id;
  exEditRef = editRef || null;
  var editing = exEditRef && r.examples && r.examples[exEditRef.category] && r.examples[exEditRef.category][exEditRef.index];
  if(editing){
    exSub.textContent = 'Editing an example for ' + r.name + '.';
    exCategorySel.value = exEditRef.category;
    exTextArea.value = editing.text;
    var radios1 = document.getElementsByName('ex_type');
    var wantOk = editing.ok;
    (radios1.forEach ? radios1 : Array.prototype.slice.call(radios1)).forEach(function(rb){ rb.checked = (rb.value === 'ok') === wantOk; });
    exSaveBtn.textContent = 'Save changes';
  } else {
    exSub.textContent = 'Add a correct or incorrect weekly-report example for ' + r.name + '.';
    exTextArea.value = '';
    var radios = document.getElementsByName('ex_type');
    radios.forEach && radios.forEach(function(rb){ rb.checked = (rb.value === 'ok'); });
    if(!radios.forEach){ for(var i=0;i<radios.length;i++){ radios[i].checked = (radios[i].value === 'ok'); } }
    exSaveBtn.textContent = 'Add example';
  }
  exOverlay.hidden = false;
}
function closeExampleModal(){ exOverlay.hidden = true; exEditRef = null; }

if(exOverlay){
  exCloseBtn.addEventListener('click', closeExampleModal);
  exCancelBtn.addEventListener('click', closeExampleModal);
  // Same reasoning as the rule modal above — no click-outside-to-close.

  exSaveBtn.addEventListener('click', async function(){
    var text = exTextArea.value.trim();
    if(!text){ exTextArea.focus(); return; }
    var category = exCategorySel.value;
    var typeRadio = document.querySelector('input[name="ex_type"]:checked');
    var isOk = !typeRadio || typeRadio.value === 'ok';

    var r = rules.filter(function(x){ return x.id === exTargetRuleId; })[0];
    if(!r){ closeExampleModal(); return; }

    var isEdit = !!(exEditRef && r.examples && r.examples[exEditRef.category] && r.examples[exEditRef.category][exEditRef.index]);
    var existingEx = isEdit ? r.examples[exEditRef.category][exEditRef.index] : null;
    var defaultBtnLabel = isEdit ? 'Save changes' : 'Add example';

    exSaveBtn.disabled = true;
    exSaveBtn.textContent = 'Saving...';
    var writeError = null;

    if(isEdit){
      if(existingEx.spItemId && typeof upaeUpdateRuleExample === 'function'){
        try{
          await upaeUpdateRuleExample(existingEx.spItemId, category, isOk, text);
        } catch(err){
          console.error('UPAE: failed to update example in SharePoint', err);
          writeError = err;
        }
      }
      if(!writeError){
        // Remove from its old slot (category may have changed) and re-add.
        r.examples[exEditRef.category].splice(exEditRef.index, 1);
        if(!r.examples[exEditRef.category].length) delete r.examples[exEditRef.category];
        if(!r.examples[category]) r.examples[category] = [];
        r.examples[category].push({ ok: isOk, text: text, spItemId: existingEx.spItemId });
      }
    } else {
      if(typeof upaeCreateRuleExample === 'function'){
        try{
          var created = await upaeCreateRuleExample(r.name, category, isOk, text);
          var newSpId = created && created.id;
        } catch(err){
          console.error('UPAE: failed to save example to SharePoint', err);
          writeError = err;
        }
      }
      if(!writeError){
        if(!r.examples) r.examples = {};
        if(!r.examples[category]) r.examples[category] = [];
        r.examples[category].push({ ok: isOk, text: text, spItemId: newSpId });
      }
    }

    if(r.id === currentId) renderDetail(r);

    exSaveBtn.disabled = false;
    if(writeError){
      exSub.textContent = 'Saved to this screen, but the SharePoint write failed: ' + writeError.message + '. Check your connection/sign-in and try again.';
      exSaveBtn.textContent = defaultBtnLabel;
    } else {
      exSaveBtn.textContent = (typeof upaeCreateRuleExample === 'function') ? 'Saved to SharePoint' : (isEdit ? 'Updated' : 'Added');
      setTimeout(function(){ closeExampleModal(); exSaveBtn.textContent = defaultBtnLabel; }, 600);
    }
  });
}
