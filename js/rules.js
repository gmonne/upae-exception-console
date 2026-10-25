// Shared helper (kept first so every later script can use it)
function escapeHtml(s){
  return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];
  });
}

// ---- Rule Library ----
var rules = [
  { id:'R-001', name:'Local 14 Operating Engineers', status:'structured', statusLabel:'Structured', cba:'Local 14 CBA, Art. 12 (Hours of Work & Overtime)', verified:'Gus, Sept 22, 2026', codes:'L041, L044, L052 (3 of 11 codes matched — confirm remaining)',
    fields:{
      'Work week start day':'Monday', 'Standard workday (hrs)':'8',
      'Regular threshold':'Up to 8 hrs/day, Mon–Fri', 'Regular applicable days':'Mon, Tue, Wed, Thu, Fri',
      'OT 1.5 trigger type':'Not Applicable', 'OT 1.5 threshold':'', 'OT 1.5 applicable days':'', 'OT 1.5 rate mult.':'', 'OT 1.5 benefit mult.':'',
      'OT 2x trigger type':'Daily (over N hrs/day)', 'OT 2x threshold':'Over 8 hrs/day Mon–Fri; all hours Saturday; all hours Sunday',
      'OT 2x applicable days':'Mon–Fri (over 8), Sat (all), Sun (all)', 'OT 2x rate mult.':'2.0', 'OT 2x benefit mult.':'CONFIRM WITH PAYROLL — do not default to 1.0',
      'Night trigger type':'Send-home / Recall event',
      'Night trigger condition':'Crew sent home and recalled to work a second shift, any weekday. Starts immediately regardless of RT hours already worked that day or the length of the gap between shifts.',
      'Night applicability':'Weekdays only — Saturday/Sunday hours are always OT2x, never Night Differential',
      'Night tier 1 threshold':'First 8 hrs of the recalled second shift', 'Night tier 1 rate mult.':'1.6', 'Night tier 1 benefit mult.':'2.0',
      'Night tier 2 threshold':'Hours beyond 8 in the recalled second shift', 'Night tier 2 rate mult.':'2.0', 'Night tier 2 benefit mult.':'2.0',
      'Precedence rule':'Continuous work past 8 hrs in one uninterrupted shift → OT2x. A send-home-and-recall second shift → Night Differential (Tier 1 then Tier 2) instead of OT2x, even though the wage rate converges at 2.0x at Tier 2 — the benefit multiplier still differs. Weekend hours are always OT2x regardless of any recall.'
    },
    examples:{
      'Regular':[
        {ok:true, text:'Mon–Fri, 8.0 hrs/day, no Sat/Sun — 40.0 RT for the week, no OT.'},
        {ok:false, text:'Mon–Fri, 8.0 hrs/day reported, but one day shows 6.0 RT + 2.0 OT — should be 8.0 RT (OT only kicks in over 8/day).'}
      ],
      'Overtime 1.5x':[
        {ok:true, text:'No hours coded at 1.5x anywhere on the week — correct. Local 14 has no 1.5x tier at all; every overtime hour is 2.0x.'},
        {ok:false, text:'Saturday, 4.0 hrs worked — reported as 4.0 OT (1.5x). Incorrect: Local 14 has no 1.5x tier — this must be 4.0 OT2 (2.0x), same as any other overtime hour for this union.'}
      ],
      'Overtime 2x':[
        {ok:true, text:'Tuesday, 10.0 hrs worked — reported as 8.0 RT + 2.0 OT2 (2.0x). Correct: hours beyond 8/day Mon–Fri go straight to 2.0x, there is no 1.5x step in between.'},
        {ok:false, text:'Tuesday, 10.0 hrs worked — reported as 8.0 RT + 2.0 OT (1.5x). Incorrect: Local 14 has no 1.5x tier; the 2 extra hours must be OT2 (2.0x).'},
        {ok:true, text:'Saturday, 4.0 hrs worked — reported as 4.0 OT2 (2.0x). Correct: all Saturday and Sunday hours are OT2, regardless of hours already worked that week and regardless of any recall event.'},
        {ok:false, text:'Sunday, 4.0 hrs worked following a Friday recall — reported as 4.0 Night Differential Tier 1. Incorrect: weekend hours are always OT2 (2.0x), never Night Differential, even in a week where a recall happened.'}
      ],
      'Night Differential':[
        {ok:true, text:'Crew works 8.0 RT Wednesday, is sent home, then recalled for a 6.0-hr second shift the same night — reported as 8.0 RT + 6.0 Night Differential Tier 1 (1.6x wages / 2.0x benefits). Correct: the recall triggers Night Differential immediately, regardless of RT hours already worked or the length of the gap between shifts.'},
        {ok:false, text:'Same recall scenario reported as 8.0 RT + 6.0 OT2 (2.0x). Incorrect: a send-home-and-recall second shift on a weekday is Night Differential, not standard overtime — even though Tier 2\'s wage rate happens to also be 2.0x, the classification and benefit multiplier still differ from OT2.'},
        {ok:true, text:'Crew is recalled for a 10.0-hr second shift — reported as 8.0 Night Differential Tier 1 (1.6x/2.0x) + 2.0 Night Differential Tier 2 (2.0x/2.0x). Correct: only the first 8 hours of the recalled shift are Tier 1; hours beyond 8 move to Tier 2.'},
        {ok:false, text:'Same 10.0-hr recalled shift reported entirely as 10.0 Night Differential Tier 1 (1.6x). Incorrect: hours beyond the first 8 in the recalled shift must be split out as Tier 2, not left at the Tier 1 rate.'}
      ]
    }
  },
  { id:'R-002', name:'Local 15 Operating Engineers', status:'ambiguous', statusLabel:'Ambiguous mapping', cba:'—', verified:'Not yet verified', codes:'Labor Type ID 15 ("15 Operator Paid As Operator 14") — ambiguous between Local 14 / Local 15, needs Payroll confirmation' },
  { id:'R-003', name:'Local 15C Mechanics', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-004', name:'Local 15D Surveyors', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-005', name:'Local 731 Laborers Union', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-006', name:'Local 1010 Pavers & Road Builders', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-007', name:'Local 157 Carps', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-008', name:'Local 1536 Timberman', status:'ambiguous', statusLabel:'Auto-resolved — confirm', cba:'—', verified:'Not yet verified', codes:'Overlaps Local 1556 codes — auto-resolved to 1536, needs confirm' },
  { id:'R-009', name:'Local 1456 Dockbuilders', status:'ambiguous', statusLabel:'Auto-resolved — confirm', cba:'—', verified:'Not yet verified', codes:'Overlaps Local 1556 codes — auto-resolved to 1456, needs confirm' },
  { id:'R-010', name:'Local 1556 Marine Divers', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-011', name:'Local 46 Metallic Lathers', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-012', name:'Local 60 Metallic Lathers', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'Auto-matched, pending confirm' },
  { id:'R-013', name:'Non Union Employees', status:'draft', statusLabel:'Draft — prose only', cba:'—', verified:'Not yet verified', codes:'45 labor types unmatched to any union fall here or need new categories — reviewing with Payroll' }
];

// ---- Admin: Unions & Labor Types (reference data) ----
var unions = rules.map(function(r){ return { id: r.id, name: r.name }; });

var laborTypes = [
  { id:'LT-001', code:'L062', name:'1 Plumbers Helper B', unionName:null },
  { id:'LT-002', code:'L060', name:'1 Plumbers Journeyman A', unionName:null },
  { id:'LT-003', code:'L100', name:'14 Driller', unionName:'Local 14 Operating Engineers' },
  { id:'LT-004', code:'L106', name:'14 Operator Utility Compressor', unionName:'Local 14 Operating Engineers' },
  { id:'LT-005', code:'L008', name:'14 Operators Backhoe/Excavator', unionName:'Local 14 Operating Engineers' },
  { id:'LT-006', code:'L073', name:'14 Operators Compressor Engineer', unionName:'Local 14 Operating Engineers' },
  { id:'LT-007', code:'L087', name:'14 Operators Concrete Pumps', unionName:'Local 14 Operating Engineers' },
  { id:'LT-008', code:'L061', name:'14 Operators Crane', unionName:'Local 14 Operating Engineers' },
  { id:'LT-009', code:'L083', name:'14 Operators Crane-2', unionName:'Local 14 Operating Engineers' },
  { id:'LT-010', code:'L037', name:'14 Operators Maintenance Engineer', unionName:'Local 14 Operating Engineers' },
  { id:'LT-011', code:'L088', name:'14 Operators Portable Compressor', unionName:'Local 14 Operating Engineers' },
  { id:'LT-012', code:'L025', name:'14 Operators Roller', unionName:'Local 14 Operating Engineers' },
  { id:'LT-013', code:'L009', name:'14 Operators Spreader', unionName:'Local 14 Operating Engineers' },
  { id:'LT-014', code:'L118', name:'14 Well Pumps', unionName:'Local 14 Operating Engineers' },
  { id:'LT-015', code:'L107', name:'15 Operator Paid As Operator 14 Crane', unionName:'Local 15 Operating Engineers' },
  { id:'LT-016', code:'L010', name:'15 Operators Backhoe', unionName:'Local 15 Operating Engineers' },
  { id:'LT-017', code:'L078', name:'15 Operators Backhoe/Large', unionName:'Local 15 Operating Engineers' },
  { id:'LT-018', code:'L084', name:'15 Operators Generator Engineer', unionName:'Local 15 Operating Engineers' },
  { id:'LT-019', code:'L081', name:'15 Operators Junior Engineer-2', unionName:'Local 15 Operating Engineers' },
  { id:'LT-020', code:'L011', name:'15 Operators Maintenance Engineer', unionName:'Local 15 Operating Engineers' },
  { id:'LT-021', code:'L047', name:'15 Operators Maintenance Engineer Foreman', unionName:'Local 15 Operating Engineers' },
  { id:'LT-022', code:'L031', name:'15 Operators Maintenance Engineer/Welder', unionName:'Local 15 Operating Engineers' },
  { id:'LT-023', code:'L021', name:'15 Operators Maintenance Foreman', unionName:'Local 15 Operating Engineers' },
  { id:'LT-024', code:'L012', name:'15 Operators Oiler', unionName:'Local 15 Operating Engineers' },
  { id:'LT-025', code:'L039', name:'15 Operators Oiler-Large Crane', unionName:'Local 15 Operating Engineers' },
  { id:'LT-026', code:'L013', name:'15 Operators Pump Engineers', unionName:'Local 15 Operating Engineers' },
  { id:'LT-027', code:'L074', name:'15 Operators Utility Welder', unionName:'Local 15 Operating Engineers' },
  { id:'LT-028', code:'L101', name:'15-15A Jr.Engineer #2', unionName:'Local 15 Operating Engineers' },
  { id:'LT-029', code:'GALT11', name:'Barge-Conveyor Operatir', unionName:null },
  { id:'LT-030', code:'GALT12', name:'Barge-Excavator Operator', unionName:null },
  { id:'LT-031', code:'GALT10', name:'Barge-Level Monitor', unionName:null },
  { id:'LT-032', code:'GALT13', name:'Barge-Material Handler Operator', unionName:null },
  { id:'LT-033', code:'GALT14', name:'Barge-Skid Steer', unionName:null },
  { id:'LT-034', code:'L132', name:'CARPENTERS LOCAL 157', unionName:null },
  { id:'LT-035', code:'CA001', name:'Castcapa-Foreman', unionName:null },
  { id:'LT-036', code:'CA002', name:'Castcapa-Laborer', unionName:null },
  { id:'LT-037', code:'GALT21', name:'Cold Patch Delivery Driver', unionName:null },
  { id:'LT-038', code:'GALT7', name:'Cold Patch-Bagger', unionName:null },
  { id:'LT-039', code:'GALT8', name:'Cold Patch-Bobcat Operator', unionName:null },
  { id:'LT-040', code:'GALT5', name:'Crush-Crusher Operator', unionName:null },
  { id:'LT-041', code:'GALT6', name:'Crush-Picker', unionName:null },
  { id:'LT-042', code:'GALT4', name:'Crushing-Excavator Operator', unionName:null },
  { id:'LT-043', code:'L068', name:'Field Assistant', unionName:null },
  { id:'LT-044', code:'L149', name:'FIELD ENGINEER', unionName:null },
  { id:'LT-045', code:'L001', name:'Flagman', unionName:null },
  { id:'LT-046', code:'GALT19', name:'General-Foreman', unionName:null },
  { id:'LT-047', code:'GALT20', name:'General-Laborer', unionName:null },
  { id:'LT-048', code:'GALT18', name:'General-Mechanic', unionName:null },
  { id:'LT-049', code:'L150', name:'Iron worker local 361 2nd yr Apprentice', unionName:null },
  { id:'LT-050', code:'L133', name:'JR. ESTIMATOR', unionName:null },
  { id:'LT-051', code:'L153', name:'LOCAL 14 MASTER MECHANIC', unionName:null },
  { id:'LT-052', code:'L148', name:'LOCAL 14 VERMEER OPERATOR', unionName:null },
  { id:'LT-053', code:'L154', name:'Local 15 Concrete Pump', unionName:null },
  { id:'LT-054', code:'L135', name:'LOCAL 60 FIRST YR O.Q GAS MECH (A)', unionName:null },
  { id:'LT-055', code:'L119', name:'Local 60 Flagperson', unionName:null },
  { id:'LT-056', code:'L136', name:'LOCAL 60 GROUP 1', unionName:null },
  { id:'LT-057', code:'L137', name:'LOCAL 60 GROUP 1 A', unionName:null },
  { id:'LT-058', code:'L138', name:'LOCAL 60 GROUP III', unionName:null },
  { id:'LT-059', code:'L139', name:'LOCAL 60 GROUP IV', unionName:null },
  { id:'LT-060', code:'L140', name:'LOCAL 60 GROUP V', unionName:null },
  { id:'LT-061', code:'L141', name:'LOCAL 60 GROUP VI', unionName:null },
  { id:'LT-062', code:'L142', name:'LOCAL 60 GROUP VIA', unionName:null },
  { id:'LT-063', code:'L120', name:'Local 60 Labor Group II', unionName:null },
  { id:'LT-064', code:'L134', name:'LOCAL 60 O.Q GAS MECH (A MECHANIC)', unionName:null },
  { id:'LT-065', code:'L143', name:'LOCAL 60 TIER 1 APPRENTICE', unionName:null },
  { id:'LT-066', code:'L144', name:'LOCAL 60 TIER 2 APPRENTICE', unionName:null },
  { id:'LT-067', code:'L030', name:'46 Lathers Journeyman', unionName:'Local 46 Metallic Lathers' },
  { id:'LT-068', code:'L070', name:'60 Laborers Laborer', unionName:'Local 60 Metallic Lathers' },
  { id:'LT-069', code:'L076', name:'137 Operators Backhoe', unionName:null },
  { id:'LT-070', code:'L016', name:'282 Teamster Driver', unionName:null },
  { id:'LT-071', code:'L027', name:'282 Teamsters Shop Steward', unionName:null },
  { id:'LT-072', code:'L125', name:'361 Ironworkers 1st year Apprentice', unionName:null },
  { id:'LT-073', code:'L126', name:'361 Ironworkers 3rd year Apprentice', unionName:null },
  { id:'LT-074', code:'127', name:'361 Ironworkers Apprentice 2', unionName:null },
  { id:'LT-075', code:'L124', name:'361 Ironworkers Foreman', unionName:null },
  { id:'LT-076', code:'L123', name:'361 Ironworkers Journeyman', unionName:null },
  { id:'LT-077', code:'L019', name:'731 Laborers Apprentice 0 - 1000', unionName:'Local 731 Laborers Union' },
  { id:'LT-078', code:'L041', name:'731 Laborers Apprentice 1000 - 2000', unionName:'Local 731 Laborers Union' },
  { id:'LT-079', code:'L055', name:'731 Laborers Apprentice 2000 - 3000', unionName:'Local 731 Laborers Union' },
  { id:'LT-080', code:'L082', name:'731 Laborers Apprentice 3000 - 4000', unionName:'Local 731 Laborers Union' },
  { id:'LT-081', code:'L006', name:'731 Laborers Foreman', unionName:'Local 731 Laborers Union' },
  { id:'LT-082', code:'L004', name:'731 Laborers Laborer', unionName:'Local 731 Laborers Union' },
  { id:'LT-083', code:'L121', name:'740 Millwright Journeyman', unionName:null },
  { id:'LT-084', code:'L007', name:'780 Masons Mason', unionName:null },
  { id:'LT-085', code:'L103', name:'1010 Laborer Micro Paver', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-086', code:'L104', name:'1010 Laborer Small Equip Oper', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-087', code:'L079', name:'1010 Laborers Apprentice', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-088', code:'L005', name:'1010 Laborers Forman', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-089', code:'L024', name:'1010 Laborers Formsetter', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-090', code:'L002', name:'1010 Laborers Laborer', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-091', code:'L028', name:'1010 Laborers Paving Raker', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-092', code:'L044', name:'1010 Laborers Paving Screed Person', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-093', code:'L075', name:'1010 Laborers Paving Shoveler', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-094', code:'L003', name:'1010/1018 Laborers Raker', unionName:'Local 1010 Pavers & Road Builders' },
  { id:'LT-095', code:'L117', name:'1536 Timberman App 4th Year', unionName:'Local 1536 Timberman' },
  { id:'LT-096', code:'L122', name:'1556 Carpenter Foreman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-097', code:'L128', name:'1556 MARINE DIVERS', unionName:'Local 1556 Marine Divers' },
  { id:'LT-098', code:'L029', name:'1556/45 Carpenter Journeyman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-099', code:'L097', name:'1556/608 Carpenter General Foreman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-100', code:'L038', name:'1556/926 Carpenter Journeyman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-101', code:'L089', name:'1556/1456 Dockbuilders Apprentice 1st Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-102', code:'L091', name:'1556/1456 Dockbuilders Apprentice 2nd Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-103', code:'L095', name:'1556/1456 Dockbuilders Apprentice 3rd Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-104', code:'L098', name:'1556/1456 Dockbuilders Apprentice 4th Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-105', code:'L034', name:'1556/1456 Dockbuilders Foreman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-106', code:'L056', name:'1556/1456 Dockbuilders Journeyman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-107', code:'L080', name:'1556/1536 Timbermen Apprentice 1st Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-108', code:'L090', name:'1556/1536 Timbermen Apprentice 2nd Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-109', code:'L092', name:'1556/1536 Timbermen Apprentice 3rd Year', unionName:'Local 1556 Marine Divers' },
  { id:'LT-110', code:'L018', name:'1556/1536 Timbermen Foreman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-111', code:'L017', name:'1556/1536 Timbermen Journeyman', unionName:'Local 1556 Marine Divers' },
  { id:'LT-112', code:'L059', name:'1556/1536 Timbermen Shop Steward', unionName:'Local 1556 Marine Divers' },
  { id:'LT-113', code:'GALT22', name:'Asphalt Plant Management', unionName:null },
  { id:'LT-114', code:'GALT23', name:'Asphalt Plant Quality Assurance', unionName:null },
  { id:'LT-115', code:'GALT3', name:'Asphalt-Loader Operator', unionName:null },
  { id:'LT-116', code:'GALT1', name:'Asphalt-Plant Operator', unionName:null },
  { id:'LT-117', code:'GALT2', name:'Asphalt-Temp Monitor', unionName:null },
  { id:'LT-118', code:'L094', name:'Assistant Project Manager', unionName:null },
  { id:'LT-119', code:'L096', name:'15-C Helper', unionName:'Local 15 Operating Engineers' },
  { id:'LT-120', code:'L108', name:'15C Operator B-Helper', unionName:'Local 15C Mechanics' },
  { id:'LT-121', code:'L014', name:'15C Operator Mechanic', unionName:'Local 15C Mechanics' },
  { id:'LT-122', code:'L015', name:'15D Surveyors Instrument Man', unionName:'Local 15D Surveyors' },
  { id:'LT-123', code:'L045', name:'15D Surveyors Party Chief', unionName:'Local 15D Surveyors' },
  { id:'LT-124', code:'L051', name:'15D Surveyors Rodman', unionName:'Local 15D Surveyors' },
  { id:'LT-125', code:'L077', name:'29 Drill Runners Jackhammer Op', unionName:null },
  { id:'LT-126', code:'L105', name:'29 Laborer Jackhammer', unionName:null },
  { id:'LT-127', code:'L066', name:'40 Ironworkers Apprentice', unionName:null },
  { id:'LT-128', code:'L064', name:'40 Ironworkers Assistant Foreman', unionName:null },
  { id:'LT-129', code:'L063', name:'40 Ironworkers Foreman', unionName:null },
  { id:'LT-130', code:'L065', name:'40 Ironworkers Journeyman', unionName:null },
  { id:'LT-131', code:'L035', name:'46 Lathers Foreman', unionName:'Local 46 Metallic Lathers' }
];

function unionNameById(){ return unions.map(function(u){ return u.name; }); }

// ---- Insights: live coverage stats across Unions / Labor Types / Rules ----
function renderRuleLibraryKpis(){
  var el = document.getElementById('ruleLibraryKpis');
  if(!el) return;

  var totalUnions = unions.length;
  var structuredRules = rules.filter(function(r){ return r.status === 'structured'; }).length;
  var draftOrAmbiguousRules = rules.length - structuredRules;

  var totalLabor = laborTypes.length;
  var mappedLabor = laborTypes.filter(function(lt){ return !!lt.unionName; }).length;
  var unmappedLabor = totalLabor - mappedLabor;

  var unionsWithExamples = rules.filter(function(r){
    var ex = r.examples || {};
    return Object.keys(ex).some(function(cat){ return (ex[cat] || []).length > 0; });
  }).length;

  el.innerHTML =
    '<div class="kpi"><div class="k-label">Unions covered</div><div class="k-value">' + totalUnions + '</div>' +
      '<div class="k-delta flat">' + structuredRules + ' fully structured &middot; ' + draftOrAmbiguousRules + ' in draft</div></div>' +
    '<div class="kpi"><div class="k-label">Labor type codes mapped</div><div class="k-value">' + mappedLabor + '<span class="unit">/' + totalLabor + '</span></div>' +
      '<div class="k-delta up">' + unmappedLabor + ' unmatched &middot; reviewing with Payroll</div></div>' +
    '<div class="kpi"><div class="k-label">Example pairs on file</div><div class="k-value">' + unionsWithExamples + '<span class="unit">/' + totalUnions + ' unions</span></div>' +
      '<div class="k-delta flat">Regular / OT 1.5 / OT 2x / Night, per union</div></div>' +
    '<div class="kpi"><div class="k-label">Open items for Payroll</div><div class="k-value">' + draftOrAmbiguousRules + '</div>' +
      '<div class="k-delta up">rules still draft or needing confirmation</div></div>';
}

function renderCoverageInsights(){
  var kpisEl = document.getElementById('coverageKpis');
  var statusBarEl = document.getElementById('ruleStatusBarList');
  var unionBarEl = document.getElementById('laborUnionBarList');
  var footEl = document.getElementById('coverageFootnote');
  if(!kpisEl) return; // Insights view not present on this build

  var totalUnions = unions.length;
  var totalLabor = laborTypes.length;
  var mappedLabor = laborTypes.filter(function(lt){ return !!lt.unionName; }).length;
  var unassignedLabor = totalLabor - mappedLabor;
  var mappedPct = totalLabor ? Math.round((mappedLabor / totalLabor) * 100) : 0;
  var totalRules = rules.length;
  var structuredRules = rules.filter(function(r){ return r.status === 'structured'; }).length;
  var structuredPct = totalRules ? Math.round((structuredRules / totalRules) * 100) : 0;

  kpisEl.innerHTML =
    '<div class="kpi"><div class="k-label">Unions</div><div class="k-value">' + totalUnions + '</div>' +
      '<div class="k-delta down">reference list, Admin screen</div></div>' +
    '<div class="kpi"><div class="k-label">Labor type codes</div><div class="k-value">' + totalLabor + '</div>' +
      '<div class="k-delta down">' + unassignedLabor + ' unassigned to a union</div></div>' +
    '<div class="kpi"><div class="k-label">Labor types mapped</div><div class="k-value">' + mappedPct + '<span class="unit">%</span></div>' +
      '<div class="k-delta down">' + mappedLabor + ' of ' + totalLabor + ' codes</div></div>' +
    '<div class="kpi"><div class="k-label">Rules structured</div><div class="k-value">' + structuredPct + '<span class="unit">%</span></div>' +
      '<div class="k-delta down">' + structuredRules + ' of ' + totalRules + ' rules</div></div>';

  var statusCounts = { structured: 0, ambiguous: 0, draft: 0 };
  rules.forEach(function(r){ statusCounts[r.status] = (statusCounts[r.status] || 0) + 1; });
  var statusRows = [
    { label: 'Structured', value: statusCounts.structured || 0 },
    { label: 'Ambiguous / needs confirmation', value: statusCounts.ambiguous || 0 },
    { label: 'Draft — prose only', value: statusCounts.draft || 0 }
  ];
  var statusMax = Math.max(1, statusCounts.structured || 0, statusCounts.ambiguous || 0, statusCounts.draft || 0);
  if(statusBarEl){
    statusBarEl.innerHTML = statusRows.map(function(row){
      var pct = Math.round((row.value / statusMax) * 100);
      return '<div class="ibar-row"><span class="ibar-label">' + row.label + '</span>' +
        '<span class="ibar-track"><span style="width:' + pct + '%"></span></span>' +
        '<span class="ibar-value">' + row.value + '</span></div>';
    }).join('');
  }

  if(unionBarEl){
    var byUnion = {};
    laborTypes.forEach(function(lt){
      var key = lt.unionName || 'Unassigned';
      byUnion[key] = (byUnion[key] || 0) + 1;
    });
    var unionRows = Object.keys(byUnion).map(function(k){ return { label: k, value: byUnion[k] }; })
      .sort(function(a, b){ return b.value - a.value; })
      .slice(0, 8);
    var unionMax = Math.max.apply(null, unionRows.map(function(r){ return r.value; }).concat([1]));
    unionBarEl.innerHTML = unionRows.map(function(row){
      var pct = Math.round((row.value / unionMax) * 100);
      return '<div class="ibar-row"><span class="ibar-label">' + row.label + '</span>' +
        '<span class="ibar-track"><span style="width:' + pct + '%"></span></span>' +
        '<span class="ibar-value">' + row.value + '</span></div>';
    }).join('');
  }

  if(footEl){
    footEl.textContent = totalUnions + ' union' + (totalUnions === 1 ? '' : 's') + ', ' +
      totalLabor + ' labor type code' + (totalLabor === 1 ? '' : 's') + ', and ' +
      totalRules + ' rule' + (totalRules === 1 ? '' : 's') + ' in the reference library right now' +
      (unassignedLabor ? ' — ' + unassignedLabor + ' labor type code' + (unassignedLabor === 1 ? ' is' : 's are') + ' still unassigned to a union.' : '.');
  }
}

function renderUnionAdminList(){
  renderCoverageInsights();
  renderRuleLibraryKpis();
  if(typeof renderEmployees === 'function') renderEmployees();
  var el = document.getElementById('unionList');
  var countEl = document.getElementById('unionCount');
  countEl.textContent = unions.length + ' union' + (unions.length === 1 ? '' : 's');
  if(!unions.length){ el.innerHTML = '<div class="admin-empty">No unions yet — add one below.</div>'; return; }
  var html = '';
  unions.forEach(function(u, i){
    var usedByRule = rules.filter(function(r){ return r.name === u.name; }).length;
    html += '<div class="admin-row"><div class="a-main"><div class="a-name">' + escapeHtml(u.name) + '</div>' +
      '<div class="a-meta">' + (usedByRule ? usedByRule + ' rule' + (usedByRule>1?'s':'') + ' reference this union' : 'No rule built yet') + '</div></div>' +
      '<button class="btn ghost" type="button" data-union-edit-idx="' + i + '" style="padding:4px 10px; font-size:12px;" title="Edit union">Edit</button>' +
      '<button class="a-del" type="button" data-union-idx="' + i + '" title="Remove union">&times;</button></div>';
  });
  el.innerHTML = html;
  el.querySelectorAll('[data-union-edit-idx]').forEach(function(btn){
    btn.addEventListener('click', function(){
      openEditUnionModal(parseInt(btn.getAttribute('data-union-edit-idx'), 10));
    });
  });
  el.querySelectorAll('[data-union-idx]').forEach(function(btn){
    btn.addEventListener('click', async function(){
      var idx = parseInt(btn.getAttribute('data-union-idx'), 10);
      var u = unions[idx];
      if(!u) return;
      // Real SharePoint list item ids are plain numbers; unions created locally
      // (e.g. auto-added when saving a rule for a brand-new union) get a
      // placeholder id like "U-<timestamp>" or a rule id like "R-101" and were
      // never written to the Unions list — nothing to delete there for those.
      if(/^\d+$/.test(String(u.id)) && typeof upaeDeleteUnion === 'function'){
        try{ await upaeDeleteUnion(u.id); }
        catch(err){
          console.error('UPAE: failed to delete union from SharePoint', err);
          alert('Could not delete this union from SharePoint: ' + err.message + '. It has not been removed.');
          return;
        }
      }
      unions.splice(idx, 1);
      renderUnionAdminList();
      populateUnionOptions();
      populateLaborUnionSelect();
      renderLaborAdminList();
    });
  });
}

var laborFilter = 'all'; // 'all' | 'unassigned'

function laborUnionOptionsHtml(selected){
  var opts = '<option value=""' + (!selected ? ' selected' : '') + '>Unassigned</option>';
  opts += unions.map(function(u){
    return '<option value="' + u.name + '"' + (u.name === selected ? ' selected' : '') + '>' + u.name + '</option>';
  }).join('');
  return opts;
}

function renderLaborAdminList(){
  renderCoverageInsights();
  renderRuleLibraryKpis();
  if(typeof renderEmployees === 'function') renderEmployees();
  var el = document.getElementById('laborList');
  var countEl = document.getElementById('laborCount');
  var unassignedCount = laborTypes.filter(function(lt){ return !lt.unionName; }).length;
  countEl.textContent = laborTypes.length + ' code' + (laborTypes.length === 1 ? '' : 's') + ' · ' + unassignedCount + ' unassigned';

  if(!laborTypes.length){ el.innerHTML = '<div class="admin-empty">No labor type codes yet — add one below.</div>'; return; }

  var visible = laborTypes
    .map(function(lt, i){ return { lt: lt, i: i }; })
    .filter(function(x){ return laborFilter === 'unassigned' ? !x.lt.unionName : true; });

  if(!visible.length){ el.innerHTML = '<div class="admin-empty">Every code has a union assigned.</div>'; return; }

  var html = '';
  visible.forEach(function(x){
    var lt = x.lt, i = x.i;
    html += '<div class="admin-row"><div class="a-code">' + lt.code + '</div><div class="a-main"><div class="a-name">' + lt.name + '</div></div>' +
      '<select class="admin-union-select" data-labor-assign-idx="' + i + '" style="max-width:180px; padding:6px 8px; border:1px solid var(--line); border-radius:6px; background:var(--bg); color:var(--ink); font:inherit; font-size:12px;">' +
      laborUnionOptionsHtml(lt.unionName) + '</select>' +
      '<button class="a-del" type="button" data-labor-idx="' + i + '" title="Remove labor type">&times;</button></div>';
  });
  el.innerHTML = html;

  el.querySelectorAll('[data-labor-idx]').forEach(function(btn){
    btn.addEventListener('click', function(){
      laborTypes.splice(parseInt(btn.getAttribute('data-labor-idx'), 10), 1);
      renderLaborAdminList();
    });
  });
  el.querySelectorAll('[data-labor-assign-idx]').forEach(function(sel){
    sel.addEventListener('change', function(){
      var idx = parseInt(sel.getAttribute('data-labor-assign-idx'), 10);
      laborTypes[idx].unionName = sel.value || null;
      renderLaborAdminList();
    });
  });
}

document.querySelectorAll('[data-labor-filter]').forEach(function(btn){
  btn.addEventListener('click', function(){
    laborFilter = btn.getAttribute('data-labor-filter');
    document.querySelectorAll('[data-labor-filter]').forEach(function(b){ b.classList.toggle('active', b === btn); });
    renderLaborAdminList();
  });
});

function populateLaborUnionSelect(){
  var sel = document.getElementById('newLaborUnion');
  if(!sel) return;
  sel.innerHTML = '<option value="">Unassigned</option>' + unions.map(function(u){ return '<option value="' + u.name + '">' + u.name + '</option>'; }).join('');
}

var addUnionBtn = document.getElementById('addUnionBtn');
var newUnionInput = document.getElementById('newUnionName');
if(addUnionBtn){
  addUnionBtn.addEventListener('click', async function(){
    var name = newUnionInput.value.trim();
    if(!name) { newUnionInput.focus(); return; }
    var newUnion = { id: 'U-' + Date.now(), name: name, unionCode: '', cbaReference: '' };
    if(typeof upaeCreateUnion === 'function'){
      addUnionBtn.disabled = true;
      addUnionBtn.textContent = 'Saving...';
      try{
        var created = await upaeCreateUnion(name, '', '');
        newUnion.id = created.id;
      } catch(err){
        console.error('UPAE: failed to create union in SharePoint', err);
        alert('Added to this screen, but the SharePoint write failed: ' + err.message + '. Edit/Delete on this union won\'t reach SharePoint until it syncs.');
      }
      addUnionBtn.disabled = false;
      addUnionBtn.textContent = '+ Add union';
    }
    unions.push(newUnion);
    newUnionInput.value = '';
    renderUnionAdminList();
    populateUnionOptions();
    populateLaborUnionSelect();
    renderLaborAdminList();
  });
}

// ---- Edit union modal ----
var euOverlay = document.getElementById('editUnionModalOverlay');
var euCloseBtn = document.getElementById('closeEditUnionModal');
var euCancelBtn = document.getElementById('cancelEditUnion');
var euSaveBtn = document.getElementById('saveEditUnion');
var euNameInput = document.getElementById('eu_name');
var euCodeInput = document.getElementById('eu_code');
var euCbaInput = document.getElementById('eu_cba');
var euEditIdx = null;

function openEditUnionModal(idx){
  var u = unions[idx];
  if(!u) return;
  euEditIdx = idx;
  euNameInput.value = u.name || '';
  euCodeInput.value = u.unionCode || '';
  euCbaInput.value = u.cbaReference || '';
  euOverlay.hidden = false;
  euNameInput.focus();
}
function closeEditUnionModal(){ euOverlay.hidden = true; euEditIdx = null; }

if(euOverlay){
  euCloseBtn.addEventListener('click', closeEditUnionModal);
  euCancelBtn.addEventListener('click', closeEditUnionModal);
  // No click-outside-to-close — same reasoning as the other modals.

  euSaveBtn.addEventListener('click', async function(){
    if(euEditIdx == null) return;
    var u = unions[euEditIdx];
    if(!u) { closeEditUnionModal(); return; }
    var newName = euNameInput.value.trim();
    if(!newName){ euNameInput.focus(); return; }
    var newCode = euCodeInput.value.trim();
    var newCba = euCbaInput.value.trim();
    var oldName = u.name;

    euSaveBtn.disabled = true;
    euSaveBtn.textContent = 'Saving...';
    var writeError = null;
    var isLiveLinked = /^\d+$/.test(String(u.id));
    if(isLiveLinked && typeof upaeUpdateUnion === 'function'){
      try{ await upaeUpdateUnion(u.id, newName, newCode, newCba); }
      catch(err){
        console.error('UPAE: failed to update union in SharePoint', err);
        writeError = err;
      }
    }

    if(!writeError){
      u.name = newName;
      u.unionCode = newCode;
      u.cbaReference = newCba;

      // A rule's Title doubles as its union's name in this app (one rule per
      // union) — keep that rule in sync so it doesn't silently detach from
      // the union it belongs to.
      if(newName !== oldName){
        var matchingRule = rules.filter(function(r){ return r.name === oldName; })[0];
        if(matchingRule){
          matchingRule.name = newName;
          if(matchingRule.spItemId && typeof upaeUpdateRule === 'function'){
            try{ await upaeUpdateRule(matchingRule.spItemId, { Title: newName }); }
            catch(err){ console.error('UPAE: failed to rename matching rule in SharePoint', err); }
          }
        }
        laborTypes.forEach(function(lt){ if(lt.unionName === oldName) lt.unionName = newName; });
      }

      renderUnionAdminList();
      populateUnionOptions();
      populateLaborUnionSelect();
      renderLaborAdminList();
      renderList(currentId);
      if(currentId){
        var currentRule = rules.filter(function(r){ return r.id === currentId; })[0];
        if(currentRule) renderDetail(currentRule);
      }
    }

    euSaveBtn.disabled = false;
    if(writeError){
      alert('Could not save this union to SharePoint: ' + writeError.message + '. Nothing was changed.');
      euSaveBtn.textContent = 'Save changes';
    } else {
      euSaveBtn.textContent = isLiveLinked ? 'Saved to SharePoint' : 'Saved';
      setTimeout(function(){ closeEditUnionModal(); euSaveBtn.textContent = 'Save changes'; }, 600);
    }
  });
}

var addLaborBtn = document.getElementById('addLaborBtn');
if(addLaborBtn){
  addLaborBtn.addEventListener('click', function(){
    var code = document.getElementById('newLaborCode').value.trim();
    var name = document.getElementById('newLaborName').value.trim();
    var unionName = document.getElementById('newLaborUnion').value || null;
    if(!code || !name){ document.getElementById(code ? 'newLaborName' : 'newLaborCode').focus(); return; }
    laborTypes.push({ id:'LT-' + Date.now(), code: code, name: name, unionName: unionName });
    document.getElementById('newLaborCode').value = '';
    document.getElementById('newLaborName').value = '';
    document.getElementById('newLaborUnion').value = '';
    renderLaborAdminList();
    renderCodesMultiselect(); // refresh in case the rule modal is behind this
  });
}

