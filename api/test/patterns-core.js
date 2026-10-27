const fs=require('fs'), vm=require('vm'), assert=require('assert');
const ctx={}; vm.createContext(ctx); vm.runInContext(fs.readFileSync(__dirname+'/../../js/patterns-core.js','utf8')+';this.patDetect=patDetect;this.patMedian=patMedian;',ctx);
const weeks=['2026-09-20','2026-09-27','2026-10-04','2026-10-11','2026-10-18'];
const rows=[]; const add=(id,nm,w,dayOff,rt,ot,dt)=>{const d=new Date(w+'T00:00:00');d.setDate(d.getDate()+dayOff);const iso=d.toISOString().slice(0,10);rows.push({employeeId:id,employeeName:nm,workDate:iso,weekStart:w,weekNo:1,rt,ot,dt});};
weeks.forEach((w,i)=>{
  for(let d=1;d<=5;d++){ add('AGUE1','Aguiluz',w,d,8,0,0); add('BASE1','Base',w,d,8,0,0); for(let k=0;k<30;k++){ add('U'+k,'U'+k,w,d,8,i==4?2:0,0); } }
  if(i>=1) add('AGUE1','Aguiluz',w,6,0,0,3);          // Saturday DT in 4 weeks incl target
  if(i==4){ for(let d=1;d<=5;d++) add('BASE1','Base',w,d,0,6,0); }  // spike: 40 -> 70
  if(i>=2) add('LONG1','Long',w,2,8,5,0);               // 13h days in 3 weeks
});
const unionOf=id=>id.startsWith('U')?'Local 14':'';
const r=ctx.patDetect(rows,unionOf,'2026-10-18');
const types=r.outliers.map(o=>o.type).sort(); console.log(types, r.notes, r.outliers.map(o=>o.facts[0]));
assert(types.includes('repeat_dt_day')&&types.includes('long_day')&&types.includes('spike')&&types.includes('union_premium'));
const r2=ctx.patDetect(rows,unionOf,'2026-10-18'); assert.deepStrictEqual(JSON.stringify(r),JSON.stringify(r2),'reproducible');
const early=ctx.patDetect(rows,unionOf,'2026-09-27'); assert.strictEqual(early.outliers.length,0); assert(early.notes.length>=1);
assert.strictEqual(ctx.patDetect(rows,unionOf,'2030-01-01').outliers.length,0);
const strict=ctx.patDetect(rows,unionOf,'2026-10-18',{repeatWeeks:5}); assert(!strict.outliers.some(o=>o.type==='repeat_dt_day'));
console.log('core ok');
