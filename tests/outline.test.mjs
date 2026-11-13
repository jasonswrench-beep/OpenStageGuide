import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture} from './helpers/mock-office.mjs';
import {addModule,addActivity} from '../src/model.mjs';
import {searchOutline} from '../src/outline.mjs';
function sample(){let p=addModule(fixture(),'Café conflict','').next;return addActivity(p,p.modules[1].id,{title:'Discuss <choices>',icon:'Discussion',minutes:7.5,participant:'[insert scenario]',facilitator:'PRIVATE leadership notes',answers:'SECRET scenario key',lines:0}).next;}
test('unfiltered outline preserves source IDs, order, numbering, empty modules and input bytes',()=>{
  const p=addModule(sample(),'Empty module','Later objective').next,before=JSON.stringify(p),out=searchOutline(p);
  assert.equal(out.moduleCount,3);assert.equal(out.activityCount,2);assert.equal(out.rows.at(-1).id,p.modules[2].id);assert.equal(out.rows[3].activityIndex,0);assert.equal(out.rows[3].moduleIndex,1);assert.equal(JSON.stringify(p),before);
});
test('literal all-term search covers title/objective/private teaching fields and returns only display allowlists',()=>{
  const p=sample(),a=p.modules[1].activities[0];
  for(const query of ['LEADERSHIP private','secret scenario','CAFE','<choices>','[insert','discussion']){const out=searchOutline(p,{query});assert.equal(out.activityCount,1);assert.equal(out.rows.at(-1).id,a.id);assert.doesNotMatch(JSON.stringify(out),/PRIVATE leadership|SECRET scenario key|insert scenario/);}
  assert.equal(searchOutline(p,{query:'.*'}).rows.length,0);assert.equal(searchOutline(p,{query:'Practice response'}).rows.at(-1).id,p.modules[0].activities[0].id);
});
test('cue/status filters combine and keep parent context without changing original numbering',()=>{
  const p=sample(),out=searchOutline(p,{icon:'Discussion',status:'unfinished',query:'scenario'});assert.equal(out.rows.length,2);assert.equal(out.rows[0].context,true);assert.equal(out.rows[1].moduleIndex,1);assert.equal(out.rows[1].activityIndex,0);
  assert.equal(searchOutline(p,{icon:'Lecturette',status:'unfinished'}).rows.length,0);
  const missing=searchOutline(p,{status:'missing-objective'});assert.equal(missing.rows[0].id,p.modules[1].id);assert.equal(missing.activityCount,1);
});
test('unfinished and required-error filters reflect own fields and Review Guide severity',()=>{
  const p=sample();p.modules[0].objective='TBD objective';p.modules[0].activities[0].participant='';
  const unfinished=searchOutline(p,{status:'unfinished'});assert.ok(unfinished.rows.some(r=>r.id===p.modules[0].id));assert.ok(!unfinished.rows.some(r=>r.id===p.modules[0].activities[0].id));
  const errors=searchOutline(p,{status:'needs-attention'});assert.equal(errors.activityCount,1);assert.equal(errors.rows.at(-1).id,p.modules[0].activities[0].id);assert.ok(!errors.rows.some(r=>r.id===p.modules[1].id));
});
test('search rejects unsupported filters and oversized queries without changing the course',()=>{
  const p=sample(),before=JSON.stringify(p);for(const options of [{query:'x'.repeat(201)},{query:5},{icon:'unknown'},{status:'materials'}])assert.throws(()=>searchOutline(p,options));assert.equal(JSON.stringify(p),before);
});
