import test from 'node:test';
import assert from 'node:assert/strict';
import {course,addModule,addActivity} from '../src/model.mjs';
import {reviewCourse} from '../src/planning.mjs';
import {searchOutline} from '../src/outline.mjs';
import {coverageReport,editObjectiveCoverage} from '../src/objective-coverage.mjs';
import {workbookPreview,worksheetFiles} from '../src/worksheet.mjs';
import {agendaPreview,agendaFiles} from '../src/agenda.mjs';
import {materialsPreview} from '../src/materials-output.mjs';
import {facilitatorPreview} from '../src/facilitator-output.mjs';
import {selectedObjectiveLinks} from '../src/delivery-plans.mjs';

function fixture(){let p=course('Objective boundary');p=addModule(p,'Module','Perform a skill').next;p=addActivity(p,p.modules[0].id,{title:'Activity',minutes:10,participant:'Try the skill',facilitator:'Teach the skill',answers:'Debrief',lines:1}).next;const m=p.modules[0],a=m.activities[0];p.objectiveCoverage={version:1,projectId:p.id,links:[{activityId:a.id,objectiveModuleId:m.id,roles:['practice'],reviewedObjective:m.objective,rationale:'PRIVATE_RATIONALE'}],exceptions:[]};return p;}

test('defined nontext objectives remain opaque, uncounted and distinct from blank across source-only projections',()=>{
 for(const bad of [null,0,42,false,true,{}, {toString:null}, ['Perform a skill']]){
  const p=fixture(),m=p.modules[0],a=m.activities[0];m.objective=bad;const before=structuredClone(p);
  const review=reviewCourse(p);assert.ok(review.some(x=>x.severity==='error'&&x.id===m.id&&/unsupported value/i.test(x.message)));assert.ok(!review.some(x=>x.id===m.id&&/Add a learning objective/.test(x.message)));
  const outline=searchOutline(p,{status:'needs-attention'});assert.ok(outline.rows.some(x=>x.id===m.id));assert.equal(searchOutline(p,{status:'missing-objective'}).rows.length,0);assert.equal(searchOutline(p,{query:'Perform a skill'}).rows.length,0);
  const report=coverageReport(p),row=report.rows[0];assert.equal(report.supported,true);assert.equal(row.blank,false);assert.equal(row.invalid,true);assert.equal(row.objective,'');assert.equal(row.roles.practice.current.length,0);assert.equal(row.roles.practice.stale.length,0);assert.equal(row.roles.practice.unavailable.length,1);assert.equal(row.roles.practice.gap,null);assert.equal(row.roles.practice.acknowledgedGap,false);
  const selected=selectedObjectiveLinks(p,[{moduleId:m.id,activityIds:[a.id]}]);assert.equal(selected.state,'unsupported-objective-field');assert.equal(selected.summary,null);
  assert.throws(()=>workbookPreview(p),/unsupported value/i);assert.throws(()=>worksheetFiles(p),/unsupported value/i);assert.throws(()=>facilitatorPreview(p,{answers:false,participant:false}),/learning objective text/i);
  const agenda=agendaPreview(p);assert.ok(agenda.issues.some(x=>x.severity==='error'&&/unsupported value/i.test(x.message)));assert.ok(agendaFiles(p)['word/document.xml'].includes('Activity'));assert.ok(!agendaFiles(p)['word/document.xml'].includes('PRIVATE_RATIONALE'));
  assert.ok(materialsPreview(p,'2026-10-26T08:30:00.000Z'));assert.deepEqual(p,before);
  const changed={kind:'links',activityId:a.id,links:[{objectiveModuleId:m.id,roles:['practice'],reviewedObjective:'Perform a skill',rationale:'PRIVATE_RATIONALE',reviewCurrent:false}]};assert.throws(()=>editObjectiveCoverage(p,changed),/unsupported value/i);assert.deepEqual(p,before);
  assert.throws(()=>editObjectiveCoverage(p,{kind:'links',activityId:a.id,links:[]}),/unsupported value/i);assert.deepEqual(p,before);
 }
});

test('absent and blank objectives retain legacy review and output behavior; current text still counts exactly',()=>{
 for(const replacement of [undefined,'','  ', 'Perform a skill']){
  const p=fixture(),m=p.modules[0];if(replacement===undefined)delete m.objective;else m.objective=replacement;
  const row=coverageReport(p).rows[0];assert.equal(row.invalid,false);assert.equal(row.blank,replacement===undefined||!replacement.trim());assert.equal(row.roles.practice.current.length,replacement==='Perform a skill'?1:0);
  assert.doesNotThrow(()=>workbookPreview(p));assert.doesNotThrow(()=>worksheetFiles(p));assert.doesNotThrow(()=>searchOutline(p));
 }
});

test('workbook file generation validates malformed course roots before checking objective fields',()=>{
 for(const p of [null,{}, {schema:1,id:'guide',title:'Course'}, ...[null,{},'modules',0,false].map(modules=>({schema:1,id:'guide',title:'Course',modules}))]){
  const before=structuredClone(p);assert.throws(()=>worksheetFiles(p),e=>e instanceof Error&&!(e instanceof TypeError)&&/unsupported OpenStage Guide data/i.test(e.message));assert.deepEqual(p,before);
 }
});
