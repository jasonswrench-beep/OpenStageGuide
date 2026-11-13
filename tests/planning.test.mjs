import test from 'node:test';
import assert from 'node:assert/strict';
import {course,addModule,addActivity} from '../src/model.mjs';
import {buildAgenda,reviewCourse} from '../src/planning.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
function fixture(){let p=addModule(course('Planning test'),'Module one','Practice a response.').next;return addActivity(p,p.modules[0].id,{title:'Practice',minutes:12.5,participant:'Write your response.',facilitator:'PRIVATE FACILITATOR',answers:'SECRET ANSWER',lines:2}).next;}
test('agenda carries fractional minutes across module boundaries and midnight',()=>{
  let p=fixture();p=addModule(p,'Module two','Reflect.').next;p=addActivity(p,p.modules[1].id,{title:'Debrief',participant:'Reflect.',minutes:3.5}).next;
  const a=buildAgenda(p,'23:50');assert.equal(a.rows[0].end,'Day 2 · 00:02:30');assert.equal(a.rows[1].start,a.rows[0].end);assert.equal(a.end,'Day 2 · 00:06');assert.equal(a.minutes,16);
  assert.equal(p.modules[0].activities[0].minutes,12.5);
});
test('agenda never includes instructor content or private metadata',()=>{
  const p=fixture();p.secret='PRIVATE METADATA';const a=buildAgenda(p);
  assert.doesNotMatch(JSON.stringify(a),/PRIVATE|SECRET/);assert.deepEqual(Object.keys(a.rows[0]),['moduleId','moduleTitle','activityId','title','start','end','minutes']);
});
test('agenda rejects invalid stored durations and impossible start times',()=>{
  const p=fixture();for(const s of ['24:00','9:00','12:60',''])assert.throws(()=>buildAgenda(p,s),/start time/);
  p.modules[0].activities[0].minutes=NaN;assert.throws(()=>buildAgenda(p),/durations/);
});
test('review finds blank required text, incomplete prompts and missing timing without modifying the course',()=>{
  const p=fixture(),m=p.modules[0],a=m.activities[0];m.objective='';a.participant=' ';a.facilitator='[Insert directions]';a.minutes=0;
  const before=JSON.stringify(p),issues=reviewCourse(p);assert.equal(JSON.stringify(p),before);
  assert.ok(issues.some(x=>x.severity==='error'&&/participant/.test(x.message)&&x.id===a.id));
  assert.ok(issues.some(x=>/learning objective/.test(x.message)&&x.id===m.id));assert.ok(issues.some(x=>/unfinished/.test(x.message)));assert.ok(issues.some(x=>/No teaching time/.test(x.message)));
  assert.deepEqual(reviewCourse(fixture()),[]);
});
test('review reports empty modules and invalid stored durations separately',()=>{
  const p=fixture();p.modules[0].activities[0].minutes=-2;
  assert.ok(reviewCourse(p).some(x=>x.severity==='error'&&/duration/.test(x.message)));
  assert.ok(reviewCourse(course('Empty')).some(x=>/module/.test(x.message)));
});
test('workbook export stops on blank learner instructions but allows optional-content suggestions',()=>{
  const p=fixture();p.modules[0].objective='';assert.doesNotThrow(()=>worksheetFiles(p));
  p.modules[0].activities[0].participant='';assert.throws(()=>worksheetFiles(p),/Review Guide/);
});
