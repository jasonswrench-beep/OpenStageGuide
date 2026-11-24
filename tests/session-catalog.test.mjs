// The session catalog (src/session-catalog.mjs) held to the product spec's reconciliation rules: every template's total is the sum of its blocks, every activity's timed
// steps add up to its minutes, and nothing in it is unfinished or borrowed.
import test from 'node:test';
import assert from 'node:assert/strict';
import {ACTIVITIES,TEMPLATES,PHASES,FAMILIES,CATALOG_VERSION} from '../src/session-catalog.mjs';
import {MODES} from '../src/guide-facts.mjs';

const byId=new Map(ACTIVITIES.map(a=>[a.id,a]));
const sum=(list,f)=>list.reduce((n,x)=>n+f(x),0);

test('the catalog holds 16 templates, 46 activities and 102 agenda blocks, each ID once',()=>{
  assert.equal(CATALOG_VERSION,1);
  assert.equal(TEMPLATES.length,16);
  assert.equal(ACTIVITIES.length,46);
  assert.equal(sum(TEMPLATES,t=>t.rows.length),102);
  assert.equal(new Set(TEMPLATES.map(t=>t.id)).size,16);
  assert.equal(new Set(ACTIVITIES.map(a=>a.id)).size,46);
  for(const t of TEMPLATES)assert.match(t.id,/^T\d\d$/);
  for(const a of ACTIVITIES)assert.match(a.id,/^A\d\d$/);
});

test('ten templates and 35 activities are launch; six and eleven are the expansion pack',()=>{
  assert.deepEqual([TEMPLATES.filter(t=>t.pack==='Launch').length,TEMPLATES.filter(t=>t.pack==='Expansion').length],[10,6]);
  assert.deepEqual([ACTIVITIES.filter(a=>a.pack==='Launch').length,ACTIVITIES.filter(a=>a.pack==='Expansion').length],[35,11]);
  assert.deepEqual(TEMPLATES.slice(0,10).map(t=>t.pack),Array(10).fill('Launch'));
});

test('every template total equals the sum of its blocks',()=>{
  for(const t of TEMPLATES){
    assert.ok(t.rows.length>=5,`${t.id} has a short agenda`);
    for(const id of t.rows)assert.ok(byId.has(id),`${t.id} names ${id}, which is not an activity`);
    assert.equal(sum(t.rows,id=>byId.get(id).minutes),t.minutes,`${t.id} ${t.title}`);
  }
});

test('every activity\'s timed steps equal its default duration',()=>{
  for(const a of ACTIVITIES){
    assert.ok(a.steps.length>=1,a.id);
    assert.equal(sum(a.steps,s=>s.min),a.minutes,`${a.id} ${a.title}`);
    for(const s of a.steps){assert.ok(Number.isInteger(s.min)&&s.min>0,a.id);assert.ok(s.text.trim().length>3,a.id);}
  }
});

test('every activity carries what the library and the workbook need, in words an author can use',()=>{
  for(const a of ACTIVITIES){
    assert.ok(PHASES.includes(a.phase),`${a.id} phase ${a.phase}`);
    assert.ok(MODES.includes(a.mode),`${a.id} mode ${a.mode}`);
    assert.ok(['Launch','Expansion'].includes(a.pack),a.id);
    for(const k of ['title','grouping','output','access'])assert.ok(a[k].trim(),`${a.id} has no ${k}`);
    assert.ok(a.materials.length>=1&&a.materials.every(m=>m.trim()),`${a.id} materials`);
    assert.ok(a.task.length>=1&&a.task.every(m=>m.trim()),`${a.id} has no participant task`);
    assert.equal(a.version,1);
  }
  // The two Break and set-up blocks may have no debrief; every other activity asks something afterwards.
  for(const a of ACTIVITIES.filter(a=>a.phase!=='Break'))assert.ok(a.debrief.length>=1,`${a.id} has no debrief prompt`);
  assert.ok(ACTIVITIES.every(a=>a.mode==='Break'?a.phase==='Break':true));
});

test('every template says who it is for, how many, what it makes and why its blocks come in that order',()=>{
  for(const t of TEMPLATES){
    assert.ok(FAMILIES.includes(t.family),`${t.id} family`);
    for(const k of ['title','audience','purpose','output','prerequisites','rationale'])assert.ok(t[k].trim(),`${t.id} has no ${k}`);
    assert.ok(t.group.min>=1&&t.group.max>t.group.min,`${t.id} group`);
    assert.ok(t.materials.length>=2,`${t.id} materials`);
    assert.equal(t.version,1);
  }
  assert.equal(new Set(TEMPLATES.map(t=>t.family)).size,4);
});

test('the sample of the spec holds: Skill Practice Workshop from 09:00 to 10:00 with the peer feedback at 09:35',()=>{
  const t=TEMPLATES.find(x=>x.id==='T01');
  assert.equal(t.title,'Skill Practice Workshop');
  let at=9*60;const starts=[];
  for(const id of t.rows){starts.push([id,`${String(Math.floor(at/60)).padStart(2,'0')}:${String(at%60).padStart(2,'0')}`]);at+=byId.get(id).minutes;}
  assert.deepEqual(starts.map(s=>s[1]),['09:00','09:05','09:17','09:35','09:47','09:52']);
  assert.equal(starts.find(s=>s[0]==='A14')[1],'09:35');
  assert.equal(at,10*60);
});

test('nothing in the catalog is unfinished, marked for checking, or names another product',()=>{
  const text=JSON.stringify({ACTIVITIES,TEMPLATES});
  // The marker is the capitalized word a reviewer leaves; the lower-case word is ordinary English ("a cause to verify").
  const found=[/\bVERIFY\b/,/\b(?:TODO|TBD)\b/i,/\[(?:insert|add|replace)/i,/metodic/i,/https?:\/\//i,/⏎|\\t/].map(re=>{const m=re.exec(text);return m?`${re}: ...${text.slice(Math.max(0,m.index-40),m.index+40)}...`:null;}).filter(Boolean);
  assert.deepEqual(found,[]);
});

test('five activities are in the library only, in no template, and every other one is in at least one',()=>{
  const used=new Map();
  for(const t of TEMPLATES)for(const id of t.rows)used.set(id,(used.get(id)||0)+1);
  // The catalog is the one definition of each activity: a template names it by ID and cannot change its minutes or phase.
  assert.equal(used.size,41);
  assert.deepEqual([...byId.keys()].filter(id=>!used.has(id)).sort(),['A17','A26','A31','A33','A44']);
});
