import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {ICONS} from '../src/model.mjs';
import {createCueOwnership} from '../src/cue-ownership.mjs';
import {createIconJournal,validateIconJournal,proposeIconTransition} from '../src/icon-transition.mjs';
import {createDocumentStore,PENDING} from '../src/document-store.mjs';
function sample(){const p=fixture();p.cueOwnership={version:1};const a=p.modules[0].activities[0];a.cueImage=createCueOwnership(a.id,a.icon);p.unknown={private:'keep',list:[{x:1}]};a.unknown={manual:'metadata'};a.materials=[];return p;}
const checkpoint=()=>({id:crypto.randomUUID(),masterHash:'a'.repeat(64),backupHash:'b'.repeat(64)});
const journal=()=>{const p=sample();return createIconJournal(p,p.modules[0].activities[0].id,'Assessment',checkpoint());};
test('icon-only proposals preserve all unknown/course/activity data and stable ownership identity for every supplied icon',()=>{
  for(const [icon] of ICONS){const p=sample(),before=structuredClone(p),a=p.modules[0].activities[0],result=proposeIconTransition(p,a.id,icon);assert.deepEqual(p,before);if(icon===a.icon){assert.equal(result,null);continue;}
    const expected=structuredClone(p),changed=expected.modules[0].activities[0];changed.icon=icon;changed.cueImage.assetHash=result.icon.after.assetHash;expected.revision++;
    assert.deepEqual(result.project,expected);assert.deepEqual(result.base,p);assert.equal(changed.cueImage.nonce,a.cueImage.nonce);assert.equal(changed.cueImage.activityId,a.id);result.project.unknown.list[0].x=5;assert.equal(p.unknown.list[0].x,1);
  }
});
test('legacy/manual/unowned/custom ownership is ineligible and never silently adopted',()=>{
  for(const change of [p=>delete p.cueOwnership,p=>p.cueOwnership.version=2,p=>p.cueOwnership.extra=true,p=>delete p.modules[0].activities[0].cueImage,p=>p.modules[0].activities[0].cueImage.version=2,p=>p.modules[0].activities[0].cueImage.assetHash='c'.repeat(64),p=>p.modules[0].activities[0].cueImage.extra=true,p=>p.modules[0].activities[0].cueImage.nonce='unknown',p=>p.modules[0].activities[0].cueImage.activityId=crypto.randomUUID()]){
    const p=sample();change(p);const before=structuredClone(p);assert.throws(()=>proposeIconTransition(p,p.modules[0].activities[0].id,'Assessment'),e=>e.code==='OSG_ICON_PROPOSAL');assert.deepEqual(p,before);
  }
});
test('icon proposals refuse unknown targets, icons, duplicate owner tokens and revision overflow',()=>{
  const p=sample(),id=p.modules[0].activities[0].id;for(const [target,icon] of [[crypto.randomUUID(),'Assessment'],[id,'Custom'],[id,'toString']])assert.throws(()=>proposeIconTransition(p,target,icon));
  const other=structuredClone(p.modules[0].activities[0]);other.id=crypto.randomUUID();other.cueImage.activityId=other.id;p.modules[0].activities.push(other);assert.throws(()=>proposeIconTransition(p,id,'Assessment'));
  p.modules[0].activities.pop();p.revision=Number.MAX_SAFE_INTEGER;assert.throws(()=>proposeIconTransition(p,id,'Assessment'));
});
test('typed schema4 icon journals preserve full base/after/checkpoint evidence and clone caller values',()=>{
  const j=journal(),original=structuredClone(j),checked=validateIconJournal(j);assert.deepEqual(j,original);assert.deepEqual(checked,j);checked.base.unknown.list[0].x=9;assert.equal(j.base.unknown.list[0].x,1);
  const receipt=validateIconJournal({...j,phase:'committed'});assert.deepEqual(receipt.base,j.base);assert.deepEqual(receipt.project,j.project);assert.deepEqual(receipt.icon,j.icon);assert.deepEqual(receipt.checkpoint,j.checkpoint);assert.equal(receipt.txId,j.txId);
  const p=sample();assert.equal(createIconJournal(p,p.modules[0].activities[0].id,p.modules[0].activities[0].icon,checkpoint()),null);
});
test('icon journal validation rejects unrelated changes anywhere in the full proposed guide',()=>{
  for(const change of [j=>j.project.title='Changed',j=>j.project.revision++,j=>j.project.id=crypto.randomUUID(),j=>j.project.unknown.list[0].x=2,j=>j.project.modules.reverse(),j=>j.project.modules[0].title='Changed',j=>j.project.modules[0].activities[0].minutes++,j=>j.project.modules[0].activities[0].lines++,j=>j.project.modules[0].activities[0].answers='Changed',j=>j.project.modules[0].activities[0].materials.push({name:'Extra'}),j=>delete j.project.modules[0].activities[0].unknown,j=>j.project.modules[0].activities[0].cueImage.nonce='c'.repeat(32)]){
    const j=journal();j.base.modules.push({id:crypto.randomUUID(),title:'Second',activities:[]});j.project.modules.push(structuredClone(j.base.modules[1]));change(j);assert.throws(()=>validateIconJournal(j),e=>e.code==='OSG_ICON_PROPOSAL');
  }
});
test('typed icon journal rejects incomplete/mixed versions, phases, extra fields and inconsistent icon/checkpoint evidence',()=>{
  for(const change of [j=>j.schema=3,j=>j.kind='activity-settings',j=>j.phase='writing',j=>j.txId='-'.repeat(36),j=>delete j.base,j=>j.extra='unknown journal version',j=>j.checkpoint.id='old',j=>j.checkpoint.masterHash='bad',j=>j.checkpoint.extra=true,j=>j.icon.ownerTag+='x',j=>j.icon.before.icon='Assessment',j=>j.icon.after.assetHash='c'.repeat(64),j=>j.icon.after.icon=j.icon.before.icon,j=>j.icon.extra=true]){const j=journal();change(j);assert.throws(()=>validateIconJournal(j),e=>e.code==='OSG_ICON_PROPOSAL');}
});
test('existing document store continues to pause on schema4 with no durable save or Word write',async()=>{
  const j=journal(),h=makeHost(j.base);h.durable[PENDING]=structuredClone(j);h.memory[PENDING]=structuredClone(j);let writes=0;const store=createDocumentStore({settings:()=>h.settings,scan:async()=>h.records,write:async()=>{writes++;}}),before=structuredClone(h.durable);
  await assert.rejects(store.readProject(),/paused/);assert.equal(store.isBlocked(),true);assert.equal(h.saves,0);assert.equal(writes,0);assert.deepEqual(h.durable,before);assert.equal(h.runs,0);
});

test('icon journals reject non-JSON or lossy metadata instead of hiding it in equality comparisons',()=>{
  for(const change of [j=>j.project.unknown.invisible=undefined,j=>j.base.unknown.value=NaN,j=>j.base.unknown.date=new Date(),j=>j.base.unknown.list=Array(2),j=>j.project.unknown.self=j.project]){const j=journal();change(j);assert.throws(()=>validateIconJournal(j),e=>e.code==='OSG_ICON_PROPOSAL');}
});

test('sparse arrays cannot camouflage missing indexes with extra properties in a guide or journal',()=>{
  for(const target of ['base','project'])for(const shape of ['sparse-extra','dense-extra','non-index','negative-index']){
    const j=journal(),a=shape==='dense-extra'?[1]:Array(1);a[shape==='non-index'?'01':shape==='negative-index'?'-1':'extra']='concealed';j[target].unknown.list=a;
    assert.throws(()=>validateIconJournal(j),e=>e.code==='OSG_ICON_PROPOSAL');
    const p=sample();p.unknown.list=a;assert.throws(()=>proposeIconTransition(p,p.modules[0].activities[0].id,'Assessment'),e=>e.code==='OSG_ICON_PROPOSAL');
  }
});
test('invisible and accessor properties are refused without invoking getters or stripping checkpoint metadata',()=>{
  let reads=0;
  for(const target of ['object','array','index','checkpoint','journal'])for(const kind of ['symbol','nonenumerable','accessor']){
    const j=journal(),value=target==='journal'?j:target==='checkpoint'?j.checkpoint:['array','index'].includes(target)?j.base.unknown.list:j.project.unknown;
    if(kind==='symbol')value[Symbol('hidden')]='secret';
    else Object.defineProperty(value,target==='index'?'0':'hidden',kind==='accessor'?{enumerable:true,get(){reads++;return 'secret';}}:{value:'secret',enumerable:false});
    assert.throws(()=>validateIconJournal(j),e=>e.code==='OSG_ICON_PROPOSAL');
    if(target==='checkpoint'){const p=sample();assert.throws(()=>createIconJournal(p,p.modules[0].activities[0].id,'Assessment',value),e=>e.code==='OSG_ICON_PROPOSAL');}
  }
  assert.equal(reads,0);
});
test('scalar icon identifiers and ownership/checkpoint evidence never accept array or object coercion',()=>{
  let coercions=0;const object=value=>({toString(){coercions++;return value;}});
  const p=sample(),id=p.modules[0].activities[0].id;
  for(const wrap of [value=>[value],object,value=>new String(value)]){
    assert.throws(()=>proposeIconTransition(p,id,wrap('Assessment')),e=>e.code==='OSG_ICON_PROPOSAL');
    assert.throws(()=>proposeIconTransition(p,wrap(id),'Assessment'),e=>e.code==='OSG_ICON_PROPOSAL');
    for(const path of [['txId'],['checkpoint','id'],['checkpoint','masterHash'],['checkpoint','backupHash'],['icon','id'],['icon','ownerTag'],['icon','before','icon'],['icon','after','icon'],['icon','before','assetHash'],['icon','after','assetHash'],['base','id'],['base','modules',0,'id'],['base','modules',0,'activities',0,'id'],['base','modules',0,'activities',0,'icon'],['base','modules',0,'activities',0,'cueImage','nonce'],['base','modules',0,'activities',0,'cueImage','activityId'],['base','modules',0,'activities',0,'cueImage','assetHash']]){
      const j=journal();let parent=j;for(const key of path.slice(0,-1))parent=parent[key];const key=path.at(-1);parent[key]=wrap(parent[key]);assert.throws(()=>validateIconJournal(j),e=>e.code==='OSG_ICON_PROPOSAL');
    }
  }
  assert.equal(coercions,0);
});
test('supported dense JSON metadata survives proposals and constructed after-state is validated',()=>{
  const p=sample();p.unknown={list:[null,true,false,0,-1,1.25,'text',[],{},[1,2]],nested:{a:'keep',b:null}};const before=structuredClone(p),id=p.modules[0].activities[0].id,j=createIconJournal(p,id,'Assessment',checkpoint());assert.deepEqual(j.base,p);assert.deepEqual(j.project.unknown,p.unknown);assert.equal(typeof j.project.modules[0].activities[0].icon,'string');assert.deepEqual(p,before);
  const malformed=structuredClone(j);malformed.project.modules[0].activities[0].icon=['Assessment'];assert.throws(()=>validateIconJournal(malformed),e=>e.code==='OSG_ICON_PROPOSAL');
});
