import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture} from './helpers/mock-office.mjs';
import {editActivityMaterials,addActivity} from '../src/model.mjs';
import {serializeMaterialsDraft,readMaterialsDraft,restoreMaterialsDraft,MAX_MATERIAL_DRAFT_BYTES} from '../src/materials-draft.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
const item={item:'Handout',quantity:2,unit:'copies',use:'consumable',notes:'PRIVATE PREPARATION'};
const editor={item:'Unfinished supply',quantity:'not decided',unit:'',use:'consumable',notes:'PRIVATE UNSAVED\nSecond line'};
test('private draft roundtrip restores complete items and unfinished raw editor without changing saved materials',()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,before=JSON.stringify(p),text=serializeMaterialsDraft(p,id,[item],editor),r=restoreMaterialsDraft(text,p,id);
  assert.deepEqual(r.draft.items,[item]);assert.deepEqual(r.unsavedEntry,editor);assert.deepEqual(r.draft.restored.savedItems,[]);assert.equal(JSON.stringify(p),before);assert.equal(serializeMaterialsDraft(r.draft.project,id,r.draft.items,r.unsavedEntry),text);
});
test('different saved revisions are shown for review without applying or rebasing silently',()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,text=serializeMaterialsDraft(p,id,[item],editor),current=editActivityMaterials(p,id,[{...item,quantity:5}]),r=restoreMaterialsDraft(text,current,id);
  assert.equal(r.draft.restored.recordedRevision,p.revision);assert.equal(r.draft.project.revision,current.revision);assert.equal(r.draft.restored.savedItems[0].quantity,5);assert.equal(r.draft.items[0].quantity,2);assert.equal(current.modules[0].activities[0].materials[0].quantity,5);
});
test('draft restore rejects wrong course/activity, corrupt/unsupported/oversized and unknown data atomically',()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,text=serializeMaterialsDraft(p,id,[item],editor),before=JSON.stringify(p);assert.throws(()=>restoreMaterialsDraft(text,fixture(),id));assert.throws(()=>restoreMaterialsDraft(text,p,'missing'));
  for(const mutate of [x=>x.version=2,x=>x.settings={private:true},x=>x.items[0].quantity=-1,x=>x.items.push({...x.items[0]}),x=>x.unsavedEntry.secret='private',x=>x.unsavedEntry.item='line\nbreak',x=>x.unsavedEntry.quantity='x'.repeat(129),x=>x.revision=-1]){const r=JSON.parse(text);mutate(r);assert.throws(()=>readMaterialsDraft(JSON.stringify(r)));}
  assert.throws(()=>readMaterialsDraft('{bad'));assert.throws(()=>readMaterialsDraft(' '.repeat(MAX_MATERIAL_DRAFT_BYTES+1)));assert.equal(JSON.stringify(p),before);
});
test('restored requirements and raw preparation records never enter learner package parts',()=>{
  const p=fixture(),id=p.modules[0].activities[0].id,r=restoreMaterialsDraft(serializeMaterialsDraft(p,id,[item],editor),p,id);p.privateMaterialRecovery=r;for(const part of Object.values(worksheetFiles(p)))assert.doesNotMatch(part,/PRIVATE|UNSAVED|Unfinished supply|not decided|Handout/);
});
