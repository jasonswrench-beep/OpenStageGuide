import test from 'node:test';
import assert from 'node:assert/strict';
import {randomUUID} from 'node:crypto';
import {FAMILY_BRAND_FIELDS} from '../src/workbook-family-flow.mjs';
import {BRAND_PRESET_KEY,readBrandPresetCatalog,writeBrandPresetCatalog,nextBrandPresetCatalog,validateBrandPresetCatalog} from '../src/workbook-brand-presets.mjs';

const brand=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(key=>[key,'']));
const memory=()=>{
  const data=new Map();
  return {data,getItem:key=>data.get(key)??null,setItem:(key,value)=>data.set(key,value)};
};
const create=(state,name='Example College',values=brand())=>nextBrandPresetCatalog(state.catalog,'create',
  {id:randomUUID(),uuid:randomUUID(),name,brand:values});

test('local catalog persists only authored overrides and reopens with exact readback',()=>{
  const storage=memory(),initial=readBrandPresetCatalog(storage),raw=brand();raw.organizationName='Example College';raw.text='#1A1A1A';
  const next=create(initial,'College',raw),saved=writeBrandPresetCatalog(storage,initial.raw,next);
  assert.equal(storage.data.get(BRAND_PRESET_KEY),saved.raw);
  assert.deepEqual(readBrandPresetCatalog(storage),saved);
  assert.deepEqual(Object.keys(saved.catalog.presets[0].brand),FAMILY_BRAND_FIELDS);
  assert.equal(JSON.stringify(saved).includes('effective'),false);
  assert.equal(JSON.stringify(saved).includes('project'),false);
  const overwrite=nextBrandPresetCatalog(saved.catalog,'overwrite',{id:next.presets[0].id,uuid:randomUUID(),name:'College updated',brand:raw});
  assert.equal(writeBrandPresetCatalog(storage,saved.raw,overwrite).catalog.presets.length,1);
});

test('unavailable, quota and uncertain writes never erase the last valid catalog',()=>{
  const storage=memory(),initial=readBrandPresetCatalog(storage),saved=writeBrandPresetCatalog(storage,null,create(initial));
  const next=create(saved,'Second');
  assert.throws(()=>writeBrandPresetCatalog({getItem:()=>{throw Error('blocked');}},saved.raw,next),/could not be read/);
  assert.throws(()=>writeBrandPresetCatalog({getItem:storage.getItem,setItem:()=>{throw Error('quota');}},saved.raw,next),/outcome is unconfirmed/);
  assert.equal(storage.data.get(BRAND_PRESET_KEY),saved.raw);
  let calls=0;
  assert.throws(()=>writeBrandPresetCatalog({getItem:key=>++calls===1?saved.raw:'different',setItem:storage.setItem},saved.raw,next),/outcome is unconfirmed/);
  assert.deepEqual(saved.catalog.presets.length,1);
});

test('stale tabs, duplicate identities, unknown versions and altered JSON refuse without writes',()=>{
  const storage=memory(),initial=readBrandPresetCatalog(storage),saved=writeBrandPresetCatalog(storage,null,create(initial));
  let writes=0;const guarded={getItem:storage.getItem,setItem:()=>writes++};
  assert.throws(()=>writeBrandPresetCatalog(guarded,null,create(initial,'Stale')),/changed in another session/);
  assert.equal(writes,0);
  const row=saved.catalog.presets[0],duplicate={...saved.catalog,revision:randomUUID(),presets:[row,{...row}]};
  assert.throws(()=>validateBrandPresetCatalog(duplicate),/duplicated/);
  assert.throws(()=>validateBrandPresetCatalog({...saved.catalog,version:2}),/unsupported/);
  assert.throws(()=>validateBrandPresetCatalog({...saved.catalog,extra:'hidden'}),/unsupported/);
  assert.throws(()=>validateBrandPresetCatalog(Object.assign(Object.create({version:1}),saved.catalog)),/plain record/);
  const accessor={...row.brand};Object.defineProperty(accessor,'text',{enumerable:true,get(){throw Error('read');}});
  assert.throws(()=>validateBrandPresetCatalog({...saved.catalog,presets:[{...row,brand:accessor}]}),/visible own data/);
  storage.data.set(BRAND_PRESET_KEY,'{"format":"openstage-guide-brand-presets","version":2,"revision":"empty","presets":[]}');
  assert.throws(()=>readBrandPresetCatalog(storage),/unsupported/);
  storage.data.set(BRAND_PRESET_KEY,saved.raw.slice(0,-1)+',"presets":[]}');
  assert.throws(()=>readBrandPresetCatalog(storage),/malformed|noncanonical/);
});

test('catalog count, names, values and overflow are bounded',()=>{
  const storage=memory(),initial=readBrandPresetCatalog(storage),saved=writeBrandPresetCatalog(storage,null,create(initial));
  assert.throws(()=>create(saved,'example college'),/duplicated/);
  assert.throws(()=>create(saved,' '),/name is invalid/);
  const invalid=brand();invalid.text='X'.repeat(81);
  assert.throws(()=>create(saved,'Other',invalid),/bounded raw override/);
  const rows=Array.from({length:20},(_,i)=>({id:randomUUID(),name:`Preset ${i}`,brand:brand()}));
  const full={...saved.catalog,presets:rows};
  assert.throws(()=>nextBrandPresetCatalog(full,'create',{id:randomUUID(),uuid:randomUUID(),name:'Too many',brand:brand()}),/limit/);
  storage.data.set(BRAND_PRESET_KEY,'X'.repeat(32769));
  assert.throws(()=>readBrandPresetCatalog(storage),/oversized/);
});

test('catalog rejects scalar coercion, array getters, hidden fields and custom iteration without reading them',()=>{
  const baseline=readBrandPresetCatalog(memory()).catalog;
  let reads=0,iterations=0;
  const revision={toString(){reads++;return randomUUID();}};
  assert.throws(()=>validateBrandPresetCatalog({...baseline,revision}),/revision is invalid/);assert.equal(reads,0);
  const getter=[];Object.defineProperty(getter,'0',{enumerable:true,get(){reads++;return null;}});getter.length=1;
  assert.throws(()=>validateBrandPresetCatalog({...baseline,presets:getter}),/own data/);assert.equal(reads,0);
  const custom=[];Object.defineProperty(custom,Symbol.iterator,{value:function*(){iterations++;yield null;}});
  assert.throws(()=>validateBrandPresetCatalog({...baseline,presets:custom}),/unsupported fields/);assert.equal(iterations,0);
  const hidden=[];Object.defineProperty(hidden,'private',{value:'canary'});
  assert.throws(()=>validateBrandPresetCatalog({...baseline,presets:hidden}),/unsupported fields/);
  const hole=new Array(1);assert.throws(()=>validateBrandPresetCatalog({...baseline,presets:hole}),/dense list|unsupported fields/);
  const options={id:randomUUID(),name:'Valid',brand:brand(),uuid:randomUUID()};
  Object.defineProperty(options,'name',{enumerable:true,get(){reads++;return 'Getter';}});
  assert.throws(()=>nextBrandPresetCatalog(baseline,'create',options),/visible own data/);assert.equal(reads,0);
  const inherited=Object.assign(Object.create({id:randomUUID()}),{name:'X',brand:brand(),uuid:randomUUID()});
  assert.throws(()=>nextBrandPresetCatalog(baseline,'create',inherited),/plain record/);
});
