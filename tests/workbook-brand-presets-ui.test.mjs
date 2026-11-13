import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {fixture,makeHost} from './helpers/mock-office.mjs';
import {dom,settle} from './helpers/app-dom.mjs';
import {BRAND_PRESET_KEY} from '../src/workbook-brand-presets.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';

let sequence=0;
const storage=()=>{
  const data=new Map();return {data,getItem:key=>data.get(key)??null,setItem:(key,value)=>data.set(key,value)};
};
async function app(localStorage=storage()){
  const h=makeHost(fixture()),ui=dom(),n=id=>ui.nodes.get(id);
  Object.assign(globalThis,{Office:h.Office,Word:h.Word,document:ui.document,location:{search:''},JSZip,localStorage});
  await import(`../src/app.mjs?preset-ui=${++sequence}`);await settle();await ui.routes.find(button=>button.dataset.view==='output').onclick();await settle();
  return {h,n,localStorage};
}
const click=async(n,id)=>{n(id).onclick();await settle();};
const set=(n,id,value,event='oninput')=>{n(id).value=value;n(id)[event]?.();};
const preview=async n=>{await click(n,'family-preview');assert.equal(n('family-reviewed').disabled,false);n('family-reviewed').checked=true;n('family-reviewed').onchange();};

test('explicit save survives a panel reopen and load applies brand to current family only',async()=>{
  const store=storage(),{n}=await app(store);
  await preview(n);set(n,'family-brand-organizationName','Example College');set(n,'family-brand-bodyPt','13');
  set(n,'family-preset-name','College brand');await click(n,'family-preset-save-new');
  assert.equal(JSON.parse(store.data.get(BRAND_PRESET_KEY)).presets.length,1);
  const id=n('family-preset-list').value;assert.ok(id);assert.equal(n('family-reviewed').checked,false);
  const reopened=await app(store),m=reopened.n;
  assert.equal(m('family-preset-list').children.length,2);
  set(m,'family-template','field-guide','onchange');
  set(m,'family-brand-organizationName','Unsaved draft');set(m,'family-preset-name','Name work');
  m('family-preset-list').value=id;m('family-preset-list').onchange();
  await click(m,'family-preset-load');assert.equal(m('family-preset-confirm').hidden,false);
  assert.equal(m('family-brand-organizationName').value,'Unsaved draft');
  await click(m,'family-preset-keep');assert.equal(m('family-brand-organizationName').value,'Unsaved draft');
  await click(m,'family-preset-load');await click(m,'family-preset-discard');
  assert.equal(m('family-template').value,'field-guide');
  assert.equal(m('family-brand-organizationName').value,'Example College');
  assert.equal(m('family-brand-bodyPt').value,'13');
  assert.equal(m('family-preset-name').value,'College brand');
  assert.equal(m('family-reviewed').checked,false);await preview(m);
  assert.match(m('family-preview-content').textContent,/Field Guide|field-guide/);
});

test('reset and overwrite are explicit; failed/stale confirmation keeps raw drafts and review',async()=>{
  const store=storage(),{n}=await app(store);
  set(n,'family-brand-organizationName','First');set(n,'family-preset-name','Brand A');await click(n,'family-preset-save-new');
  const id=n('family-preset-list').value;
  set(n,'family-brand-organizationName','Second');set(n,'family-preset-name','Brand B');
  await click(n,'family-preset-overwrite');assert.equal(n('family-preset-confirm').hidden,false);
  assert.equal(JSON.parse(store.data.get(BRAND_PRESET_KEY)).presets[0].brand.organizationName,'First');
  await click(n,'family-preset-discard');assert.equal(JSON.parse(store.data.get(BRAND_PRESET_KEY)).presets[0].brand.organizationName,'Second');
  await preview(n);
  await click(n,'family-preset-reset');set(n,'family-brand-organizationName','Late raw');
  await click(n,'family-preset-discard');assert.equal(n('family-brand-organizationName').value,'Late raw');
  assert.match(n('family-preview-content').textContent,/Refresh and review/);assert.equal(n('family-reviewed').checked,false);
  await preview(n);await click(n,'family-preset-reset');await click(n,'family-preset-keep');
  assert.equal(n('family-brand-organizationName').value,'Late raw');
  await click(n,'family-preset-reset');await click(n,'family-preset-discard');
  assert.equal(n('family-brand-organizationName').value,'');assert.equal(n('family-preset-name').value,'Brand B');
  assert.equal(n('family-preset-list').value,id);assert.equal(JSON.parse(store.data.get(BRAND_PRESET_KEY)).presets.length,1);
});

test('stale catalog, invalid current policy and storage failure refuse without changing active brand',async()=>{
  const store=storage(),{n}=await app(store);
  set(n,'family-brand-organizationName','Saved');set(n,'family-preset-name','Brand');await click(n,'family-preset-save-new');
  set(n,'family-brand-organizationName','Current raw');
  await click(n,'family-preset-load');const before=n('family-brand-organizationName').value;
  store.data.set(BRAND_PRESET_KEY,store.data.get(BRAND_PRESET_KEY).replace('"Saved"','"Externally changed"'));
  await click(n,'family-preset-discard');assert.equal(n('family-brand-organizationName').value,before);
  assert.match(n('family-preset-status').textContent,/changed|reload/i);
  const old=JSON.parse(store.data.get(BRAND_PRESET_KEY));old.presets[0].brand.text='#FFFFFF';store.data.set(BRAND_PRESET_KEY,JSON.stringify(old));
  await click(n,'family-preset-reload');n('family-preset-list').value=old.presets[0].id;n('family-preset-list').onchange();
  await click(n,'family-preset-load');assert.equal(n('family-preset-confirm').hidden,true);
  assert.match(n('family-preset-status').textContent,/current rules|contrast/i);
  assert.equal(n('family-brand-organizationName').value,before);
});

test('unavailable local storage is local to presets and leaves Workbook review usable',async()=>{
  const {n}=await app({getItem(){throw Error('denied');},setItem(){throw Error('denied');}});
  assert.match(n('family-preset-status').textContent,/could not be read/);
  assert.equal(n('family-preset-save-new').disabled,true);
  await preview(n);assert.equal(n('family-apply').disabled,false);
  set(n,'family-brand-organizationName','Unsaved raw');set(n,'family-preset-name','Name work');
  await click(n,'family-preset-reset');await click(n,'family-preset-discard');
  assert.equal(n('family-brand-organizationName').value,'');assert.equal(n('family-preset-name').value,'Name work');
});

test('one authored organization brand can be loaded into each family without changing the family',async()=>{
  const store=storage(),{n}=await app(store);
  set(n,'family-brand-organizationName','Campus brand');set(n,'family-preset-name','Campus');await click(n,'family-preset-save-new');
  const id=n('family-preset-list').value;
  for(const family of WORKBOOK_FAMILIES){
    set(n,'family-template',family.id,'onchange');set(n,'family-brand-organizationName','Temporary');
    n('family-preset-list').value=id;n('family-preset-list').onchange();
    await click(n,'family-preset-load');await click(n,'family-preset-discard');
    assert.equal(n('family-template').value,family.id);
    assert.equal(n('family-brand-organizationName').value,'Campus brand');
    await preview(n);assert.match(n('family-preview-content').textContent,/Campus brand/);
  }
});

test('quota and unconfirmed storage writes preserve the active review and typed name',async()=>{
  const store=storage(),{n}=await app(store);
  await preview(n);const view=n('family-preview-content').textContent;
  set(n,'family-preset-name','Pending name');
  store.setItem=()=>{throw Error('quota');};await click(n,'family-preset-save-new');
  assert.match(n('family-preset-status').textContent,/outcome is unconfirmed/);
  assert.equal(n('family-preset-name').value,'Pending name');assert.equal(n('family-preview-content').textContent,view);
  assert.equal(n('family-reviewed').checked,true);assert.equal(store.data.size,0);
  store.setItem=(key,value)=>store.data.set(key,value);
  await click(n,'family-preset-save-new');assert.match(n('family-preset-status').textContent,/Reload saved presets/);
  await click(n,'family-preset-reload');
  let reads=0;store.getItem=key=>++reads===1?null:'unconfirmed';
  await click(n,'family-preset-save-new');assert.match(n('family-preset-status').textContent,/outcome is unconfirmed/);
  assert.equal(n('family-preset-name').value,'Pending name');assert.equal(n('family-preview-content').textContent,view);
  assert.equal(n('family-reviewed').checked,true);
});

test('nested output hierarchy excludes local preset controls from the workbook review fingerprint',async()=>{
  const {n}=await app();
  // Model the actual nested parent/child textContent, which the flat mock omits.
  n('workbook-family-section').append(n('family-brand-controls'),n('family-brand-status'),n('family-preview-content'),n('family-review-label'));
  n('view-output').append(n('workbook-family-section'),n('family-preset-controls'));
  await preview(n);const reviewed=n('family-preview-content').textContent;
  set(n,'family-preset-name','Default brand');await click(n,'family-preset-save-new');
  n('family-reviewed').onchange();assert.equal(n('family-apply').disabled,false);
  assert.equal(n('family-preview-content').textContent,reviewed);
  await click(n,'family-preset-load');await click(n,'family-preset-keep');
  n('family-reviewed').onchange();assert.equal(n('family-apply').disabled,false);
  await click(n,'family-preset-load');await click(n,'family-preset-discard');
  assert.equal(n('family-reviewed').checked,true);assert.equal(n('family-apply').disabled,false);
  await click(n,'family-preset-reset');await click(n,'family-preset-discard');
  assert.equal(n('family-reviewed').checked,true);assert.equal(n('family-apply').disabled,false);
  set(n,'family-brand-organizationName','Actually changed');
  assert.equal(n('family-reviewed').checked,false);assert.equal(n('family-apply').disabled,true);
});

test('base language is session output configuration outside version-one brand presets',async()=>{
  const store=storage(),{n}=await app(store);
  set(n,'family-base-language','fr-FR','onchange');
  set(n,'family-brand-organizationName','Campus');set(n,'family-preset-name','Campus preset');
  await click(n,'family-preset-save-new');
  const catalog=JSON.parse(store.data.get(BRAND_PRESET_KEY));
  assert.equal(Object.keys(catalog.presets[0].brand).length,16);
  assert.equal(JSON.stringify(catalog).includes('baseLanguage'),false);
  set(n,'family-base-language','de-DE','onchange');set(n,'family-brand-organizationName','Temporary');
  await click(n,'family-preset-load');await click(n,'family-preset-discard');
  assert.equal(n('family-base-language').value,'de-DE');assert.equal(n('family-brand-organizationName').value,'Campus');
  await click(n,'family-preset-reset');await click(n,'family-preset-discard');
  assert.equal(n('family-base-language').value,'de-DE');assert.equal(n('family-brand-organizationName').value,'');
});
