import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {GUIDE_LOOKS,cleanLook,parseLook,resolveGuideLook,styleSpecs,normalSpec,lookSpecs,headerPackage,footerPackage,DEFAULT_LOOK} from '../src/guide-looks.mjs';
import {styleDefinitions} from '../src/guide-styles.mjs';
import {createWordAdapter} from '../src/guide-word.mjs';
import {startGuide,makeContext} from '../src/guide-ooxml.mjs';
import {started,click} from './helpers/guide-harness.mjs';

const theme=id=>resolveGuideLook({id}).theme;
const spec=(specs,name)=>specs.find(s=>s.name===name);

test('a look is cleaned into a complete, safe set of choices',()=>{
  assert.deepEqual(cleanLook(null),DEFAULT_LOOK);
  assert.deepEqual(cleanLook({id:'nonsense',accent:'zzz',organization:'  Acme\n  Learning ',header:1,copyright:'x'.repeat(300),revised:' 3 Nov 2026 '}),
    {id:'clearline',accent:'',organization:'Acme Learning',header:true,copyright:'x'.repeat(120),revised:'3 Nov 2026'});
  assert.equal(cleanLook({id:'field-guide',accent:'164b6b'}).accent,'#164B6B','a missing # is added and the letters are made capitals');
  assert.equal(parseLook('not json').id,'clearline');
  assert.deepEqual(parseLook(JSON.stringify({id:'practice-lab',accent:'#12594D'})).id,'practice-lab');
  assert.deepEqual(GUIDE_LOOKS.map(l=>l.id),['clearline','editorial-study','practice-lab','field-guide','reflection-journal','storyboard']);
});

test('every look turns into a theme, and an accent too pale to read as heading text is refused',()=>{
  for(const l of GUIDE_LOOKS){
    const r=resolveGuideLook({id:l.id});
    assert.equal(r.ok,true,l.id);
    assert.match(r.theme.navy,/^[0-9A-F]{6}$/);
    assert.ok(r.theme.headingFont&&r.theme.bodyFont,l.id);
  }
  assert.equal(theme('editorial-study').headingFont,'Georgia');
  const pale=resolveGuideLook({id:'clearline',accent:'#FFFF99'});
  assert.equal(pale.ok,false);
  assert.match(pale.issues[0],/too light to read as heading text.*needs at least 4\.5 to 1/);
  assert.equal(resolveGuideLook({id:'clearline',accent:'#000080'}).ok,true);
  // The accent is also the colour of the labels on the beige boxes, and the guide is printed in black and white: it must read there as well.
  const onBeige=resolveGuideLook({id:'clearline',accent:'#8C64B9'});
  assert.equal(onBeige.ok,false,'readable enough on white, but not on the beige box');
  assert.match(onBeige.issues[0],/too light to read on the beige boxes.*needs at least 4.5 to 1/);
  const printed=resolveGuideLook({id:'clearline',accent:'#7B6090'});
  assert.equal(printed.ok,false,'fine in colour on both, too faint as the grey a black-and-white printer makes of it');
  assert.match(printed.issues[0],/too faint when the guide is printed in black and white.*4.3 to 1.*needs at least 4.5 to 1/);
  assert.equal(resolveGuideLook({id:'clearline',organization:'Acme'}).effective.organizationName,'Acme');
});

test('the style settings come from the guide\'s own definitions, one for each paragraph style',()=>{
  const t=theme('clearline'),specs=styleSpecs(t,DOMParser);
  const count=(styleDefinitions(t).match(/<w:style w:type="paragraph"/g)||[]).length;
  assert.equal(specs.length,count);
  const h1=spec(specs,'Heading 1');
  assert.deepEqual(h1.font,{name:'Aptos Display',size:16,color:'#164B6B',bold:true});
  assert.deepEqual(h1.paragraph,{spaceBefore:18,spaceAfter:6,keepWithNext:true,keepTogether:true});
  const notes=spec(specs,'OSG Facilitator Notes');
  assert.equal(notes.shading,'#F1EBDD');
  assert.deepEqual(notes.borders,[{location:'Left',color:'#164B6B',width:'Pt300'}]);
  assert.equal(notes.paragraph.leftIndent,120);
  const cue=spec(specs,'OSG Cue');
  assert.deepEqual(cue.paragraph,{spaceBefore:6,leftIndent:120,firstLineIndent:-120,keepTogether:true});
  assert.deepEqual(spec(specs,'OSG Cue Facilitator').borders,[{location:'Left',color:'#164B6B',width:'Pt150'}]);
  assert.deepEqual(spec(specs,'OSG Time').borders,[{location:'Bottom',color:'#164B6B',width:'Pt100'}]);
  assert.equal(spec(specs,'OSG Rating Box').paragraph.alignment,'Centered');
  assert.equal(spec(specs,'OSG Cue Label')?.name,undefined,'character styles are not in this list');
  // A body style based on another says only what it changes; a style based on a heading says what that heading says (real Word,
  // 2026-11-06: Word writes the heading values of the guide's first look into OSG Module and the others, so a later look must restate them).
  assert.deepEqual(spec(specs,'OSG Module').font,h1.font);
  assert.deepEqual(spec(specs,'OSG Module').paragraph,h1.paragraph);
  assert.deepEqual(spec(specs,'OSG Lesson').font,spec(specs,'Heading 2').font);
  assert.deepEqual(spec(specs,'OSG Assessment').font,h1.font);
  assert.deepEqual(spec(specs,'OSG Guide Title').font,spec(specs,'Title').font);
  assert.deepEqual(spec(specs,'OSG Guide Title').paragraph,spec(specs,'Title').paragraph);
});

test('every body style carries the body font and size of the look itself, and every heading style the look\'s heading font',()=>{
  // Word writes the old font into a style it imports, so the look has to say the new one on each body style.
  const specs=lookSpecs(theme('editorial-study'),DOMParser);
  for(const name of ['OSG Cue','OSG Cue Facilitator','OSG Facilitator Notes','OSG Time','OSG Question','OSG Table Cell','OSG Overview']){
    assert.equal(spec(specs,name).font.name,'Aptos',name);assert.equal(spec(specs,name).font.size,12,name);
  }
  for(const name of ['OSG Module','OSG Lesson','OSG Assessment','OSG Guide Title','Heading 1','Title'])assert.notEqual(spec(specs,name).font.name,'Aptos',name+' keeps its heading font');
  for(const name of ['OSG Module','OSG Lesson','OSG Assessment','OSG Guide Title'])assert.equal(spec(specs,name).font.name,'Georgia',name+' says the look\'s heading font itself');
  assert.equal(spec(specs,'OSG Module').font.color,spec(specs,'Heading 1').font.color);
  assert.equal(spec(specs,'OSG Module').font.size,spec(specs,'Heading 1').font.size);
});

test('the body text, font, size and line spacing come through the Normal style first',()=>{
  const t=theme('clearline'),n=normalSpec(t);
  assert.deepEqual(n,{name:'Normal',font:{name:'Aptos',size:12},paragraph:{lineSpacing:15,spaceAfter:6},shading:null,borders:[]});
  const all=lookSpecs(t,DOMParser);
  assert.equal(all[0].name,'Normal');
  assert.equal(all.filter(s=>s.name==='Normal').length,1);
  assert.notEqual(spec(lookSpecs(theme('editorial-study'),DOMParser),'Heading 1').font.name,spec(all,'Heading 1').font.name,'another look changes the heading font');
});

test('the header carries the title and the footer carries who, the date and Page X of Y',()=>{
  const h=headerPackage({title:'Handling <conflict> & more'});
  assert.match(h,/Handling &lt;conflict&gt; &amp; more/);
  const f=footerPackage({organization:'Acme',copyright:'© 2026 Acme',revised:'3 Nov 2026'});
  assert.match(f,/Acme · © 2026 Acme · Revised 3 Nov 2026/);
  assert.match(f,/ PAGE /);assert.match(f,/ NUMPAGES /);
  assert.doesNotMatch(footerPackage({}).replace(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g,''),/Revised|undefined/);
  for(const p of [h,f])assert.equal(new DOMParser({errorHandler:()=>{}}).parseFromString(p,'application/xml').documentElement.localName,'package');
});

// ---- the Word adapter, against a stand-in for Word's style objects
function fakeStyles({existing=['Normal','Title','Heading 1','Heading 2','OSG Module','OSG Cue','OSG Facilitator Notes'],refuse=()=>false}={}){
  const log=[],pending=[],setCalls=[];
  const recorder=(name,group)=>new Proxy({},{set(t,prop,value){log.push({name,group,prop,value});if(refuse(name,group,prop))pending.push(Object.assign(new Error('refused'),{code:'InvalidArgument'}));return true;}});
  const section=()=>{
    const bodyFor=which=>({insertOoxml:(pkg,where)=>{setCalls.push({which,where,pkg});if(refuse('section',which,'insertOoxml'))pending.push(Object.assign(new Error('refused'),{code:'GeneralException'}));},clear:()=>setCalls.push({which,cleared:true})});
    return {getHeader:()=>bodyFor('header'),getFooter:()=>bodyFor('footer')};
  };
  const ctx={
    document:{
      getStyles:()=>({getByNameOrNullObject:name=>({isNullObject:!existing.includes(name),load(){},font:recorder(name,'font'),paragraphFormat:recorder(name,'paragraphFormat'),shading:recorder(name,'shading'),borders:{getByLocation:loc=>recorder(name,`border ${loc}`)}})}),
      sections:{getFirst:section}
    },
    sync:async()=>{if(pending.length)throw pending.shift();}
  };
  return {api:{run:fn=>fn(ctx)},log,setCalls};
}
const adapterFor=fake=>createWordAdapter({Word:fake.api,fetchFn:async()=>({ok:true,arrayBuffer:async()=>new Uint8Array([137,80,78,71]).buffer}),DOMParserClass:DOMParser});

test('applyLook changes the styles the document has, skips the ones it does not, and carries on past a refusal',async()=>{
  const fake=fakeStyles({refuse:(name,group,prop)=>name==='OSG Facilitator Notes'&&group==='border Left'});
  const out=await adapterFor(fake).applyLook(lookSpecs(theme('clearline'),DOMParser));
  assert.deepEqual(out.applied.sort(),['Heading 1','Heading 2','Normal','OSG Cue','OSG Facilitator Notes','OSG Module','Title'].sort());
  assert.ok(out.skipped.includes('OSG Time')&&out.skipped.includes('OSG Question'),'styles the document has not made yet are left for the next insert');
  assert.deepEqual(out.refused,[{style:'OSG Facilitator Notes',what:'left border',why:'InvalidArgument'}]);
  const h1=fake.log.filter(l=>l.name==='Heading 1'&&l.group==='font');
  assert.deepEqual(Object.fromEntries(h1.map(l=>[l.prop,l.value])),{name:'Aptos Display',size:16,color:'#164B6B',bold:true});
  assert.ok(fake.log.some(l=>l.name==='OSG Facilitator Notes'&&l.group==='shading'&&l.prop==='backgroundPatternColor'&&l.value==='#F1EBDD'),'the shading was still set after the border was refused');
  assert.ok(fake.log.some(l=>l.group==='border Left'&&l.prop==='width'&&l.value==='Pt300'));
});

test('a style setting Word does not offer is reported, not thrown',async()=>{
  const fake=fakeStyles({refuse:(name,group)=>group==='shading'});
  const out=await adapterFor(fake).applyLook(lookSpecs(theme('clearline'),DOMParser));
  assert.ok(out.refused.some(r=>r.style==='OSG Facilitator Notes'&&r.what==='shading'));
  assert.ok(out.applied.includes('OSG Facilitator Notes'),'its font, spacing and border were still changed');
});

test('the header and footer are written to the first section, taken away on request, and left alone when not mentioned',async()=>{
  const fake=fakeStyles(),a=adapterFor(fake);
  assert.deepEqual(await a.setHeaderFooter({header:headerPackage({title:'T'}),footer:footerPackage({})}),{done:['header','footer'],failed:[]});
  assert.deepEqual(fake.setCalls.map(c=>[c.which,c.where]),[['header','Replace'],['footer','Replace']]);
  await a.setHeaderFooter({header:null,footer:undefined});
  assert.deepEqual(fake.setCalls.slice(2),[{which:'header',cleared:true}]);
  const bad=fakeStyles({refuse:(n,g)=>g==='footer'});
  assert.deepEqual(await adapterFor(bad).setHeaderFooter({header:'<x/>',footer:'<y/>'}),{done:['header'],failed:[{what:'footer',why:'GeneralException'}]});
});

test('the look is kept as a custom document property, and styles an insert makes come from it',async()=>{
  const store=new Map();
  const ctx={document:{properties:{customProperties:{
    getItemOrNullObject:k=>({isNullObject:!store.has(k),value:store.get(k),load(){}}),add:(k,v)=>{store.set(k,v);}}}},sync:async()=>{}};
  const a=createWordAdapter({Word:{run:fn=>fn(ctx)},fetchFn:async()=>({}),DOMParserClass:DOMParser});
  assert.equal(await a.getLook(),null);
  await a.setLook('{"id":"field-guide"}');
  assert.equal(store.get('osg.look'),'{"id":"field-guide"}');
  assert.equal(await a.getLook(),'{"id":"field-guide"}');
  // The package an insert builds uses the theme the adapter was told about.
  const {buildPackage}=await import('../src/guide-ooxml.mjs');
  const c=makeContext(),body=startGuide({title:'T',standard:false},c);
  assert.doesNotMatch(buildPackage(body,c,{}),/AB12CD/);
  assert.match(buildPackage(body,c,{},{theme:{navy:'AB12CD'}}),/w:color w:val="AB12CD"/);
});

// ---- the pane
const settingsOf=a=>JSON.parse(a.adapter.lookState().lookText);
const lookRadio=(a,id)=>a.n('guide-looks').children.map(l=>l.children[0]).find(i=>i.value===id);
// The test DOM has no radio groups, so choosing one unchecks the others by hand.
const chooseLook=(a,id)=>{for(const l of a.n('guide-looks').children)l.children[0].checked=l.children[0].value===id;};

test('a guide with no look says so; applying one tells the adapter, changes the styles and remembers it in the document',async()=>{
  const a=await started();
  assert.match(a.n('guide-look-now').textContent,/no look yet/);
  chooseLook(a,'editorial-study');a.type('guide-look-accent','#543D67');a.type('guide-look-org','Acme Learning');
  await click(a,'guide-look-apply');
  assert.equal(a.err(),'');
  const state=a.adapter.lookState();
  assert.equal(state.theme.headingFont,'Georgia');assert.equal(state.theme.navy,'543D67');
  assert.equal(state.applied.length,1);assert.equal(state.applied[0][0].name,'Normal');
  assert.deepEqual(settingsOf(a),{id:'editorial-study',accent:'#543D67',organization:'Acme Learning',header:false,copyright:'',revised:''});
  assert.match(a.n('guide-look-now').textContent,/This guide uses the Editorial Study look, accent color #543D67\./);
  assert.match(a.n('guide-status').textContent,/The Editorial Study look is applied\./);
  assert.equal(state.headerFooter.length,0,'no header or footer was asked for, so none was touched');
});

test('a look the pane cannot use is refused with the reason and changes nothing',async()=>{
  const a=await started();
  chooseLook(a,'clearline');a.type('guide-look-accent','#FFFF99');
  await click(a,'guide-look-apply');
  assert.match(a.err(),/too light to read as heading text/);
  assert.equal(a.adapter.lookState().applied.length,0);assert.equal(a.adapter.lookState().lookText,null);
});

test('an old "look is applied" line does not stay beside a refusal',async()=>{
  const a=await started();
  chooseLook(a,'clearline');
  await click(a,'guide-look-apply');
  assert.match(a.n('guide-look-note').textContent,/The Clearline look is applied\./);
  a.type('guide-look-accent','#FFFF99');
  await click(a,'guide-look-apply');
  assert.match(a.err(),/too light to read as heading text/);
  assert.equal(a.n('guide-look-note').textContent,'','the earlier success line is gone');
});

test('a running header and footer are added with the look and taken away when unticked',async()=>{
  const a=await started('Handling conflict');
  chooseLook(a,'clearline');a.n('guide-look-header').checked=true;a.type('guide-look-copyright','© 2026 Acme');a.type('guide-look-revised','3 Nov 2026');a.type('guide-look-org','Acme');
  await click(a,'guide-look-apply');
  assert.equal(a.err(),'');
  let [hf]=a.adapter.lookState().headerFooter;
  assert.match(hf.header,/Handling conflict/);assert.match(hf.footer,/Acme · © 2026 Acme · Revised 3 Nov 2026/);
  assert.match(a.n('guide-look-now').textContent,/with a running header and footer/);
  a.n('guide-look-header').checked=false;
  await click(a,'guide-look-apply');
  hf=a.adapter.lookState().headerFooter.at(-1);
  assert.deepEqual([hf.header,hf.footer],[null,null],'taking it away is asked for');
  // Applying again with none ticked and none present leaves them alone.
  const asked=a.adapter.lookState().headerFooter.length;
  await click(a,'guide-look-apply');
  assert.equal(a.adapter.lookState().headerFooter.length,asked,'nothing was asked of the header or footer');
});

test('what Word would not change is named, and the rest is reported as done',async()=>{
  const a=await started();
  a.adapter.refuse(()=>[{style:'OSG Facilitator Notes',what:'left border',why:'x'},{style:'OSG Facilitator Notes',what:'shading',why:'x'},{style:'Heading 1',what:'font',why:'x'}]);
  await click(a,'guide-look-apply');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/OSG Facilitator Notes \(left border, shading\); Heading 1 \(font\) could not be changed, and the rest was changed\. To finish by hand, change those styles in the Styles panel/);
});

test('a built-in style Word cannot find is named, since the look would be incomplete without it',async()=>{
  const a=await started();
  const real=a.adapter.applyLook;
  a.adapter.applyLook=async specs=>({...(await real(specs)),skipped:['Heading 1','Heading 2','OSG Time']});
  await click(a,'guide-look-apply');
  assert.match(a.n('guide-status').textContent,/Word did not find its own Heading 1, Heading 2 styles, so they were not changed\. Change them by hand in the Styles pane\./);
  assert.doesNotMatch(a.n('guide-status').textContent,/OSG Time/,'a guide style not made yet is not a problem');
});

test('a guide started with a look has it from the first insert, and one started without does not',async()=>{
  const a=(await import('./helpers/guide-harness.mjs')).pane();
  await a.p.refresh();
  a.type('guide-title','Look test');a.type('guide-start-look','practice-lab');
  await a.n('guide-start-form').onsubmit({preventDefault(){}});
  assert.equal(a.err(),'');
  assert.equal(a.adapter.lookState().theme.navy,'12594D');
  assert.equal(settingsOf(a).id,'practice-lab');
  assert.match(a.n('guide-status').textContent,/Guide started\. .*The Practice Lab look is applied\. Add a module to begin\./);
  // The pane says so too (real Word, 2026-11-06: it kept saying "no look yet" and showed Clearline for a guide started in a look).
  assert.match(a.n('guide-look-now').textContent,/This guide uses the Practice Lab look\./);
  assert.deepEqual([...a.n('guide-looks').children].filter(l=>l.children[0].checked).map(l=>l.children[0].value),['practice-lab']);
  const b=await started('No look');
  assert.equal(b.adapter.lookState().lookText,null);assert.equal(b.adapter.lookState().applied.length,0);
  assert.equal(b.n('guide-start-look').value,'','the standard styles are the starting choice until the look is proven in Word');
});

test('a guide that already has a look shows it, tells the adapter, and the workbook starts from it',async()=>{
  const a=await started();
  await a.adapter.setLook(JSON.stringify({id:'field-guide',accent:'#625014',organization:'Acme'}));
  await a.p.refresh();
  assert.match(a.n('guide-look-now').textContent,/Field Guide look, accent color #625014/);
  assert.equal(a.adapter.lookState().theme.navy,'625014');
  assert.equal(lookRadio(a,'field-guide').checked,true);assert.equal(a.n('guide-look-org').value,'Acme');
  assert.equal(a.n('guide-wb-accent').value,'#625014');assert.equal(a.n('guide-wb-org').value,'Acme');
});
