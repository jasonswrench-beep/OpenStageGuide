// Editing the guide's styles (src/guide-styles-edit.mjs): changes laid over the program's own definitions, and styles of the author's own.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {stylesPart} from '../src/guide-styles.mjs';
import {cleanSpec,cleanStyleEdits,applyStyleEdits,listStyles,readSpec,idForName,specCss,STYLE_LIMITS,mergeSpec,diffSpec,contrastRatio,contrastNeeded,previewStyles} from '../src/guide-styles-edit.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const base=stylesPart();
const edit=(overrides={},custom=[])=>applyStyleEdits(base,{overrides,custom},io);
const find=(xml,id,mine=['Callout','Strong2','Loud','Plainer'])=>listStyles(xml,io,new Set(mine)).find(s=>s.id===id);

test('with no edits the styles part comes back as it was, byte for byte',()=>{
  assert.equal(applyStyleEdits(base,{overrides:{},custom:[]},io),base);
  assert.equal(applyStyleEdits(base,undefined,io),base);
  assert.equal(applyStyleEdits(base,{overrides:{OSGCueBody:{}},custom:[]},io),base,'an empty change is no change');
});

test('a built-in style takes the settings laid over it, and nothing else about it changes',()=>{
  const before=find(base,'OSGCueBody');
  const out=edit({OSGCueBody:{font:{name:'Georgia',size:13,color:'#1f3a5f',bold:true,italic:false,underline:true,strike:true},paragraph:{alignment:'justify',spaceBefore:3,spaceAfter:9,lineSpacing:1.5,leftIndent:36,firstLineIndent:-18,keepWithNext:true},shading:'#fff4ce'}});
  const after=find(out,'OSGCueBody');
  assert.deepEqual(after.own.font,{name:'Georgia',size:13,color:'#1F3A5F',bold:true,italic:false,strike:true,underline:true});
  assert.deepEqual(after.own.paragraph,{alignment:'justify',spaceBefore:3,spaceAfter:9,lineSpacing:1.5,leftIndent:36,firstLineIndent:-18,keepWithNext:true});
  assert.equal(after.own.shading,'#FFF4CE');
  assert.equal(after.name,before.name);assert.equal(after.basedOn,before.basedOn);
  const others=listStyles(base,io).filter(s=>s.id!=='OSGCueBody');
  const othersAfter=listStyles(out,io).filter(s=>s.id!=='OSGCueBody');
  assert.deepEqual(othersAfter.map(s=>[s.id,s.own]),others.map(s=>[s.id,s.own]),'every other style is as it was');
});

test('a style\'s properties stay in the order Word\'s schema gives, and the edit is the same however it is made up',()=>{
  const out=edit({OSGCueBody:{font:{size:12,color:'#000000',bold:true,name:'Arial'},paragraph:{alignment:'center',spaceAfter:6,leftIndent:10}}});
  const style=io.parse(out).documentElement;
  const el=Array.from(style.childNodes).filter(n=>n.nodeType===1&&n.localName==='style').find(s=>s.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','styleId')==='OSGCueBody');
  const names=e=>Array.from(e.childNodes).filter(n=>n.nodeType===1).map(n=>n.localName);
  const rPr=Array.from(el.childNodes).find(n=>n.localName==='rPr'),pPr=Array.from(el.childNodes).find(n=>n.localName==='pPr');
  assert.deepEqual(names(rPr),['rFonts','b','bCs','color','sz','szCs']);
  assert.ok(names(pPr).indexOf('spacing')<names(pPr).indexOf('ind')&&names(pPr).indexOf('ind')<names(pPr).indexOf('jc'));
  const ps=names(el);
  assert.ok(ps.indexOf('qFormat')<ps.indexOf('pPr')&&ps.indexOf('pPr')<ps.indexOf('rPr'),'name, based on, next, qFormat, then pPr, then rPr');
  assert.equal(edit({OSGCueBody:{paragraph:{spaceAfter:6,alignment:'center',leftIndent:10},font:{name:'Arial',bold:true,color:'#000000',size:12}}}),out);
});

test('taking an edit away puts the style back exactly: the edits are laid on the definitions each time, not written into them',()=>{
  const out=edit({OSGCueBody:{font:{size:20}}});
  assert.notEqual(out,base);
  assert.equal(applyStyleEdits(base,{overrides:{},custom:[]},io),base);
  const twice=edit({OSGCueBody:{font:{size:20}}});
  assert.equal(twice,out,'the same edits on the same definitions give the same part');
});

test('a style of the author\'s own is a real Word style: custom, named, based on another, next is itself, in the style gallery',()=>{
  const out=edit({},[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'OSGCueBody',spec:{shading:'#FFF4CE',font:{italic:true,size:11}}}]);
  const s=find(out,'Callout');
  assert.deepEqual([s.name,s.type,s.basedOn,s.custom],['Callout','paragraph','OSGCueBody',true]);
  assert.deepEqual(s.own,{font:{italic:true,size:11},shading:'#FFF4CE'});
  assert.match(out,/<w:style [^>]*w:customStyle="1"[^>]*w:styleId="Callout"|<w:style [^>]*w:styleId="Callout"[^>]*w:customStyle="1"/);
  assert.match(out,/<w:name w:val="Callout"\/><w:basedOn w:val="OSGCueBody"\/><w:next w:val="Callout"\/><w:qFormat\/>/);
  assert.equal(s.effective.font.size,11,'its own size');
  assert.ok(s.effective.paragraph.leftIndent>0,'and what it is based on counts for the rest');
  const ch=edit({},[{id:'Strong2',name:'Strong 2',type:'character',basedOn:'DefaultParagraphFont',spec:{font:{bold:true}}}]);
  assert.equal(find(ch,'Strong2').type,'character');
  assert.ok(!/w:val="Strong2"\/><w:qFormat/.test(ch.replace('<w:next w:val="Strong2"/>','')),'a character style has no next');
});

test('a style based on a style of the author\'s own is allowed; one based on nothing that exists, or with an id that is taken, is refused',()=>{
  const custom=[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'OSGCueBody',spec:{}},{id:'Loud',name:'Loud callout',type:'paragraph',basedOn:'Callout',spec:{font:{bold:true}}}];
  assert.equal(find(edit({},custom),'Loud').effective.font.bold,true);
  assert.throws(()=>edit({},[{id:'X1',name:'X',type:'paragraph',basedOn:'NoSuchStyle',spec:{}}]),/based on a style that is not there/);
  assert.throws(()=>edit({},[{id:'OSGCueBody',name:'Mine',type:'paragraph',basedOn:'Normal',spec:{}}]),/already a style with the id/);
});

test('settings that Word or a page cannot use are refused, with what is wrong, and nothing is half-written',()=>{
  const bad=spec=>()=>cleanSpec(spec);
  assert.throws(bad({font:{size:5}}),/font size must be a number from 6 to 96/);
  assert.throws(bad({font:{size:97}}),/from 6 to 96/);
  assert.throws(bad({font:{color:'red'}}),/six hexadecimal digits/);
  assert.throws(bad({font:{color:'#12345'}}),/six hexadecimal digits/);
  assert.throws(bad({font:{name:'Bad"Font'}}),/font name/);
  assert.throws(bad({font:{name:'<script>'}}),/font name/);
  assert.throws(bad({font:{name:'x'.repeat(65)}}),/font name/);
  assert.throws(bad({font:{bold:'yes'}}),/on or off/);
  assert.throws(bad({paragraph:{alignment:'diagonal'}}),/Alignment must be/);
  assert.throws(bad({paragraph:{spaceAfter:-1}}),/Space after must be/);
  assert.throws(bad({paragraph:{lineSpacing:9}}),/Line spacing must be/);
  assert.throws(bad({paragraph:{keepWithNext:1}}),/on or off/);
  assert.throws(bad({shading:'beige'}),/shading color/);
  assert.throws(bad(null),/not in the right form/);
  assert.throws(bad([]),/not in the right form/);
  assert.deepEqual(cleanSpec({font:{name:'Georgia',extra:1},unknown:true}),{font:{name:'Georgia'}},'unknown settings are dropped');
  assert.throws(()=>cleanStyleEdits({overrides:JSON.parse('{"__proto__":{},"OSGNotes":{}}')}),/id this program cannot use/,'a key that could reach the prototype is not a style id');
  assert.throws(()=>cleanStyleEdits({overrides:{'a b':{}}}),/id this program cannot use/);
  assert.throws(()=>cleanStyleEdits({custom:[{id:'A1',name:'a<b',type:'paragraph',basedOn:'Normal'}]}),/name is up to/);
  assert.throws(()=>cleanStyleEdits({custom:[{id:'A1',name:'  ',type:'paragraph',basedOn:'Normal'}]}),/name is up to/);
  assert.throws(()=>cleanStyleEdits({custom:[{id:'A1',name:'A',type:'table',basedOn:'Normal'}]}),/for paragraphs or for characters/);
  assert.throws(()=>cleanStyleEdits({custom:[{id:'A1',name:'Same',type:'paragraph',basedOn:'Normal'},{id:'A2',name:'same',type:'paragraph',basedOn:'Normal'}]}),/already a style called/);
  const many=Array.from({length:STYLE_LIMITS.custom+1},(_,i)=>({id:`S${i}`,name:`Style ${i}`,type:'paragraph',basedOn:'Normal'}));
  assert.throws(()=>cleanStyleEdits({custom:many}),/up to 60 styles/);
});

test('a name becomes an id nothing else has, whatever the name holds',()=>{
  assert.equal(idForName('Callout',[]),'Callout');
  assert.equal(idForName('My callout! 2',[]),'Mycallout2');
  assert.equal(idForName('Callout',['callout']),'Callout2','compared without regard to case');
  assert.equal(idForName('Callout',['Callout','Callout2']),'Callout3');
  assert.match(idForName('123',[]),/^Style123$/);
  assert.match(idForName('日本語',[]),/^Style/,'a name with no letters Word ids can use still gets an id');
  assert.ok(idForName('x'.repeat(200),[]).length<=44);
  for(const n of ['<b>','a b','Ünïcode','---'])assert.match(idForName(n,[]),/^[A-Za-z][A-Za-z0-9]*$/);
});

test('the styles read back as a list, with each style\'s own settings and what it comes to; hidden styles are left out',()=>{
  const list=listStyles(base,io);
  assert.ok(list.length>=30);
  const normal=list.find(s=>s.id==='Normal');
  assert.equal(normal.defaultStyle,true);
  const h=list.find(s=>s.id==='Heading1');
  assert.equal(h.name,'heading 1');
  assert.ok(h.effective.font.size>=14);
  assert.ok(list.every(s=>s.id&&s.name&&['paragraph','character'].includes(s.type)));
  assert.equal(list.some(s=>s.id==='DefaultParagraphFont'),false,'a style that is hidden is not listed');
});

test('a spec is drawn as CSS a page can use: points become pixels, bold and italic and decoration say what they mean',()=>{
  assert.deepEqual(specCss({font:{name:'Georgia',size:12,color:'#112233',bold:true,italic:false,underline:true,strike:true},paragraph:{alignment:'center',spaceBefore:6,spaceAfter:12,lineSpacing:1.5,leftIndent:36,firstLineIndent:-18},shading:'#FFF4CE'}),
    {fontFamily:'"Georgia", system-ui, sans-serif',fontSize:'16px',color:'#112233',fontWeight:'700',fontStyle:'normal',textDecoration:'underline line-through',textAlign:'center',marginTop:'8px',marginBottom:'16px',lineHeight:'1.8',marginLeft:'48px',textIndent:'-24px',backgroundColor:'#FFF4CE'});
  assert.deepEqual(specCss({}),{});
  assert.equal(specCss({font:{underline:false,strike:false}}).textDecoration,'none');
});

test('readSpec reads a style the way it was written, so what an edit wrote is what comes back',()=>{
  const spec={font:{name:'Verdana',size:10.5,color:'#336699',bold:false,italic:true,underline:false,strike:false},paragraph:{alignment:'right',spaceBefore:0,spaceAfter:4,lineSpacing:1.15,leftIndent:0,firstLineIndent:12,keepWithNext:false,keepTogether:true},shading:'#EEEEEE'};
  const out=edit({OSGNotes:spec});
  const el=Array.from(io.parse(out).documentElement.childNodes).find(n=>n.nodeType===1&&n.localName==='style'&&n.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','styleId')==='OSGNotes');
  const back=readSpec(el);
  assert.deepEqual(back.font,spec.font);
  assert.equal(back.shading,spec.shading);
  for(const [k,v] of Object.entries(spec.paragraph))if(k!=='keepWithNext'||v)assert.deepEqual(back.paragraph[k],v,k);
});

test('two specs laid together keep each group, the second over the first',()=>{
  assert.deepEqual(mergeSpec({font:{name:'A',size:10},paragraph:{alignment:'left'},shading:'#111111'},{font:{size:12,bold:true},shading:'#222222'}),{font:{name:'A',size:12,bold:true},paragraph:{alignment:'left'},shading:'#222222'});
  assert.deepEqual(mergeSpec({},{}),{});
  assert.deepEqual(mergeSpec(undefined,{font:{size:9}}),{font:{size:9}});
});

test('what the author changed is only what is not what the style came to: the rest goes on following the style it is based on',()=>{
  const was={font:{name:'Georgia',size:12,color:'#111111'},paragraph:{spaceAfter:6,alignment:'left'},shading:'#EEEEEE'};
  assert.deepEqual(diffSpec(was,was),{});
  assert.deepEqual(diffSpec(was,{font:{name:'Georgia',size:14,color:'#111111'},paragraph:{spaceAfter:6,alignment:'center'},shading:'#EEEEEE'}),{font:{size:14},paragraph:{alignment:'center'}});
  assert.deepEqual(diffSpec({},{font:{bold:true},shading:'#FFF4CE'}),{font:{bold:true},shading:'#FFF4CE'});
  assert.deepEqual(diffSpec(was,{font:{size:undefined}}),{},'a setting that is not given is not a change');
});

test('contrast is measured as the web accessibility guidelines measure it, and large text needs less',()=>{
  assert.equal(contrastRatio('#000000','#FFFFFF'),21);
  assert.equal(contrastRatio('#FFFFFF','#FFFFFF'),1);
  assert.equal(contrastRatio('#767676','#FFFFFF'),4.54,'the lightest grey that passes on white');
  assert.equal(contrastRatio('#777777','#FFFFFF')<4.5,true);
  assert.equal(contrastRatio('red','#FFFFFF'),null);
  assert.equal(contrastNeeded(11,false),4.5);
  assert.equal(contrastNeeded(14,true),3);assert.equal(contrastNeeded(14,false),4.5);assert.equal(contrastNeeded(18,false),3);
  assert.equal(contrastNeeded(undefined,undefined),4.5);
});

test('the guide page draws a changed style over its usual look, and a style of the author\'s from the style it is based on',()=>{
  const edits={overrides:{OSGCueBody:{font:{size:14,color:'#112233'}}},custom:[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'OSGCueBody',spec:{shading:'#FFF4CE'}},{id:'Loud',name:'Loud callout',type:'paragraph',basedOn:'Callout',spec:{font:{bold:true}}},{id:'Plainer',name:'Plainer',type:'paragraph',basedOn:'Normal',spec:{font:{italic:true}}}]};
  const xml=applyStyleEdits(base,edits,io);
  const list=listStyles(xml,io,new Set(edits.custom.map(c=>c.id)));
  const p=previewStyles(edits,list);
  assert.deepEqual(p.cssFor('OSG Cue Body'),{fontSize:'18.67px',color:'#112233'},'a changed built-in style: only what was changed');
  assert.equal(p.cssFor('OSG Notes'),null,'a style nobody changed has no changes to draw');
  assert.deepEqual(p.cssFor('Callout'),{fontSize:'18.67px',color:'#112233',backgroundColor:'#FFF4CE'},'a style of the author\'s starts from what its base was changed to');
  assert.equal(p.cssFor('Loud callout').fontWeight,'700');assert.equal(p.cssFor('Loud callout').backgroundColor,'#FFF4CE','through two styles');
  assert.equal(p.baseOf('Callout'),'OSG Cue Body');assert.equal(p.baseOf('Loud callout'),'OSG Cue Body');
  assert.equal(p.baseOf('Plainer'),'Normal');
  assert.equal(p.baseOf('OSG Cue Body'),'OSG Cue Body');
  assert.equal(p.baseOf('No such style'),null);
});

test('borders are written as Word borders, read back, and drawn',()=>{
  const out=edit({OSGNotes:{borders:[{location:'left',color:'#112233',width:3},{location:'top',color:'#445566',width:0.5}]}});
  const own=find(out,'OSGNotes').own;
  assert.deepEqual(own.borders,[{location:'top',color:'#445566',width:0.5},{location:'left',color:'#112233',width:3}]);
  assert.match(out,/<w:left w:val="single" w:sz="24" w:space="4" w:color="112233"\/>/);
  assert.equal(specCss({borders:own.borders}).borderLeft,'4px solid #112233');
  assert.throws(()=>cleanSpec({borders:[{location:'middle',color:'#112233',width:1}]}),/left, right, top or bottom/);
  assert.throws(()=>cleanSpec({borders:[{location:'left',color:'#112233',width:1},{location:'left',color:'#112233',width:2}]}),/once/);
  assert.throws(()=>cleanSpec({borders:[{location:'left',color:'blue',width:1}]}),/border color/);
  assert.throws(()=>cleanSpec({borders:[{location:'left',color:'#112233',width:9}]}),/width must be/);
});
