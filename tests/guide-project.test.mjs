import test from 'node:test';
import assert from 'node:assert/strict';
import {PROJECT_FORMAT,PROJECT_VERSION,LIMITS,ProjectError,cleanGuide,serializeProject,parseProject} from '../src/guide-project.mjs';

const para=(text,style='Normal',xml='')=>({xml,text,style});
const guide=(over={})=>({title:'Reframing conflict',paragraphs:[para('Reframing conflict','OSG Guide Title','<w:p xmlns:w="w"><w:r><w:t>Reframing conflict</w:t></w:r></w:p>'),para('Module 1: Naming it','OSG Module')],cursor:1,ownCues:['Pause here'],look:'{"name":"calm"}',slideMemory:'{"slides":3}',versionMemory:null,forms:null,design:null,sessions:null,applied:[[{name:'OSG Module',size:16}]],headerFooter:[{header:'A'}],media:{},styles:{overrides:{},custom:[]},...over});
const refused=(fn,pattern)=>assert.throws(fn,e=>e instanceof ProjectError&&pattern.test(e.message),String(pattern));

test('a guide written to a project file reads back exactly',()=>{
  const g=guide();
  const text=serializeProject(g,{savedAt:'2026-11-13T10:00:00.000Z'});
  const back=parseProject(text);
  assert.deepEqual(back.guide,g);
  assert.equal(back.version,PROJECT_VERSION);
  assert.equal(back.savedAt,'2026-11-13T10:00:00.000Z');
  assert.ok(text.endsWith('\n'));
  assert.equal(JSON.parse(text).format,PROJECT_FORMAT);
});

test('writing a project does not change the guide it is given',()=>{
  const g=guide(),before=JSON.stringify(g);
  serializeProject(g);
  assert.equal(JSON.stringify(g),before);
});

test('things that are not projects are refused in plain words, and nothing is half-read',()=>{
  refused(()=>parseProject(null),/not an OpenStage Guide project/);
  refused(()=>parseProject(''),/not an OpenStage Guide project/);
  refused(()=>parseProject('{not json'),/not readable/);
  refused(()=>parseProject('[1,2]'),/not an OpenStage Guide project/);
  refused(()=>parseProject(JSON.stringify({format:'something-else',version:1,guide:guide()})),/not an OpenStage Guide project/);
});

test('a project from a newer program is refused and says to update; an unusable version is refused',()=>{
  const make=version=>JSON.stringify({format:PROJECT_FORMAT,version,guide:guide()});
  assert.throws(()=>parseProject(make(PROJECT_VERSION+1)),e=>e.code==='newer'&&/newer OpenStage Guide/.test(e.message));
  refused(()=>parseProject(make(0)),/no usable version/);
  refused(()=>parseProject(make('1')),/no usable version/);
  refused(()=>parseProject(make(1.5)),/no usable version/);
});

test('a project whose guide is the wrong shape is refused',()=>{
  const wrap=g=>JSON.stringify({format:PROJECT_FORMAT,version:1,guide:g});
  refused(()=>parseProject(wrap(null)),/does not hold a guide/);
  refused(()=>parseProject(wrap({})),/no list of paragraphs/);
  refused(()=>parseProject(wrap({paragraphs:'x'})),/no list of paragraphs/);
  refused(()=>parseProject(wrap({paragraphs:[1]})),/Paragraph 1 .* not in the right form/);
  refused(()=>parseProject(wrap({paragraphs:[{text:5}]})),/paragraph 1 is not text/);
  refused(()=>parseProject(wrap({paragraphs:[],ownCues:'x'})),/own cue names/);
  refused(()=>parseProject(wrap({paragraphs:[],look:5})),/look is not text/);
  refused(()=>parseProject(wrap({paragraphs:[],applied:{}})),/look settings/);
});

test('limits are enforced: too many paragraphs, a paragraph that is too long, too many cues',()=>{
  refused(()=>cleanGuide({paragraphs:Array.from({length:LIMITS.paragraphs+1},()=>({}))}),/more paragraphs/);
  refused(()=>cleanGuide({paragraphs:[{text:'x'.repeat(LIMITS.words+1)}]}),/too long/);
  refused(()=>cleanGuide({paragraphs:[],ownCues:Array.from({length:LIMITS.cues+1},()=>'a')}),/own cue names/);
  refused(()=>cleanGuide({paragraphs:[],title:'x'.repeat(LIMITS.title+1)}),/title is too long/);
});

test('missing optional parts take their empty value, unknown parts are dropped, the cursor is kept inside the guide',()=>{
  const g=cleanGuide({paragraphs:[para('a'),para('b')],cursor:99,extra:'ignored',paragraphs_note:1});
  assert.deepEqual(g,{title:'',paragraphs:[para('a'),para('b')],cursor:1,ownCues:[],look:null,slideMemory:null,versionMemory:null,forms:null,design:null,sessions:null,applied:[],headerFooter:[],media:{},styles:{overrides:{},custom:[]}});
  assert.equal(cleanGuide({paragraphs:[para('a')],cursor:-5}).cursor,-1);
  assert.equal(cleanGuide({paragraphs:[para('a')],cursor:'2'}).cursor,-1);
  assert.equal('extra' in g,false);
});

test('settings are copied as plain data: keys that reach an object\'s prototype are dropped, odd values are refused',()=>{
  const evil=JSON.parse('{"paragraphs":[],"applied":[{"__proto__":{"polluted":1},"constructor":2,"ok":3}]}');
  const g=cleanGuide(evil);
  assert.deepEqual(g.applied,[{ok:3}]);
  assert.equal({}.polluted,undefined);
  refused(()=>cleanGuide({paragraphs:[],applied:[{n:Infinity}]}),/not a number/);
  refused(()=>cleanGuide({paragraphs:[],applied:[{f(){}}]}),/kind this program does not know/);
  let deep={};const top=deep;for(let i=0;i<30;i++){deep.a={};deep=deep.a;}
  refused(()=>cleanGuide({paragraphs:[],applied:[top]}),/nested too deeply/);
});

test('a project carries the author\'s pictures: PNG, JPEG and GIF as base64 under ids the paragraphs name; anything else is refused whole',()=>{
  const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const media={rIdPicOne:{mime:'image/png',base64:png},rIdPicTwo:{mime:'image/jpeg',base64:'AAAA'},rIdPicThree:{mime:'image/gif',base64:'R0lG'}};
  const back=parseProject(serializeProject(guide({media}))).guide;
  assert.deepEqual(back.media,media,'what was stored is what comes back');
  assert.deepEqual(cleanGuide({...guide(),media:undefined}).media,{},'a project from before pictures has none');
  const bad=value=>()=>cleanGuide(guide({media:value}));
  assert.throws(bad({'x y':{mime:'image/png',base64:png}}),/picture in the project is not in the right form/,'an id that is not one the paragraphs could name');
  assert.throws(bad({rIdIcon_x:{mime:'image/png',base64:png}}),/not in the right form/,'icons are not the author\'s pictures');
  assert.throws(bad({rIdPicA:{mime:'image/svg+xml',base64:'AAAA'}}),/not in the right form/,'SVG can carry script, so it is not a picture here');
  assert.throws(bad({rIdPicA:{mime:'text/html',base64:'AAAA'}}),/not in the right form/);
  assert.throws(bad({rIdPicA:{mime:'image/png',base64:'not base64!'}}),/not in the right form/);
  assert.throws(bad({rIdPicA:{mime:'image/png',base64:'AAA'}}),/not in the right form/,'base64 comes in fours');
  assert.throws(bad({rIdPicA:{mime:'image/png',base64:''}}),/not in the right form/);
  assert.throws(bad({rIdPicA:'AAAA'}),/not in the right form/);
  assert.throws(bad([]),/not in the right form/);
  assert.throws(bad({rIdPicA:{mime:'image/png',base64:'A'.repeat(LIMITS.pictureBase64+4)}}),/too large/);
  const many=Object.fromEntries(Array.from({length:LIMITS.pictures+1},(_,i)=>[`rIdPic${i}`,{mime:'image/png',base64:png}]));
  assert.throws(bad(many),/more pictures than/);
  assert.equal(cleanGuide(guide({media:{rIdPicA:{mime:'image/png',base64:png,extra:'dropped'}}})).media.rIdPicA.extra,undefined,'unknown fields are dropped');
  const proto=JSON.parse('{"__proto__":{"mime":"image/png","base64":"AAAA"},"rIdPicA":{"mime":"image/png","base64":"AAAA"}}');
  assert.throws(()=>cleanGuide(guide({media:proto})),/not in the right form/,'a key that could reach the prototype is not an id');
});

test('a project carries the changes to styles and the author\'s own styles, and refuses styles that are not in the right form, naming what is wrong',()=>{
  const styles={overrides:{OSGCueBody:{font:{size:13,name:'Georgia'},paragraph:{alignment:'center'}}},custom:[{id:'Callout',name:'Callout',type:'paragraph',basedOn:'OSGCueBody',spec:{shading:'#FFF4CE'}}]};
  assert.deepEqual(parseProject(serializeProject(guide({styles}))).guide.styles,styles);
  assert.deepEqual(cleanGuide({...guide(),styles:undefined}).styles,{overrides:{},custom:[]},'a project from before styles has none');
  assert.throws(()=>cleanGuide(guide({styles:{overrides:{OSGCueBody:{font:{size:2}}}}})),/styles are not in the right form \(The font size must be a number from 6 to 96\), so it was not opened/);
  assert.throws(()=>cleanGuide(guide({styles:'x'})),/styles are not in the right form/);
  assert.throws(()=>cleanGuide(guide({styles:{custom:[{id:'X',name:'<',type:'paragraph',basedOn:'Normal'}]}})),/styles are not in the right form/);
});

test('the notes of the design desk are kept as text, left empty when there are none, and refused when larger than a project can hold',()=>{
  assert.equal(cleanGuide(guide()).design,null);
  const notes=JSON.stringify({version:1,objectives:[{id:'o1',text:'Explain the steps.'}]});
  const g=cleanGuide(guide({design:notes}));
  assert.equal(g.design,notes);
  assert.equal(parseProject(serializeProject(guide({design:notes}))).guide.design,notes);
  assert.throws(()=>cleanGuide(guide({design:'x'.repeat(LIMITS.words+1)})),ProjectError);
  assert.throws(()=>cleanGuide(guide({design:{not:'text'}})),ProjectError);
});

test('the record of the materials and decks made from the guide is kept as text, left empty when there is none, and refused when larger than a project can hold',()=>{
  assert.equal(cleanGuide(guide()).sessions,null);
  const record=JSON.stringify({version:1,outputs:[],decks:[]});
  assert.equal(cleanGuide(guide({sessions:record})).sessions,record);
  assert.equal(parseProject(serializeProject(guide({sessions:record}))).guide.sessions,record);
  assert.throws(()=>cleanGuide(guide({sessions:'x'.repeat(LIMITS.words+1)})),ProjectError);
  assert.throws(()=>cleanGuide(guide({sessions:{not:'text'}})),ProjectError);
  assert.equal(cleanGuide({...guide(),sessions:undefined}).sessions,null,'a project from before it has none');
});