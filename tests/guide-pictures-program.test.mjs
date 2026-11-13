// The author's pictures in the program: kept by the document under ids of their own, saved in the project, given back in the package a Word
// document is made from, and sent across the message link with the block that names them.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {createProjectAdapter} from '../src/guide-project-adapter.mjs';
import {createRemoteAdapter,serveAdapter} from '../src/guide-remote.mjs';
import {parseProject,serializeProject} from '../src/guide-project.mjs';
import {pictureRun,para,run,cueBlock} from '../src/guide-ooxml.mjs';

const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const JPG='/9j/4AAQSkZJRgABAQ==';
const demo=opts=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer,...opts});
const pic=(base64=PNG,mime='image/png',alt='A cat')=>({base64,type:mime==='image/png'?'png':'jpg',mime,size:{cx:914400,cy:457200},alt,name:'cat'});
const picture=(p=pic())=>ctx=>para('OSGCueBody',pictureRun(p,ctx));
const rids=xml=>[...xml.matchAll(/r:embed="([^"]+)"/g)].map(m=>m[1]);

test('a picture a block was made with is kept under an id of its own, and the paragraph names that id',async()=>{
  const a=demo();
  await a.insert(picture(),{where:'start',spare:false});
  const [p]=a.exportState().paragraphs;
  const [id]=rids(p.xml);
  assert.match(id,/^rIdPic[a-z0-9]{1,8}$/);
  assert.notEqual(id,'rIdPic1','not the id the builder used, which every insert starts again from');
  assert.deepEqual(a.mediaOf(id),{mime:'image/png',base64:PNG});
  assert.deepEqual(a.exportState().media,{[id]:{mime:'image/png',base64:PNG}});
});

test('two blocks that each had "their first picture" do not share one, and a picture of another kind is its own kind',async()=>{
  const a=demo();
  await a.insert(picture(pic(PNG,'image/png','First')),{where:'start',spare:false});
  await a.insert(picture(pic(JPG,'image/jpeg','Second')),{where:'start',spare:false});
  const xml=a.exportState().paragraphs.map(p=>p.xml).join('');
  const ids=rids(xml);
  assert.equal(new Set(ids).size,2);
  const media=a.exportState().media;
  assert.deepEqual(Object.values(media).map(m=>m.mime).sort(),['image/jpeg','image/png']);
  assert.equal(media[ids.find(id=>media[id].mime==='image/png')].base64,PNG);
  assert.equal(media[ids.find(id=>media[id].mime==='image/jpeg')].base64,JPG);
});

test('a picture whose paragraph is deleted is left out of what is saved, and comes back with Undo',async()=>{
  const a=demo();
  await a.insert(picture(),{where:'start',spare:false});
  assert.equal(Object.keys(a.exportState().media).length,1);
  await a.deleteRange({start:0,end:0,guard:''});
  assert.deepEqual(a.exportState().media,{},'nothing names it any more, so it is not saved');
  a.undo();
  assert.equal(Object.keys(a.exportState().media).length,1,'Undo brought the paragraph back, and its picture');
});

test('only a PNG, JPEG or GIF made of base64 is a picture; anything else is refused and nothing is inserted',async()=>{
  const a=demo();
  const mk=p=>ctx=>{const x=para('OSGCueBody',pictureRun(pic(),ctx));ctx.pictures[0]={...ctx.pictures[0],...p};return x;};
  for(const bad of [{mime:'image/svg+xml'},{mime:'text/html'},{base64:'not base64!'},{base64:'AAA'},{base64:''},{rid:'rIdIcon_x'},{rid:'x"y'}]){
    await assert.rejects(a.insert(mk(bad),{where:'start',spare:false}),/not a PNG, JPG or GIF/,JSON.stringify(bad));
  }
  assert.deepEqual(a.exportState().paragraphs,[],'nothing was inserted');
  assert.deepEqual(a.exportState().media,{});
});

test('a project with pictures reads back with them: the document gives them out and takes them back',async()=>{
  const a=demo();
  await a.insert(picture(),{where:'start',spare:false});
  const text=serializeProject(a.exportState());
  assert.match(text,/"media"/);
  const b=demo();
  b.importState(parseProject(text).guide);
  const [id]=rids(b.exportState().paragraphs[0].xml);
  assert.deepEqual(b.mediaOf(id),{mime:'image/png',base64:PNG});
  assert.equal(b.mediaOf('rIdPicNope'),null);
});

test('the package a Word document is made from holds the real picture, with the id its paragraph names, and the real icons when they can be loaded',async()=>{
  const a=demo({icons:async names=>Object.fromEntries(names.map(n=>[n,'QUJDRA==']))});
  await a.insert(ctx=>cueBlock({cue:'ask',text:'What?',picture:pic(JPG,'image/jpeg')},ctx),{where:'start',spare:false});
  const pkg=await a.sourcePackage();
  const body=a.exportState().paragraphs.map(p=>p.xml).join('');
  const picId=rids(body).find(id=>id.startsWith('rIdPic'));
  assert.ok(picId);
  assert.match(pkg,new RegExp(`<Relationship Id="${picId}" Type="[^"]*/image" Target="media/image\\d+\\.jpg"/>`),'the picture has a relationship');
  assert.match(pkg,new RegExp(`pkg:name="/word/media/image\\d+\\.jpg" pkg:contentType="image/jpeg"[^>]*><pkg:binaryData>${JPG.replace(/[+/]/g,'\\$&')}</pkg:binaryData>`),'and is its own part, as it was given');
  assert.match(pkg,/<pkg:binaryData>QUJDRA==<\/pkg:binaryData>/,'the cue\'s icon is the loaded picture');
  const without=demo();
  await without.insert(ctx=>cueBlock({cue:'ask',text:'What?'},ctx),{where:'start',spare:false});
  assert.match(await without.sourcePackage(),/<pkg:binaryData>iVBORw0KGgo/,'with no way to load icons a one-pixel stand-in is used');
  const failing=demo({icons:async()=>{throw new Error('offline');}});
  await failing.insert(ctx=>cueBlock({cue:'ask',text:'What?'},ctx),{where:'start',spare:false});
  assert.match(await failing.sourcePackage(),/<pkg:binaryData>iVBORw0KGgo/,'and so is it when they cannot be loaded');
});

test('the pictures travel with the block that names them across the message link, and are saved by the page that holds the document',async()=>{
  const saved=[];
  const guide=createProjectAdapter(demo(),{read:async()=>null,write:async t=>{saved.push(t);},delayMs:0,soon:fn=>fn(),cancel:()=>{}});
  const ends={frame:[],host:[]};
  const send=(to,m)=>{const copy=structuredClone(m);queueMicrotask(()=>ends[to].forEach(fn=>fn(copy)));};
  serveAdapter(guide,{post:m=>send('frame',m),listen:fn=>{ends.host.push(fn);}});
  const remote=createRemoteAdapter({post:m=>send('host',m),listen:fn=>{ends.frame.push(fn);}});
  await remote.ready;
  await remote.insert(picture(),{where:'start',spare:false});
  const [id]=rids(guide.exportState().paragraphs[0].xml);
  assert.deepEqual(guide.mediaOf(id),{mime:'image/png',base64:PNG},'the picture arrived with the text');
  await guide.flush();
  assert.match(saved.at(-1),/"media"/);
  await assert.rejects(remote.insert(ctx=>{const x=para('OSGCueBody',pictureRun(pic(),ctx));ctx.pictures[0].mime='image/svg+xml';return x;},{where:'start',spare:false}),/not a PNG, JPG or GIF/);
  assert.equal(Object.keys(guide.exportState().media).length,1,'a picture of a kind that is not allowed was not kept');
});
