import test from 'node:test';
import assert from 'node:assert/strict';
import {ICON_FILES} from '../src/guide-cues.mjs';
import {makeContext,cueBlock,buildPackage} from '../src/guide-ooxml.mjs';
import {readBlock,isPictureParagraph,kindOf} from '../src/guide-model.mjs';
import {imageInfo,prepareImage,pictureSize,toBase64,PICTURE_WIDTHS,MAX_SIDE,KEEP_BYTES,MAX_INPUT_BYTES} from '../src/image-prepare.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {TINY,parse,partXml,paragraphsOf,packageToDocx} from './helpers/guide-package.mjs';
import {started,click,submit,guideRows,addModule,addLesson} from './helpers/guide-harness.mjs';

const iconData=Object.fromEntries(ICON_FILES.map(i=>[i,TINY]));
const PNG=Buffer.from(TINY,'base64'); // a real 1 by 1 PNG
// A JPEG header with a start-of-frame marker holding the size, enough for imageInfo (not a picture a viewer could draw).
const jpeg=(w,h)=>Buffer.from([0xff,0xd8,0xff,0xe0,0x00,0x04,0,0,0xff,0xc0,0x00,0x11,0x08,h>>8,h&255,w>>8,w&255,0x03,0x01,0x22,0x00,0x02,0x11,0x01,0x03,0x11,0x01,0xff,0xd9]);
const gif=(w,h)=>Buffer.from([0x47,0x49,0x46,0x38,0x39,0x61,w&255,w>>8,h&255,h>>8,0,0,0]);
const file=(bytes,name='a.png')=>({name,size:bytes.length,arrayBuffer:async()=>bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.length)});
const pic=(extra={})=>({base64:TINY,type:'png',mime:'image/png',size:{cx:914400,cy:457200},alt:'A flow chart of the steps',decorative:false,name:'flow.png',...extra});

test('imageInfo reads the kind and size of PNG, JPEG and GIF files, and refuses anything else',()=>{
  assert.deepEqual(imageInfo(PNG),{type:'png',mime:'image/png',width:1,height:1});
  assert.deepEqual(imageInfo(jpeg(640,480)),{type:'jpeg',mime:'image/jpeg',width:640,height:480});
  assert.deepEqual(imageInfo(gif(300,200)),{type:'gif',mime:'image/gif',width:300,height:200});
  for(const bad of [Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'),Buffer.from('hello'),Buffer.alloc(0),Buffer.from([0x89,0x50,0x4e,0x47])])
    assert.throws(()=>imageInfo(bad),/not a picture the guide can use\. Choose a PNG, JPG or GIF picture\./);
});

test('a picture\'s size follows the width chosen, keeps its shape and is never taller than eight inches',()=>{
  assert.deepEqual(pictureSize({width:1600,height:900},'large'),{cx:4389120,cy:2468880});
  assert.equal(pictureSize({width:1600,height:900},'small').cx,Math.round(2.2*914400));
  assert.equal(pictureSize({width:100,height:100},'nonsense').cx,Math.round(PICTURE_WIDTHS[0].inches*914400),'an unknown width falls back to the first');
  const tall=pictureSize({width:100,height:5000},'large');
  assert.equal(tall.cy,8*914400);
  assert.ok(Math.abs(tall.cx/tall.cy-100/5000)<0.001);
});

test('a small picture is used as it is; a large one is shrunk where a canvas exists; a huge one is refused where none does',async()=>{
  const small=await prepareImage(file(PNG,'dot.png'));
  assert.equal(small.shrunk,false);assert.equal(small.name,'dot.png');assert.equal(small.bytes.length,PNG.length);
  // A GIF is never re-drawn (it may move).
  assert.equal((await prepareImage(file(gif(3000,3000),'big.gif'))).shrunk,false);
  // Large by size: no canvas, over 5 MB, refused; under 5 MB used as it is.
  const wide=Buffer.concat([jpeg(4000,3000),Buffer.alloc(KEEP_BYTES+10)]);
  assert.equal((await prepareImage(file(wide,'wide.jpg'),{env:{}})).shrunk,false);
  await assert.rejects(prepareImage(file(Buffer.concat([jpeg(4000,3000),Buffer.alloc(6*1024*1024)]),'huge.jpg'),{env:{}}),/too large to use as it is/);
  // With a canvas: it is drawn smaller, as the same kind of picture.
  const calls=[];
  const env={
    Blob:class{constructor(parts,opts){this.parts=parts;this.type=opts.type;}},
    createImageBitmap:async()=>({width:4000,height:3000,close(){calls.push('close');}}),
    document:{createElement:tag=>{calls.push(tag);return {getContext:()=>({drawImage:(b,x,y,w,h)=>calls.push(`draw ${w}x${h}`)}),toBlob(cb,type,q){calls.push(`toBlob ${type} ${q}`);cb({arrayBuffer:async()=>new Uint8Array(1000).buffer});}};}}
  };
  const shrunk=await prepareImage(file(wide,'wide.jpg'),{env});
  assert.equal(shrunk.shrunk,true);assert.equal(shrunk.width,MAX_SIDE);assert.equal(shrunk.height,1200);assert.equal(shrunk.type,'jpeg');assert.equal(shrunk.bytes.length,1000);
  assert.deepEqual(calls,['canvas','draw 1600x1200','close','toBlob image/jpeg 0.85']);
  await assert.rejects(prepareImage(file(Buffer.alloc(8),'x.png')),/not a picture/);
  await assert.rejects(prepareImage({name:'big.png',size:MAX_INPUT_BYTES+1,arrayBuffer:async()=>new ArrayBuffer(0)}),/Choose one under 12 MB/);
  await assert.rejects(prepareImage(null),/Choose a picture file first/);
});

test('a block with a picture: the cue line, then a paragraph of body style holding the picture, then the rest',()=>{
  const ctx=makeContext(),body=cueBlock({cue:'picture',text:'The conflict cycle',notes:'Point to each arrow.',minutes:3,picture:pic()},ctx);
  const pkg=buildPackage(body,ctx,iconData),paras=paragraphsOf(pkg);
  assert.deepEqual(paras.map(p=>p.style),['OSG Cue','OSG Cue Body','OSG Facilitator Notes']);
  assert.match(paras[0].text,/Picture\s+The conflict cycle \(3 min\)/);
  assert.equal(isPictureParagraph(paras[1]),true);
  assert.equal(isPictureParagraph(paras[0]),false);
  const doc=partXml(pkg,'/word/document.xml');
  assert.match(doc,/<wp:docPr id="\d+" name="Picture \d+" descr="A flow chart of the steps">/);
  assert.match(doc,/<wp:extent cx="914400" cy="457200"\/>/);
  assert.match(doc,/<a:blip r:embed="rIdPic1"\/>/);
  assert.match(pkg,/<Relationship Id="rIdPic1" Type="[^"]*\/image" Target="media\/pic1\.png"\/>/);
  assert.match(pkg,/<pkg:part pkg:name="\/word\/media\/pic1\.png" pkg:contentType="image\/png" pkg:compression="store"><pkg:binaryData>/);
  parse(pkg);
});

test('every kind of picture gets its own part, name and content type, and two pictures do not share one',()=>{
  const ctx=makeContext(),body=cueBlock({cue:'picture',text:'One',picture:pic({type:'jpeg',mime:'image/jpeg'})},ctx)+cueBlock({cue:'media',text:'Two',picture:pic({type:'gif',mime:'image/gif',alt:'A moving arrow'})},ctx);
  const pkg=buildPackage(body,ctx,iconData);
  assert.match(pkg,/Target="media\/pic1\.jpeg"/);assert.match(pkg,/Target="media\/pic2\.gif"/);
  assert.match(pkg,/pkg:name="\/word\/media\/pic1\.jpeg" pkg:contentType="image\/jpeg"/);
  assert.match(pkg,/pkg:name="\/word\/media\/pic2\.gif" pkg:contentType="image\/gif"/);
  const ids=[...partXml(pkg,'/word/document.xml').matchAll(/<wp:docPr id="(\d+)"/g)].map(m=>m[1]);
  assert.equal(new Set(ids).size,ids.length,'picture ids are unique');
});

test('a picture needs alt text or a decorative mark, and a Picture block needs a picture',()=>{
  assert.throws(()=>cueBlock({cue:'picture',text:'x',picture:pic({alt:'  '})},makeContext()),/Describe the picture in a few words/);
  const ctx=makeContext(),body=cueBlock({cue:'picture',text:'x',picture:pic({alt:'',decorative:true})},ctx);
  assert.match(body,/descr=""[^>]*><a:extLst>.*adec:decorative.*val="1"/);
  assert.throws(()=>cueBlock({cue:'picture',text:'x'},makeContext()),/Choose a picture file for a Picture block/);
  assert.doesNotThrow(()=>cueBlock({cue:'say',text:'x'},makeContext()),'other cues need no picture');
});

test('the structure check is satisfied by a picture with alt text and fails one without',async()=>{
  const ok=makeContext(),pkg=buildPackage(cueBlock({cue:'picture',text:'Cycle',picture:pic()},ok),ok,iconData);
  const good=await checkDocx(await packageToDocx(pkg,'A guide'));
  assert.ok(good.findings.some(f=>f.rule==='IMAGE-ALT'&&f.level==='pass'));
  assert.equal(good.findings.filter(f=>f.rule==='IMAGE-ALT'&&f.level!=='pass').length,0);
  // Hand-removing the alt text in Word is what the check is for: the same package with the text taken out fails.
  const bad=await checkDocx(await packageToDocx(pkg.replace('descr="A flow chart of the steps"','descr=""'),'A guide'));
  assert.ok(bad.findings.some(f=>f.rule==='IMAGE-ALT'&&f.level==='fail'));
});

test('readBlock sees a picture, leaves it out of the block\'s lines, and knows the block has one',()=>{
  const ctx=makeContext(),pkg=buildPackage(cueBlock({cue:'picture',text:'The cycle\nSecond line',notes:'Notes here',picture:pic()},ctx),ctx,iconData),paras=paragraphsOf(pkg);
  const b=readBlock(paras,0);
  assert.equal(b.hasPicture,true);
  assert.equal(b.text,'The cycle\nSecond line');
  assert.equal(b.notes,'Notes here');
  const plain=paragraphsOf(buildPackage(cueBlock({cue:'say',text:'Hello'},makeContext()),Object.assign(makeContext(),{used:new Set(['lecture'])}),iconData));
  assert.equal(readBlock(plain,0).hasPicture,false);
  assert.equal(kindOf(paras[1]),'body');
});

// The file window's work: the picture comes back read (and shrunk if large); the pane then asks for a description.
async function chooseAndDescribe(a,bytes,name,{alt='A flow chart',decorative=false,width='large'}={}){
  const prepared=await prepareImage(file(bytes,name));
  a.adapter.setPicker(async opts=>{assert.equal(opts.kind,'picture');return {files:[prepared]};});
  await click(a,'guide-pic-choose');
  a.type('guide-pic-alt',alt);a.n('guide-pic-decorative').checked=decorative;a.type('guide-pic-width',width);
}

test('the pane inserts a Picture block from a chosen file and clears the picture fields afterwards',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.pick('picture');
  assert.equal(a.n('guide-pic-box').open,true,'the picture box is open for a Picture block');
  assert.match(a.n('guide-pic-summary').textContent,/needed for this block/);
  await chooseAndDescribe(a,PNG,'cycle.png',{alt:'The conflict cycle',width:'medium'});
  assert.match(a.n('guide-pic-chosen').textContent,/cycle\.png: 1 by 1 pixels/);
  a.type('guide-block-text','Figure 1: The conflict cycle');a.type('guide-block-notes','');a.type('guide-block-minutes','');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  const rows=guideRows(a),i=rows.findIndex(r=>/Picture\s+Figure 1/.test(r));
  assert.ok(i>=0,rows.join(' / '));
  assert.equal(rows[i+1],'OSG Cue Body: ','the picture is its own paragraph of body style');
  assert.equal(a.n('guide-pic-alt').value,'');assert.equal(a.n('guide-pic-chosen').textContent,'');
  assert.match(a.n('guide-status').textContent,/Picture block added/);
});

test('the pane asks for a file, then a description, before it inserts a picture',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.pick('picture');a.type('guide-block-text','Figure');
  await submit(a,'guide-block-form');
  assert.match(a.err(),/Choose a picture file first\. Choose a PNG, JPG or GIF picture\./);
  await chooseAndDescribe(a,PNG,'x.png',{alt:''});
  await submit(a,'guide-block-form');
  assert.match(a.err(),/Describe the picture in a few words/);
  a.n('guide-pic-decorative').checked=true;
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'','decoration needs no description');
  // A file that is not a picture says so and leaves the fields as they were.
  const b=await started();
  b.pick('picture');
  b.adapter.setPicker(async()=>prepareImage(file(Buffer.from('hello'),'note.txt')));
  await click(b,'guide-pic-choose');
  assert.match(b.err(),/not a picture the guide can use/);
  // Closing the window without choosing is not a problem.
  b.adapter.setPicker(async()=>({cancelled:true}));
  await click(b,'guide-pic-choose');
  assert.equal(b.err(),'');assert.match(b.n('guide-status').textContent,/No picture was chosen/);
});

test('a picture can ride on any cue, with or without words, and its block cannot be edited in the pane',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.pick('slide');
  assert.match(a.n('guide-pic-summary').textContent,/optional/);
  await chooseAndDescribe(a,jpeg(800,600),'slide3.jpg',{alt:'Slide 3: the three questions'});
  a.type('guide-block-text','');a.type('guide-block-notes','');a.type('guide-block-minutes','');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'','a block with a picture needs no words');
  const rows=guideRows(a);
  assert.ok(rows.some(r=>/^OSG Cue: .*Slide/.test(r)));
  // Put the cursor in that block and try to edit it.
  const at=rows.findIndex(r=>/^OSG Cue: .*Slide/.test(r));
  a.adapter.place(a.rows().length-rows.length+at);await a.p.quietRefresh();
  await click(a,'guide-edit-selected');
  assert.match(a.err(),/This block has a picture/);
  assert.equal(a.p.state().editing,null);
});

test('choosing a file may take longer than the pane\'s usual wait, because the author is the slow one',async()=>{
  const a=await started('T',{timeoutMs:60});
  const prepared=await prepareImage(file(PNG,'slow.png'));
  a.adapter.setPicker(()=>new Promise(resolve=>setTimeout(()=>resolve({files:[prepared]}),250)));
  await click(a,'guide-pic-choose');
  assert.equal(a.err(),'','no "did not answer in time" while the file window is open');
  assert.match(a.n('guide-pic-chosen').textContent,/slow\.png: 1 by 1 pixels/);
  // An ordinary action that takes as long is still cut off.
  const real=a.adapter.paragraphs;a.adapter.paragraphs=()=>new Promise(resolve=>setTimeout(()=>resolve(real.call(a.adapter)),250));
  await click(a,'guide-refresh');
  assert.match(a.err(),/Word did not answer in time/);
});
