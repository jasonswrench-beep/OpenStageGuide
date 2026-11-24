import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {pptxFixture} from './helpers/pptx-fixture.mjs';
import {readPptx} from '../src/pptx.mjs';
import {readChosen} from '../src/file-pick.mjs';
import {prepareImage} from '../src/image-prepare.mjs';
import {numberInName,matchPictures,slideAlt,slideBlock,MAX_SLIDES_AT_ONCE} from '../src/guide-slides.mjs';
import {TINY} from './helpers/guide-package.mjs';
import {started,addModule,addLesson,addBlock,guideRows,click} from './helpers/guide-harness.mjs';
globalThis.DOMParser=DOMParser;

const PNG=Buffer.from(TINY,'base64');
const file=(bytes,name)=>({name,size:bytes.length,arrayBuffer:async()=>bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.length)});
const pics=names=>names.map(n=>file(PNG,n));
const slide=(order,extra={})=>({order,title:`Title ${order}`,text:`Title ${order}\nBody ${order}`,notes:'',clicks:0,hidden:false,...extra});
// Two animation steps on a click and a third effect that goes with the first: two clicks, not three.
const TIMING=`<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3"><p:childTnLst><p:par><p:cTn id="5" nodeType="clickEffect"/><p:cTn id="6" nodeType="withEffect"/></p:par></p:childTnLst></p:cTn></p:par><p:par><p:cTn id="7"><p:childTnLst><p:par><p:cTn id="8" nodeType="clickEffect"/></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
const deckBytes=change=>pptxFixture(change);
const withClicks=async z=>{const path='ppt/slides/slide1.xml',x=await z.file(path).async('string');z.file(path,x.replace('</p:sld>',`${TIMING}</p:sld>`));};
const hideSecond=async z=>{const path='ppt/slides/slide2.xml',x=await z.file(path).async('string');z.file(path,x.replace('<p:sld ','<p:sld show="0" '));};
const asFile=(bytes,name='deck.pptx')=>({name,size:bytes.length,arrayBuffer:async()=>bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.length)});

test('the reader counts a slide\'s clicks (effects that go with another count once) and knows a hidden slide',async()=>{
  const bytes=await deckBytes(async z=>{await withClicks(z);await hideSecond(z);});
  const source=await readPptx(bytes,'d.pptx',JSZip,{details:true});
  // The fixture lists slide2.xml first, then slide1.xml.
  assert.deepEqual(source.slides.map(s=>[s.order,s.title,s.clicks,s.hidden]),[[1,'Title 2',0,true],[2,'Title 1',2,false]]);
  // Without the option the older shape is untouched, which the preview and review code compare as it stands.
  assert.deepEqual(Object.keys((await readPptx(bytes,'d.pptx',JSZip)).slides[0]),['id','sourceId','sourcePart','order','title','text','notes','warnings']);
});

test('a number in a file name is the last run of digits before the extension',()=>{
  assert.equal(numberInName('Slide12.PNG'),12);
  assert.equal(numberInName('deck 3 final 07.jpg'),7);
  assert.equal(numberInName('intro.png'),null);
  assert.equal(numberInName(''),null);
});

test('pictures are matched to slides by the number in their names, in any order',()=>{
  const slides=[slide(1),slide(2),slide(3)];
  const r=matchPictures(slides,pics(['Slide3.PNG','Slide1.PNG','Slide2.PNG']));
  assert.deepEqual([...r.byOrder].map(([o,f])=>[o,f.name]),[[3,'Slide3.PNG'],[1,'Slide1.PNG'],[2,'Slide2.PNG']]);
  assert.deepEqual(r.problems,[]);
  // Fewer pictures than slides is fine: the rest simply have none.
  assert.deepEqual([...matchPictures(slides,pics(['Slide2.PNG'])).byOrder.keys()],[2]);
  // A picture for a slide that is not there is named.
  const strays=matchPictures(slides,pics(['Slide1.PNG','Slide9.PNG']));
  assert.deepEqual([...strays.byOrder.keys()],[1]);
  assert.match(strays.problems[0],/not for any slide in the presentation \(it has 3\): Slide9\.PNG/);
});

test('names without numbers are taken in name order only when there is one for each slide',()=>{
  const slides=[slide(1),slide(2)];
  const same=matchPictures(slides,pics(['b.png','a.png']));
  assert.deepEqual([...same.byOrder].map(([o,f])=>[o,f.name]),[[1,'a.png'],[2,'b.png']]);
  const fewer=matchPictures(slides,pics(['only.png']));
  assert.equal(fewer.byOrder.size,0);
  assert.match(fewer.problems[0],/have no slide numbers and you chose 1 for a presentation of 2 slides/);
  // Two files with the same number cannot be told apart either.
  assert.match(matchPictures(slides,pics(['a1.png','b1.png','c.png'])).problems[0],/cannot be matched/);
  assert.deepEqual(matchPictures(slides,[]).problems,[]);
});

test('a slide\'s block says its number, title and clicks, carries its notes, and describes its picture',()=>{
  const s=slide(4,{title:'The three questions',text:'The three questions\nWho? What? Why?',notes:' Pause here. \n\nAsk for a volunteer.',clicks:3});
  assert.equal(slideAlt(s),'Slide 4: The three questions. On the slide: Who? What? Why?');
  const b=slideBlock(s,{picture:{base64:TINY,type:'png',mime:'image/png',size:{cx:1,cy:1},name:'Slide4.PNG'}});
  assert.equal(b.cue,'slide');assert.equal(b.text,'Slide 4: The three questions\nAnimation: 3 clicks.');
  assert.equal(b.notes,'Pause here.\nAsk for a volunteer.');
  assert.equal(b.audience,'facilitator');assert.equal(b.picture.alt,slideAlt(s));assert.equal(b.picture.decorative,false);
  assert.equal(slideBlock(s,{share:true}).audience,'everyone');
  assert.equal(slideBlock(slide(5,{title:'',text:'',clicks:1})).text,'Slide 5\nAnimation: 1 click.');
  assert.equal(slideBlock(slide(5)).picture,undefined,'no picture, none made up');
  assert.ok(slideAlt(slide(1,{text:'x '.repeat(400)})).length<=250);
  assert.equal(slideAlt(slide(2,{title:'',text:''})),'Slide 2.');
});

// The file window's work, done for real (reading the presentation, reading the pictures), then handed to the pane as the window would.
const readDeck=async f=>readPptx(await f.arrayBuffer(),f.name,JSZip,{details:true});
async function loadDeck(a,bytes,name,pictureNames=[]){
  const chosen=await readChosen({deckFile:asFile(bytes,name),pictureFiles:pics(pictureNames),prepare:prepareImage,readDeck});
  a.adapter.setPicker(async opts=>{assert.equal(opts.kind,'slides');return chosen;});
  await click(a,'guide-slides-choose');
}
const slideBox=(a,n)=>a.n('guide-slides-list').children[n-1].children[0];

test('the pane reads a presentation, lists its slides (hidden ones not chosen) and inserts the chosen ones with their pictures and notes',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await loadDeck(a,await deckBytes(async z=>{await withClicks(z);await hideSecond(z);}),'My deck.pptx',['Slide1.PNG','Slide2.PNG']);
  assert.equal(a.err(),'');
  assert.equal(a.n('guide-slides-fieldset').hidden,false);
  assert.deepEqual(a.n('guide-slides-list').children.map(l=>l.textContent.trim()),['Slide 1: Title 2 (hidden)','Slide 2: Title 1, 2 clicks']);
  assert.equal(slideBox(a,1).checked,false,'a hidden slide is not chosen');assert.equal(slideBox(a,2).checked,true);
  assert.match(a.n('guide-slides-summary').textContent,/My deck\.pptx: 2 slides, 1 chosen\./);
  assert.match(a.n('guide-slides-pics-note').textContent,/2 pictures chosen; every chosen slide has one/);
  slideBox(a,1).checked=true;slideBox(a,1).onchange();
  assert.match(a.n('guide-slides-pics-note').textContent,/every chosen slide has one/);
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  const rows=guideRows(a);
  const i=rows.findIndex(r=>/^OSG Cue Facilitator: .*Slide\s+Slide 1: Title 2/.test(r));
  assert.ok(i>=0,rows.join(' / '));
  assert.equal(rows[i+1],'OSG Cue Body: ','slide 1\'s picture follows its cue line');
  const j=rows.findIndex(r=>/Slide 2: Title 1/.test(r));
  assert.ok(j>i);
  assert.equal(rows[j+1],'OSG Cue Body: ','the picture comes straight after the cue line');
  assert.match(rows[j+2],/^OSG Cue Body: Animation: 2 clicks\.$/);
  // the slide's own words come under the picture, shaded like the speaker notes that follow them
  assert.match(rows[j+3],/^OSG Facilitator Notes: Slide text: Text & task 1$/);
  assert.match(rows[j+4],/^OSG Facilitator Notes: Facilitator notes: PRIVATE SPEAKER NOTE$/);
  assert.match(rows[i+2],/^OSG Facilitator Notes: Slide text: Text & task 2$/);
  assert.match(a.n('guide-status').textContent,/2 slides inserted\./);
  // Both are now un-chosen, so inserting again says so.
  await click(a,'guide-slides-go');
  assert.match(a.err(),/Choose at least one slide to insert/);
});

test('slides with no matching picture come in without one and say so; the workbook box makes them for everyone',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await loadDeck(a,await deckBytes(),'deck.pptx',['Slide2.PNG']);
  slideBox(a,1).checked=true;slideBox(a,1).onchange();
  assert.match(a.n('guide-slides-pics-note').textContent,/1 picture chosen; 1 chosen slide will have no picture/);
  a.n('guide-slides-share').checked=true;
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  const rows=guideRows(a);
  assert.ok(rows.some(r=>/^OSG Cue: .*Slide 1: Title 2/.test(r)),'shown to everyone uses the plain cue style');
  assert.match(a.n('guide-status').textContent,/2 slides inserted, 1 of them without a picture/);
});

test('a presentation chosen without any pictures says plainly that the slides will go in as text, and so does the result',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await loadDeck(a,await deckBytes(),'deck.pptx');
  assert.match(a.n('guide-slides-pics-note').textContent,/No pictures of the slides were chosen, so each slide goes in as text.*give a PDF of it/);
  slideBox(a,1).checked=true;slideBox(a,1).onchange();
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  assert.match(a.n('guide-status').textContent,/2 slides inserted, with no picture \(no pictures of the slides were chosen; they went in as text\)\./);
  // Before a presentation is chosen there is nothing to say.
  const b=await started();
  assert.equal(b.n('guide-slides-pics-note').textContent,'');
});

test('a slide whose deck gave no title is named by its number alone, and the list says the same',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  const bytes=await deckBytes(async z=>{
    const strip=async p=>{const x=await z.file(p).async('string');z.file(p,x.replace(/<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr>.*?<\/p:sp>/s,''));};
    await strip('ppt/slides/slide1.xml');
  });
  await loadDeck(a,bytes,'deck.pptx');
  const labels=a.n('guide-slides-list').children.map(l=>l.textContent.trim());
  assert.ok(labels.every(l=>!/no title/i.test(l)),labels.join(' / '));
  slideBox(a,1).checked=true;slideBox(a,2).checked=true;slideBox(a,2).onchange();
  await click(a,'guide-slides-go');
  assert.ok(guideRows(a).every(r=>!/no title/i.test(r)));
});

test('the pane refuses pictures it cannot match, a non-presentation, and too many slides at once',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await click(a,'guide-slides-go');
  assert.match(a.err(),/Choose a presentation first/);
  a.adapter.setPicker(async()=>{throw new Error('Presentation preview: choose a standard .pptx presentation.');});
  await click(a,'guide-slides-choose');
  assert.match(a.err(),/Presentation preview/);
  a.adapter.setPicker(async()=>({cancelled:true}));
  await click(a,'guide-slides-choose');
  assert.match(a.n('guide-status').textContent,/No presentation was chosen/);
  await loadDeck(a,await deckBytes(),'deck.pptx',['Slide7.PNG']);
  await click(a,'guide-slides-go');
  assert.match(a.err(),/not for any slide in the presentation \(it has 2\): Slide7\.PNG/);
  assert.ok(MAX_SLIDES_AT_ONCE>=20);
});

test('a deck with a chart and its embedded workbook is read for its slides, while macros are still refused',async()=>{
  const withChart=async z=>{
    z.file('ppt/charts/chart1.xml','<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
    z.file('ppt/charts/_rels/chart1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Sheet1.xlsx"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/></Relationships>');
    z.file('ppt/embeddings/Microsoft_Excel_Sheet1.xlsx',Buffer.from('not really a workbook'));
    z.file('ppt/embeddings/oleObject1.bin',Buffer.from('not really an object'));
  };
  const bytes=await deckBytes(withChart);
  // The guide's reader (details on) reads the slides and never opens the embedded parts.
  const source=await readPptx(bytes,'chart.pptx',JSZip,{details:true});
  assert.deepEqual(source.slides.map(s=>s.title),['Title 2','Title 1']);
  // The older preview keeps refusing a deck like that, as it always did.
  await assert.rejects(readPptx(bytes,'chart.pptx',JSZip),/embedded objects are unsupported|embedded-object relationships/);
  // Macros are refused either way.
  const macro=await deckBytes(async z=>{z.file('ppt/vbaProject.bin',Buffer.from('macro'));});
  await assert.rejects(readPptx(macro,'macro.pptx',JSZip,{details:true}),/macros, active content/);
  await assert.rejects(readPptx(macro,'macro.pptx',JSZip),/macros, active content/);
  const activeX=await deckBytes(async z=>{
    z.file('ppt/slides/_rels/slide2.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControl" Target="../activeX/a.xml"/></Relationships>');
  });
  await assert.rejects(readPptx(activeX,'ax.pptx',JSZip,{details:true}),/active or embedded-object relationships/);
});

test('the slide chooser reads a big presentation (the pictures and video are never read); the older preview still stops at 16 MB',async()=>{
  // A real deck of the author's was refused as "too big" at 16 MB. The chooser needs only the slides and notes.
  const stored=await deckBytes(async z=>{z.file('ppt/media/video1.bin',Buffer.alloc(17*1024*1024,7),{compression:'STORE'});});
  assert.ok(stored.length>16*1024*1024);
  const source=await readPptx(stored,'big.pptx',JSZip,{details:true});
  assert.deepEqual(source.slides.map(s=>s.title),['Title 2','Title 1']);
  await assert.rejects(readPptx(stored,'big.pptx',JSZip),/up to 16 MB/);
  // Small on disk but a lot of media once opened: more than the 64 MB the older reader allows in all, and none of it is read.
  const squeezed=await deckBytes(async z=>{for(const n of [1,2,3])z.file(`ppt/media/image${n}.bin`,Buffer.alloc(25*1024*1024),{compression:'DEFLATE'});});
  assert.ok(squeezed.length<16*1024*1024);
  assert.equal((await readPptx(squeezed,'squeezed.pptx',JSZip,{details:true})).slides.length,2);
  await assert.rejects(readPptx(squeezed,'squeezed.pptx',JSZip),/oversized ZIP parts/);
  // A chart's data can be bigger than a slide's XML may be, and is not read either.
  const chart=await deckBytes(async z=>{z.file('ppt/charts/chart1.xml',`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">${' '.repeat(3*1024*1024)}</c:chartSpace>`);});
  assert.equal((await readPptx(chart,'chart.pptx',JSZip,{details:true})).slides.length,2);
  await assert.rejects(readPptx(chart,'chart.pptx',JSZip),/XML part exceeds 2 MB/);
  // A slide's own XML is still held to its limit.
  const fat=await deckBytes(async z=>{const x=await z.file('ppt/slides/slide1.xml').async('string');z.file('ppt/slides/slide1.xml',x.replace('</p:spTree>',`<!-- ${' '.repeat(3*1024*1024)} --></p:spTree>`));});
  await assert.rejects(readPptx(fat,'fat.pptx',JSZip,{details:true}),/XML part exceeds 2 MB/);
  // Past the chooser's own limit the message says what to do, in megabytes.
  const huge=new Uint8Array(251*1024*1024);
  await assert.rejects(readPptx(huge,'huge.pptx',JSZip,{details:true}),/251 MB and the slide chooser reads up to 250 MB.*compress the media/);
});

test('a deck with no title placeholders is titled by the largest, highest line of text, and the reader says it guessed',async()=>{
  const shape=(text,{size,y,id,ph=''}={})=>`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${y===undefined?'':`<a:xfrm><a:off x="0" y="${y}"/><a:ext cx="100" cy="100"/></a:xfrm>`}</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr${size?` sz="${size}"`:''}/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const slideXml=body=>`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`;
  const bytes=await deckBytes(async z=>{
    // slide1.xml (shown second): a page number, a long paragraph, a footer, a small heading lower down and the large title at the top
    z.file('ppt/slides/slide1.xml',slideXml(
      shape('7',{size:6000,y:0,id:2})+
      shape('This long paragraph is body text and goes on and on, so it is never taken for a title because it is far longer than a heading would be, however large it is set.',{size:5000,y:10,id:3})+
      shape('Company footer',{size:4000,y:9000000,id:4,ph:'<p:ph type="ftr"/>'})+
      shape('A small heading',{size:2000,y:500,id:5})+
      shape('Why conflict feels personal',{size:4400,y:100,id:6})+
      shape('Second line of body',{size:1800,y:900,id:7})));
    // slide2.xml (shown first): equal sizes, so the higher one wins; sizes missing, so document order decides
    z.file('ppt/slides/slide2.xml',slideXml(shape('Lower line',{size:2400,y:800,id:2})+shape('Higher line',{size:2400,y:100,id:3})));
  });
  const [first,second]=(await readPptx(bytes,'d.pptx',JSZip,{details:true})).slides;
  assert.deepEqual([first.title,first.titleGuessed],['Higher line',true]);
  assert.deepEqual([second.title,second.titleGuessed],['Why conflict feels personal',true]);
  // Without sizes or places the first line in the file is used; text that is only a number gives no title at all.
  const plain=await deckBytes(async z=>{
    z.file('ppt/slides/slide1.xml',slideXml(shape('First thing',{id:2})+shape('Second thing',{id:3})));
    z.file('ppt/slides/slide2.xml',slideXml(shape('12',{id:2})+shape('03 / 2024',{id:3})));
  });
  const [a,b]=(await readPptx(plain,'p.pptx',JSZip,{details:true})).slides;
  assert.deepEqual([a.title,a.titleGuessed],['',false]);
  assert.deepEqual([b.title,b.titleGuessed],['First thing',true]);
  // A real title placeholder always wins, and is not marked as a guess; the older reader never guesses.
  const real=await deckBytes();
  const [x]=(await readPptx(real,'r.pptx',JSZip,{details:true})).slides;
  assert.deepEqual([x.title,x.titleGuessed],['Title 2',false]);
  assert.equal((await readPptx(plain,'p.pptx',JSZip)).slides[1].title,'');
});

// ---- any number of slides: they go into Word in runs of up to 40 slides (and 20 MB of pictures), in order, from one press
const bigDeck=n=>({name:'big.pptx',slides:Array.from({length:n},(_,i)=>({order:i+1,sourceId:256+i,title:`Topic ${i+1}`,text:`Topic ${i+1}\nBody ${i+1}`,notes:'',clicks:0,hidden:false,titleGuessed:false}))});
const slideNumbers=a=>guideRows(a).map(r=>/Slide (\d+): Topic/.exec(r)).filter(Boolean).map(m=>Number(m[1]));
const countInserts=a=>{const real=a.adapter.insert.bind(a.adapter),log={calls:0,fail:null};a.adapter.insert=async(...args)=>{log.calls++;if(log.fail&&log.calls===log.fail)throw new Error('Word stopped the insert.');return real(...args);};return log;};

test('ninety-five slides go in with one press, in runs of forty, in order',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.adapter.setPicker(async()=>({deck:bigDeck(95),files:[]}));
  await click(a,'guide-slides-choose');
  assert.match(a.n('guide-slides-summary').textContent,/big\.pptx: 95 slides, 95 chosen/);
  const log=countInserts(a);
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  assert.equal(log.calls,3,'40 + 40 + 15');
  assert.deepEqual(slideNumbers(a),Array.from({length:95},(_,i)=>i+1),'every slide, once, in the presentation\'s order');
  assert.match(a.n('guide-status').textContent,/95 slides inserted, with no picture/);
  assert.ok(a.n('guide-slides-list').children.every(l=>!l.children[0].checked),'none is left ticked');
  // exactly forty is still one run
  const b=await started();
  await addModule(b,'Module one');await addLesson(b,'Lesson one');
  b.adapter.setPicker(async()=>({deck:bigDeck(40),files:[]}));
  await click(b,'guide-slides-choose');
  const logB=countInserts(b);
  await click(b,'guide-slides-go');
  assert.equal(b.err(),'');assert.equal(logB.calls,1);assert.equal(slideNumbers(b).length,40);
});

test('if Word stops partway, what went in stays, the rest stay ticked, and pressing again carries on',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.adapter.setPicker(async()=>({deck:bigDeck(95),files:[]}));
  await click(a,'guide-slides-choose');
  const log=countInserts(a);log.fail=2;
  await click(a,'guide-slides-go');
  assert.match(a.err(),/40 slides of 95 went in, then it stopped: Word stopped the insert\. The others are still ticked; choose Insert again to carry on\./);
  assert.deepEqual(slideNumbers(a),Array.from({length:40},(_,i)=>i+1));
  const boxes=a.n('guide-slides-list').children.map(l=>l.children[0]);
  assert.equal(boxes.slice(0,40).some(b=>b.checked),false);assert.equal(boxes.slice(40).every(b=>b.checked),true);
  log.fail=null;
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  assert.deepEqual(slideNumbers(a),Array.from({length:95},(_,i)=>i+1),'the rest follow the first forty, in order');
  assert.match(a.n('guide-status').textContent,/55 slides inserted/);
});

test('a failure on the very first run says only what went wrong, and nothing is unticked',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  a.adapter.setPicker(async()=>({deck:bigDeck(3),files:[]}));
  await click(a,'guide-slides-choose');
  const log=countInserts(a);log.fail=1;
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'Word stopped the insert.');
  assert.ok(a.n('guide-slides-list').children.every(l=>l.children[0].checked));
});

test('pictures add up run by run: runs end at 20 MB, and one picture bigger than that is refused plainly',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  const MB=1024*1024,pic=(n,mb)=>({name:`Slide${n}.png`,type:'png',mime:'image/png',width:4,height:3,shrunk:false,bytes:new Uint8Array(mb*MB)});
  a.adapter.setPicker(async()=>({deck:bigDeck(5),files:[1,2,3,4,5].map(n=>pic(n,9))}));
  await click(a,'guide-slides-choose');
  const log=countInserts(a);
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  assert.equal(log.calls,3,'9 + 9 MB, 9 + 9 MB, 9 MB: never more than 20 MB in one run');
  assert.deepEqual(slideNumbers(a),[1,2,3,4,5]);
  const b=await started();
  await addModule(b,'Module one');await addLesson(b,'Lesson one');
  b.adapter.setPicker(async()=>({deck:bigDeck(2),files:[pic(1,1),pic(2,21)]}));
  await click(b,'guide-slides-choose');
  const logB=countInserts(b);
  await click(b,'guide-slides-go');
  assert.match(b.err(),/The picture for slide 2 is larger than 20 MB\. Choose a smaller picture or a PDF\./);
  assert.equal(logB.calls,0,'nothing was inserted before the problem was found');
});

test('a slide shared with learners keeps its typed slide text out of the workbook, and the leak check still passes',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await loadDeck(a,await deckBytes(),'deck.pptx',['Slide1.PNG','Slide2.PNG']);
  slideBox(a,1).checked=true;slideBox(a,1).onchange();
  a.n('guide-slides-share').checked=true;
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  assert.ok(guideRows(a).some(r=>/Slide text: Text & task 1$/.test(r)),'the slide text is under the picture');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
  const xml=await (await JSZip.loadAsync(Buffer.from(opened.base64,'base64'))).file('word/document.xml').async('string');
  assert.ok(!xml.includes('Slide text'),'the label is not in the workbook');
  assert.ok(!xml.includes('PRIVATE SPEAKER NOTE'),'nor the speaker notes');
});

test('slides put in right after a module that has no lesson stay inside it, the next module goes after them, and the document ends in a plain paragraph',async()=>{
  // The author's recording of 2026-11-08: Add module, Import slides, then Add module again. The slides used to land after the first module's
  // Transition, belong to no module, and every new module landed above them; the last notes paragraph was the end of the document, so Enter
  // only made more notes lines.
  const a=await started();
  await addModule(a,'Module one');
  await loadDeck(a,await deckBytes(),'deck.pptx',['Slide1.PNG','Slide2.PNG']);
  slideBox(a,1).checked=true;slideBox(a,1).onchange();slideBox(a,2).checked=true;slideBox(a,2).onchange();
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  const snap=a.adapter.snapshot();
  assert.equal(snap.at(-1).style,'Normal');assert.equal(snap.at(-1).text.trim(),'','a plain empty paragraph ends the document, below the last notes');
  const rows=guideRows(a),firstSlide=rows.findIndex(r=>/Slide\s+Slide 1: /.test(r)),transition=rows.findIndex(r=>/Transition/.test(r));
  assert.ok(firstSlide>0&&transition>firstSlide,'the slides are inside module one, before its Transition');
  await addModule(a,'Module two');
  const after=guideRows(a),two=after.findIndex(r=>/^OSG Module: Module 2: Module two/.test(r));
  assert.ok(two>after.findIndex(r=>/Transition/.test(r)),'module two follows module one, Transition included');
  assert.ok(after.findIndex(r=>/Slide\s+Slide 2: /.test(r))<two,'module two is below all the slides, not above them');
  assert.equal(a.adapter.snapshot().at(-1).style,'Normal');
});

test('slides in the middle of a module leave a way out of their notes: Write below this block puts a plain line under them, ready to type in',async()=>{
  // The author's recording of 2026-11-08 ("test 3r"): a Say block, Import slides, then Enter at the end of the slide's notes only made more
  // notes lines, and the Transition just below could not be passed. Enter cannot be changed, so the pane has its own way out.
  const a=await started();
  await addModule(a,'Module one');
  await addBlock(a,'say','test');
  await loadDeck(a,await deckBytes(),'deck.pptx',[]);
  slideBox(a,2).checked=true;slideBox(a,2).onchange();
  await click(a,'guide-slides-go');
  assert.equal(a.err(),'');
  const kinds=()=>a.adapter.snapshot().map(r=>r.style.replace(/^OSG /,'')+(/Transition/.test(r.text)?'!':''));
  const lastNotes=()=>a.adapter.snapshot().map((r,i)=>r.style==='OSG Facilitator Notes'?i:-1).filter(i=>i>=0).at(-1);
  const before=kinds();
  assert.equal(a.n('guide-selected').hidden,false);
  assert.match(a.n('guide-selected-text').textContent,/^Your cursor is in the Slide block\.$/);
  assert.equal(a.n('guide-leave-block').hidden,false,'a block offers the way out');
  assert.equal(a.n('guide-selected-hint').hidden,false);
  assert.match(a.n('guide-selected-hint').textContent,/Enter starts another line of notes/);
  const lineAt=lastNotes()+1;
  assert.match(a.adapter.snapshot()[lineAt].text,/Transition/,'nothing plain stands between the slide notes and the Transition');
  await click(a,'guide-leave-block');
  assert.equal(a.err(),'');
  const snap=a.adapter.snapshot();
  assert.equal(snap.length,before.length+1,'one line was added');
  assert.deepEqual([snap[lineAt].style,snap[lineAt].text],['Normal',''],'a plain empty line follows the notes');
  assert.match(snap[lineAt+1].text,/Transition/,'the Transition is still just below it');
  assert.equal(await a.adapter.cursorIndex(),lineAt,'the cursor is in the plain line');
  assert.match(a.n('guide-status').textContent,/A plain line was added below the Slide block, and your cursor is in it\. .*belongs to no block/);
  assert.equal(a.n('guide-selected').hidden,true,'the cursor is outside any block now');
  // Another block added from here goes below the line and above the Transition, never into the notes.
  await addBlock(a,'say','after the slides');
  const rows=guideRows(a),line=rows.findIndex(r=>r==='Normal: '),say=rows.findIndex(r=>/after the slides/.test(r)),tr=rows.findIndex(r=>/Transition/.test(r));
  assert.ok(line>0&&say>line&&tr>say,'plain line, then the new block, then the Transition');
});

test('Write below this block uses a plain line that is already there, and is only offered in a block',async()=>{
  const a=await started();
  await addModule(a,'Module one');await addLesson(a,'Lesson one');
  await addBlock(a,'say','one');
  const say=a.adapter.snapshot().findIndex(r=>/one$/.test(r.text)&&r.style==='OSG Cue Facilitator');
  assert.ok(say>0);
  await click(a,'guide-leave-block');
  const withLine=a.adapter.snapshot().length;
  a.adapter.place(say);await a.p.quietRefresh();
  await click(a,'guide-leave-block');
  assert.equal(a.adapter.snapshot().length,withLine,'no second line is made');
  assert.equal(await a.adapter.cursorIndex(),say+1);
  assert.match(a.n('guide-status').textContent,/Your cursor is on the plain line below the Say block\./);
  // On a heading there is no block, so there is nothing to write below.
  a.adapter.place(a.adapter.snapshot().findIndex(r=>r.style==='OSG Lesson'));await a.p.quietRefresh();
  assert.equal(a.n('guide-selected').hidden,false);
  assert.equal(a.n('guide-leave-block').hidden,true);assert.equal(a.n('guide-selected-hint').hidden,true);
});
