import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {CARD_LIMITS,CARD_CAPACITY,PAPERS,PER_PAGE,checkDeck,cardFit,cardSheetFiles,deckText,deckFileName} from '../src/card-print.mjs';

// ---- helpers
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const EMU=635;
const COLS={1:1,2:1,4:2,6:2,8:2},ROWS={1:1,2:2,4:2,6:3,8:4};
const parse=(text,name='part')=>{
  const problems=[];
  const doc=new DOMParser({errorHandler:{warning:e=>problems.push(String(e)),error:e=>problems.push(String(e)),fatalError:e=>problems.push(String(e))}}).parseFromString(text,'text/xml');
  assert.deepEqual(problems,[],`${name} is well-formed`);
  return doc;
};
const kids=(el,tag)=>Array.from(el.childNodes).filter(n=>n.nodeType===1&&n.localName===tag&&n.namespaceURI===W);
const all=(el,tag,ns=W)=>Array.from(el.getElementsByTagNameNS(ns,tag));
const val=(el,tag,attr='val')=>{const e=all(el,tag)[0];return e?e.getAttributeNS(W,attr):null;};
const textOf=el=>all(el,'t').map(t=>t.textContent).join('');
const partOf=(files,name)=>files.find(f=>f.name===name)?.text;
const docOf=(m,o)=>parse(partOf(cardSheetFiles(m,o),'word/document.xml'),'document.xml');
const tablesOf=doc=>all(doc,'tbl');
const rowsOf=tbl=>kids(tbl,'tr');
const cellsOf=row=>kids(row,'tc');
// Card text by row and column, then the label strips under them.
const cardsGrid=tbl=>rowsOf(tbl).filter((_,i)=>i%2===0).map(r=>cellsOf(r).map(textOf));
const stripsGrid=tbl=>rowsOf(tbl).filter((_,i)=>i%2===1).map(r=>cellsOf(r).map(textOf));
// The tests of the pages' own layout read them fronts first; the order of the pages has its own tests (the last ones in this file), which leave `sequence` out.
const deck=(over={})=>({title:'Scenario cards',type:'activity',paper:'Letter',perPage:4,cropMarks:false,flip:'long',sequence:'separate',cards:cardsOf(4),...over});
const cardsOf=(n,backs=true)=>Array.from({length:n},(_,i)=>({front:[`F${i+1}`],back:backs?[`B${i+1}`]:null}));
const zipOf=async files=>{const z=new JSZip();for(const f of files)z.file(f.name,f.text,{createFolders:false});return z;};
const text=n=>'abcdefghi '.repeat(Math.ceil(n/10)).slice(0,n).trim().padEnd(n,'x');
// n letters spread over as few lines as will do (a line holds up to 300), the way a long card is written.
const lines=(n,k=Math.ceil(n/300))=>Array.from({length:k},(_,i)=>text(Math.floor(n/k)+(i<n%k?1:0)));

test('every part is well-formed, has a content type, and every relationship points at a part that is there',async()=>{
  for(const m of [deck(),deck({perPage:6,cropMarks:true,paper:'A4',flip:'short'}),deck({perPage:1,cards:cardsOf(3,false)})]){
    const files=cardSheetFiles(m);
    for(const f of files)parse(f.text,f.name);
    const zip=await JSZip.loadAsync(await(await zipOf(files)).generateAsync({type:'nodebuffer'}));
    const names=Object.keys(zip.files);
    assert.deepEqual(names.sort(),files.map(f=>f.name).sort(),'the package holds the parts and no folder entries');
    assert.ok(names.every(n=>!zip.files[n].dir));
    const types=parse(partOf(files,'[Content_Types].xml'));
    const overrides=new Set(all(types.documentElement,'Override','http://schemas.openxmlformats.org/package/2006/content-types').map(o=>o.getAttribute('PartName')));
    for(const n of names)if(n!=='[Content_Types].xml'&&!n.endsWith('.rels'))assert.ok(overrides.has(`/${n}`),`${n} has a content type`);
    for(const [rels,from] of [['_rels/.rels',''],['word/_rels/document.xml.rels','word/']]){
      const list=all(parse(partOf(files,rels)).documentElement,'Relationship','http://schemas.openxmlformats.org/package/2006/relationships');
      assert.ok(list.length);
      for(const r of list)assert.ok(names.includes(path.posix.normalize(from+r.getAttribute('Target'))),`${rels}: ${r.getAttribute('Target')} exists`);
    }
    const ids=all(parse(partOf(files,'word/_rels/document.xml.rels')).documentElement,'Relationship','http://schemas.openxmlformats.org/package/2006/relationships').map(r=>r.getAttribute('Id'));
    for(const e of all(parse(partOf(files,'word/document.xml')).documentElement,'headerReference'))assert.ok(ids.includes(e.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')));
  }
});

test('Letter and A4 give their own page size, with half-inch side margins',()=>{
  for(const [paper,w,h] of [['Letter',12240,15840],['A4',11906,16838]]){
    const doc=docOf(deck({paper}));
    const size=all(doc,'pgSz');
    assert.equal(size.length,1);
    assert.equal(size[0].getAttributeNS(W,'w'),String(w));
    assert.equal(size[0].getAttributeNS(W,'h'),String(h));
    const mar=all(doc,'pgMar')[0];
    assert.equal(mar.getAttributeNS(W,'left'),'720');
    assert.equal(mar.getAttributeNS(W,'right'),'720');
  }
});

test('there are ceil(cards / per page) pages of fronts and the same of backs, in a grid of the right shape',()=>{
  for(const perPage of PER_PAGE){
    for(const [n,backs] of [[1,true],[7,true],[13,true],[13,false],[120,true]]){
      const doc=docOf(deck({perPage,cards:cardsOf(n,backs)}));
      const pages=Math.ceil(n/perPage),tbls=tablesOf(doc);
      assert.equal(tbls.length,backs?2*pages:pages,`${n} cards, ${perPage} to a page`);
      for(const t of tbls){
        assert.equal(all(t,'gridCol').length,COLS[perPage]);
        assert.equal(rowsOf(t).length,2*ROWS[perPage],'a card row and a label strip for each row of cards');
        assert.ok(rowsOf(t).every(r=>cellsOf(r).length===COLS[perPage]));
      }
      const captions=tbls.map(t=>val(t,'tblCaption'));
      assert.deepEqual(captions.filter(c=>c.startsWith('Fronts')).length,pages);
      assert.deepEqual(captions.slice(0,pages).every(c=>c.startsWith('Fronts')),true,'every front page comes first');
    }
  }
});

test('the body is an anchor paragraph and a grid for each page, a page break before every page but the first, and no text outside the grids',()=>{
  const doc=docOf(deck({perPage:6,cards:cardsOf(8)}));
  const body=all(doc,'body')[0],parts=Array.from(body.childNodes).filter(n=>n.nodeType===1).map(n=>n.localName);
  assert.deepEqual(parts,['p','tbl','p','tbl','p','tbl','p','tbl','p','sectPr']);
  const paras=kids(body,'p');
  assert.deepEqual(paras.map(p=>all(p,'pageBreakBefore').length),[0,1,1,1,0],'the last paragraph only closes the document');
  assert.ok(paras.every(p=>textOf(p)===''),'the title and the instructions are in the header, not on a page of cards');
  assert.deepEqual(all(doc,'sectPr').length,1);
  assert.equal(all(doc,'titlePg').length,1);
});

test('flipping on the long edge mirrors each row of the back left to right (4 up and 6 up)',()=>{
  let t=tablesOf(docOf(deck({perPage:4,flip:'long'})));
  assert.deepEqual(cardsGrid(t[0]),[['F1','F2'],['F3','F4']]);
  assert.deepEqual(cardsGrid(t[1]),[['B2','B1'],['B4','B3']]);
  assert.deepEqual(stripsGrid(t[1]),[['Card 2 of 4 · Back','Card 1 of 4 · Back'],['Card 4 of 4 · Back','Card 3 of 4 · Back']],'a label stays under its own card');
  t=tablesOf(docOf(deck({perPage:6,flip:'long',cards:cardsOf(6)})));
  assert.deepEqual(cardsGrid(t[0]),[['F1','F2'],['F3','F4'],['F5','F6']]);
  assert.deepEqual(cardsGrid(t[1]),[['B2','B1'],['B4','B3'],['B6','B5']]);
});

test('flipping on the short edge mirrors the rows of the back top to bottom (4 up and 6 up)',()=>{
  let t=tablesOf(docOf(deck({perPage:4,flip:'short'})));
  assert.deepEqual(cardsGrid(t[0]),[['F1','F2'],['F3','F4']]);
  assert.deepEqual(cardsGrid(t[1]),[['B3','B4'],['B1','B2']]);
  t=tablesOf(docOf(deck({perPage:6,flip:'short',cards:cardsOf(6)})));
  assert.deepEqual(cardsGrid(t[1]),[['B5','B6'],['B3','B4'],['B1','B2']]);
  t=tablesOf(docOf(deck({perPage:8,flip:'short',cards:cardsOf(8)})));
  assert.deepEqual(cardsGrid(t[1]),[['B7','B8'],['B5','B6'],['B3','B4'],['B1','B2']]);
});

test('with one column nothing mirrors on the long edge, with one row nothing mirrors on the short edge, and one to a page keeps each card on its own page',()=>{
  let t=tablesOf(docOf(deck({perPage:2,flip:'long',cards:cardsOf(2)})));
  assert.deepEqual(cardsGrid(t[1]),[['B1'],['B2']]);
  t=tablesOf(docOf(deck({perPage:2,flip:'short',cards:cardsOf(2)})));
  assert.deepEqual(cardsGrid(t[1]),[['B2'],['B1']]);
  for(const flip of ['long','short']){
    t=tablesOf(docOf(deck({perPage:1,flip,cards:cardsOf(3)})));
    assert.deepEqual(t.map(x=>cardsGrid(x)),[[['F1']],[['F2']],[['F3']],[['B1']],[['B2']],[['B3']]]);
  }
});

test('a last page that is not full keeps the whole grid, and the backs still land behind their fronts',()=>{
  let t=tablesOf(docOf(deck({perPage:6,flip:'long',cards:cardsOf(8)})));
  assert.deepEqual(cardsGrid(t[1]),[['F7','F8'],['',''],['','']]);
  assert.deepEqual(cardsGrid(t[3]),[['B8','B7'],['',''],['','']]);
  t=tablesOf(docOf(deck({perPage:6,flip:'short',cards:cardsOf(8)})));
  assert.deepEqual(cardsGrid(t[3]),[['',''],['',''],['B7','B8']]);
});

test('a card without a back prints a blank back cell, label and all; one with an empty list of lines counts as having no back',()=>{
  const cards=cardsOf(4);cards[1].back=null;cards[2].back=[];
  const t=tablesOf(docOf(deck({cards})));
  assert.deepEqual(cardsGrid(t[1]),[['',  'B1'],['B4','']]);
  assert.deepEqual(stripsGrid(t[1]),[['','Card 1 of 4 · Back'],['Card 4 of 4 · Back','']]);
  assert.equal(checkDeck(deck({cards})).cards[2].back,null);
});

test('a deck with no backs prints no back pages and does not ask for two-sided printing',()=>{
  const files=cardSheetFiles(deck({cards:cardsOf(9,false)}));
  const t=tablesOf(parse(partOf(files,'word/document.xml')));
  assert.equal(t.length,3);
  assert.ok(t.every(x=>val(x,'tblCaption').startsWith('Fronts')));
  const head=textOf(parse(partOf(files,'word/header1.xml')));
  assert.match(head,/Print on Letter paper\. Cut along the dashed lines\./);
  assert.doesNotMatch(head,/two-sided|backs/);
});

test('the grid is fixed, cards are the same size and cannot split, the cut lines are dashed grey, and the grid is centred so both sides line up',()=>{
  for(const paper of PAPERS)for(const perPage of PER_PAGE){
    const doc=docOf(deck({paper,perPage,cards:cardsOf(9)}));
    const size=all(doc,'pgSz')[0],mar=all(doc,'pgMar')[0],pageW=Number(size.getAttributeNS(W,'w')),pageH=Number(size.getAttributeNS(W,'h'));
    const n=a=>Number(mar.getAttributeNS(W,a));
    const heights=new Set(),strips=new Set();
    for(const t of tablesOf(doc)){
      assert.equal(val(t,'tblLayout','type'),'fixed');
      assert.equal(val(t,'jc'),'center');
      assert.ok(val(t,'tblCaption')&&/only lays out/.test(val(t,'tblDescription')),'the table says it only lays the cards out');
      const b=all(t,'tblBorders')[0];
      for(const side of ['top','left','bottom','right','insideH','insideV']){
        const e=kids(b,side)[0];
        assert.equal(e.getAttributeNS(W,'val'),'dashed');
        assert.match(e.getAttributeNS(W,'color'),/^[A-F0-9]{6}$/);
      }
      rowsOf(t).forEach((r,i)=>{
        assert.equal(all(r,'cantSplit').length,1);
        const h=all(r,'trHeight')[0];
        assert.equal(h.getAttributeNS(W,'hRule'),'exact');
        (i%2?strips:heights).add(h.getAttributeNS(W,'val'));
      });
      const cellW=Number(all(t,'tcW')[0].getAttributeNS(W,'w')),tableW=Number(all(t,'tblW')[0].getAttributeNS(W,'w'));
      assert.equal(tableW,cellW*COLS[perPage]);
      assert.ok(Math.abs((pageW-tableW)/2-n('left'))<=1,'centred across the page');
      const gridH=rowsOf(t).reduce((a,r)=>a+Number(all(r,'trHeight')[0].getAttributeNS(W,'val')),0),top=n('top')+20;
      assert.ok(Math.abs(top-(pageH-top-gridH))<=1,'the space above the grid equals the space below it');
      assert.ok(top>=1440,'room for the header and the crop marks above');
      assert.ok(pageH-n('top')-n('bottom')>=gridH+40+100,'the page holds the grid with room to spare, so it never spills onto a page of its own');
    }
    assert.equal(heights.size,1,'every card is the same height');
    assert.deepEqual([...strips],['300']);
  }
});

test('crop marks appear only when asked, as 0.25 inch straight-line shapes behind the text at the extension of every card edge',()=>{
  assert.equal(all(docOf(deck()),'anchor',WP).length,0);
  assert.doesNotMatch(partOf(cardSheetFiles(deck()),'word/document.xml'),/wps:wsp|Crop mark/);
  for(const [paper,perPage] of [['Letter',6],['A4',4],['Letter',1],['A4',8],['Letter',2]]){
    const doc=docOf(deck({paper,perPage,cropMarks:true,cards:cardsOf(perPage+1)}));
    const tbls=tablesOf(doc),per=2*(COLS[perPage]+1)+2*(ROWS[perPage]+1);
    assert.equal(all(doc,'anchor',WP).length,per*tbls.length);
    const ids=all(doc,'docPr',WP).map(e=>e.getAttribute('id'));
    assert.equal(new Set(ids).size,ids.length,'every shape has its own id');
    const mar=all(doc,'pgMar')[0],pageW=Number(all(doc,'pgSz')[0].getAttributeNS(W,'w'));
    for(const t of tbls){
      const anchorPara=t.previousSibling;
      assert.equal(val(anchorPara,'pStyle'),'OSGCardAnchor','the marks of a page are in the paragraph before its grid');
      const marks=all(anchorPara,'anchor',WP).map(a=>({
        x:Number(a.getElementsByTagNameNS(WP,'positionH')[0].textContent)/EMU,y:Number(a.getElementsByTagNameNS(WP,'positionV')[0].textContent)/EMU,
        cx:Number(a.getElementsByTagNameNS(WP,'extent')[0].getAttribute('cx')),cy:Number(a.getElementsByTagNameNS(WP,'extent')[0].getAttribute('cy')),
        behind:a.getAttribute('behindDoc'),h:a.getElementsByTagNameNS(WP,'positionH')[0].getAttribute('relativeFrom'),v:a.getElementsByTagNameNS(WP,'positionV')[0].getAttribute('relativeFrom')}));
      assert.equal(marks.length,per);
      assert.ok(marks.every(k=>k.behind==='1'&&k.h==='page'&&k.v==='page'));
      assert.ok(marks.every(k=>(k.cx===0&&k.cy===360*EMU)||(k.cy===0&&k.cx===360*EMU)),'each is a 0.25 inch line, upright or flat');
      const cellW=Number(all(t,'tcW')[0].getAttributeNS(W,'w')),tableW=Number(all(t,'tblW')[0].getAttributeNS(W,'w'));
      const left=(pageW-tableW)/2,top=Number(mar.getAttributeNS(W,'top'))+20;
      const pitch=Number(all(t,'trHeight')[0].getAttributeNS(W,'val'))+300,gridH=pitch*ROWS[perPage];
      const upright=marks.filter(k=>k.cx===0).map(k=>`${k.x},${k.y}`).sort(),flat=marks.filter(k=>k.cy===0).map(k=>`${k.x},${k.y}`).sort();
      const wantUp=[],wantFlat=[];
      for(let c=0;c<=COLS[perPage];c++)wantUp.push(`${left+c*cellW},${top-360}`,`${left+c*cellW},${top+gridH}`);
      for(let r=0;r<=ROWS[perPage];r++)wantFlat.push(`${left-360},${top+r*pitch}`,`${left+tableW},${top+r*pitch}`);
      assert.deepEqual(upright,wantUp.sort(),'above and below each column line');
      assert.deepEqual(flat,wantFlat.sort(),'left and right of each row line');
    }
    const raw=partOf(cardSheetFiles(deck({paper,perPage,cropMarks:true,cards:cardsOf(perPage+1)})),'word/document.xml');
    assert.ok(raw.includes('<mc:Choice Requires="wps">')&&raw.includes('prst="line"'));
  }
});

test('the header of the first page holds the title and the printing instructions, and every other page is all cards',()=>{
  for(const [over,line] of [[{},'Print on Letter paper, two-sided, flip on the long edge. Cut along the dashed lines.'],[{paper:'A4',flip:'short'},'Print on A4 paper, two-sided, flip on the short edge. Cut along the dashed lines.']]){
    const files=cardSheetFiles(deck(over));
    const head=parse(partOf(files,'word/header1.xml'),'header1.xml'),ps=kids(head.documentElement,'p');
    assert.equal(val(ps[0],'pStyle'),'Title');
    assert.equal(textOf(ps[0]),'Print sheet: Scenario cards');
    assert.equal(textOf(ps[1]),line);
    assert.equal(textOf(ps[2]),'Page 1 is the front. Page 2 is the back.');
    const doc=parse(partOf(files,'word/document.xml'));
    const ref=all(doc,'headerReference')[0];
    assert.equal(ref.getAttributeNS(W,'type'),'first');
    assert.equal(all(doc,'titlePg').length,1);
    assert.match(partOf(files,'word/_rels/document.xml.rels'),new RegExp(`Id="${ref.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')}"[^>]*Target="header1.xml"`));
    assert.match(partOf(files,'[Content_Types].xml'),/PartName="\/word\/header1\.xml"/);
    const top=Number(all(doc,'pgMar')[0].getAttributeNS(W,'top'))+20,at=Number(all(doc,'pgMar')[0].getAttributeNS(W,'header'));
    assert.ok(at+268+2*196<=top-360,'the header ends above the crop marks, so it never pushes the first page down');
  }
  const long=cardSheetFiles(deck({title:'T'.repeat(120),cards:cardsOf(8)}));
  const second=textOf(kids(parse(partOf(long,'word/header1.xml')).documentElement,'p')[2]);
  assert.equal(second,'Pages 1 to 2 are the fronts. Pages 3 to 4 are the backs.');
  assert.match(textOf(kids(parse(partOf(long,'word/header1.xml')).documentElement,'p')[0]),/^Print sheet: T{55}…$/,'a long title is shortened in the header only');
  assert.match(partOf(long,'docProps/core.xml'),new RegExp(`<dc:title>Print sheet: ${'T'.repeat(120)}</dc:title>`));
});

test('the document title and language are set, and the language can be chosen',()=>{
  let core=partOf(cardSheetFiles(deck()),'docProps/core.xml');
  assert.match(core,/<dc:title>Print sheet: Scenario cards<\/dc:title>/);
  assert.match(core,/<dc:language>en-US<\/dc:language>/);
  const files=cardSheetFiles(deck(),{language:'fr-CA'});
  assert.match(partOf(files,'docProps/core.xml'),/<dc:language>fr-CA<\/dc:language>/);
  assert.match(partOf(files,'word/styles.xml'),/w:lang w:val="fr-CA"/);
  assert.match(partOf(cardSheetFiles(deck({language:'de-DE'})),'docProps/core.xml'),/de-DE/);
  assert.throws(()=>cardSheetFiles(deck(),{language:'english please'}),/language/);
});

test('cards are real text in real styles: 14 pt for 1 to 4 to a page, 12 pt for 6 and 8, the first line of a front in bold',()=>{
  const cards=[{front:['Heading line','A second line'],back:['Back line one','Back line two']}];
  for(const perPage of PER_PAGE){
    const files=cardSheetFiles(deck({perPage,cards}));
    const styles=parse(partOf(files,'word/styles.xml'),'styles.xml');
    const byId=id=>all(styles,'style').find(s=>s.getAttributeNS(W,'styleId')===id);
    assert.equal(val(byId('OSGCardText'),'sz'),perPage<=4?'28':'24');
    assert.equal(all(byId('OSGCardHead'),'b').length,1);
    assert.equal(val(byId('OSGCardLabel'),'sz'),'16');
    assert.match(partOf(files,'word/styles.xml'),/w:styleId="Title"/);
    const doc=parse(partOf(files,'word/document.xml'));
    const [front,back]=tablesOf(doc),styleOf=cell=>kids(cell,'p').map(p=>val(p,'pStyle'));
    assert.deepEqual(styleOf(cellsOf(rowsOf(front)[0])[0]),['OSGCardHead','OSGCardText']);
    assert.deepEqual(styleOf(cellsOf(rowsOf(back)[0]).find(c=>textOf(c))),['OSGCardText','OSGCardText'],'a back has no bold first line');
    assert.deepEqual(stripsGrid(front)[0].filter(Boolean),['Card 1 of 1']);
    assert.equal(all(doc,'blip','http://schemas.openxmlformats.org/drawingml/2006/main').length,0,'no pictures');
  }
});

test('text with XML special characters is escaped everywhere it is written',async()=>{
  const nasty=`Q&A <b>"quoted"</b> 'single' & more`;
  const m=deck({title:nasty,cards:[{front:[nasty,'x < y && y > z'],back:['<script>alert(1)</script>']}]});
  const files=cardSheetFiles(m);
  for(const f of files)parse(f.text,f.name);
  const doc=parse(partOf(files,'word/document.xml'));
  assert.equal(textOf(cellsOf(rowsOf(tablesOf(doc)[0])[0])[0]),`${nasty}x < y && y > z`);
  assert.equal(textOf(cellsOf(rowsOf(tablesOf(doc)[1])[0])[1]),'<script>alert(1)</script>');
  assert.match(partOf(files,'docProps/core.xml'),/Print sheet: Q&amp;A &lt;b&gt;&quot;quoted&quot;&lt;\/b&gt;/);
  assert.equal(textOf(kids(parse(partOf(files,'word/header1.xml')).documentElement,'p')[0]),`Print sheet: ${nasty}`);
  assert.doesNotMatch(partOf(files,'word/document.xml'),/<script>|<b>/);
});

test('text that will not fit a card at its print size is flagged, and text that will is not',()=>{
  const label='Card 4 of 4'.length,small={front:[text(120)],back:null};
  const flags=cardFit(deck({cards:[small,small,small,{front:lines(410,2),back:null}]}));
  assert.equal(flags.length,1);
  assert.equal(flags[0].card,4);
  assert.equal(flags[0].side,'front');
  assert.equal(flags[0].chars,410+40+label,'the letters, a line break, and the label');
  assert.equal(flags[0].limit,CARD_CAPACITY[4]);
  assert.equal(flags[0].message,`Card 4 (front) has ${flags[0].chars} characters, which is more than fits when 4 cards go on a page. Shorten it, or print fewer cards to a page.`);
  assert.deepEqual(cardFit(deck({cards:[small,small]})),[]);
  assert.deepEqual(cardFit(deck({cards:[{front:[text(120)],back:[text(120)]}]})),[],'120 characters fit at every size');
  for(const perPage of PER_PAGE)assert.deepEqual(cardFit(deck({perPage,cards:[small]})),[]);
  // line breaks count 40 each, and the label counts
  assert.equal(cardFit(deck({cards:[{front:[text(200),text(100)],back:null}],perPage:8}))[0].chars,200+100+40+'Card 1 of 1'.length);
  // the back is checked on its own, with its own label
  const back=cardFit(deck({cards:[{front:['short'],back:lines(400,2)}]}));
  assert.deepEqual(back.map(f=>[f.card,f.side,f.chars]),[[1,'back',400+40+'Card 1 of 1 · Back'.length]]);
  // the same text is fine on a bigger card and flagged on a smaller one
  const big={front:[text(300),text(110)],back:null};
  assert.deepEqual(cardFit(deck({perPage:1,cards:[big]})),[]);
  assert.equal(cardFit(deck({perPage:8,cards:[big]})).length,1);
  assert.match(cardFit(deck({perPage:1,cards:[{front:Array.from({length:8},()=>text(300)),back:null}]}))[0].message,/on a page by itself\. Shorten it\.$/);
  // the capacity itself is not over the capacity, and one letter more is
  const room=CARD_CAPACITY[4]-'Card 1 of 1'.length-40;
  assert.deepEqual(cardFit(deck({cards:[{front:lines(room,2),back:null}]})),[]);
  assert.equal(cardFit(deck({cards:[{front:lines(room+1,2),back:null}]})).length,1);
});

// The table of capacities, proved against the layout the module builds: in the narrowest card (A4 width) and the shortest (Letter height),
// with a wide average letter (0.55 em) and a line 1.22 times the type size, the text a capacity allows still fits, however it is split into lines.
test('the capacity table holds for the layout: what it allows fits the card, and it does not flag a card that fits easily',()=>{
  const widest='Card 120 of 120 · Back'.length;
  for(const perPage of PER_PAGE){
    const letter=docOf(deck({perPage,paper:'Letter'})),a4=docOf(deck({perPage,paper:'A4'}));
    const styles=parse(partOf(cardSheetFiles(deck({perPage})),'word/styles.xml'));
    const style=id=>all(styles,'style').find(s=>s.getAttributeNS(W,'styleId')===id);
    const pt=Number(val(style('OSGCardText'),'sz'))/2,after=Number(val(style('OSGCardText'),'spacing','after'))/20;
    const mar=all(letter,'tblCellMar')[0],m=a=>Number(val(mar,a,'w'));
    const innerW=(Number(all(a4,'tcW')[0].getAttributeNS(W,'w'))-m('left')-m('right'))/20;
    const innerH=(Number(all(letter,'trHeight')[0].getAttributeNS(W,'val'))-m('top')-m('bottom'))/20;
    const perLine=Math.floor(innerW/(0.55*pt)),lines=Math.floor(innerH/(1.22*pt)),cap=CARD_CAPACITY[perPage];
    assert.ok(cap<=perLine*lines*0.8,`${perPage} up: the capacity ${cap} is well inside the ${perLine*lines} letters the card can hold`);
    assert.ok(cap>=perLine*lines*0.5,`${perPage} up: the capacity is not needlessly small`);
    for(let k=1;k<=CARD_LIMITS.lines;k++){
      const letters=cap-widest-40*(k-1),each=Math.ceil(letters/k);
      const used=k*Math.ceil(each/perLine)*1.22*pt+k*after;
      assert.ok(used<=innerH,`${perPage} up, ${k} lines: ${used.toFixed(0)} pt used of ${innerH.toFixed(0)} pt`);
    }
  }
  assert.deepEqual(Object.keys(CARD_CAPACITY).map(Number),[...PER_PAGE]);
  assert.ok(Object.isFrozen(CARD_CAPACITY));
  assert.ok(PER_PAGE.every((p,i)=>i===0||CARD_CAPACITY[p]<CARD_CAPACITY[PER_PAGE[i-1]]),'smaller cards hold less');
});

test('the limits are kept and bad input is refused in plain English',()=>{
  assert.deepEqual({...CARD_LIMITS},{cards:120,lines:8,line:300});
  assert.ok(Object.isFrozen(CARD_LIMITS));
  assert.deepEqual([...PAPERS],['Letter','A4']);
  assert.deepEqual([...PER_PAGE],[1,2,4,6,8]);
  const bad=(over,message)=>{
    for(const f of [checkDeck,cardSheetFiles,cardFit,deckText])assert.throws(()=>f(deck(over)),e=>e instanceof Error&&(message instanceof RegExp?message.test(e.message):e.message===message),JSON.stringify(Object.keys(over)));
  };
  bad({cards:[]},'Add at least one card.');
  bad({cards:undefined},'Add at least one card.');
  bad({cards:cardsOf(121)},'A deck holds up to 120 cards.');
  bad({cards:[{front:Array.from({length:9},(_,i)=>`line ${i}`),back:null}]},/front of card 1 has 9 lines\. A side holds up to 8 lines\./);
  bad({cards:[{front:['a'],back:Array.from({length:9},(_,i)=>`line ${i}`)}]},/back of card 1 has 9 lines/);
  bad({cards:[{front:[],back:null}]},'Card 1 needs at least one line on its front.');
  bad({cards:[{front:['a','  '],back:null}]},'Line 2 on the front of card 1 is empty.');
  bad({cards:[{front:['a'],back:['']}]},'Line 1 on the back of card 1 is empty.');
  bad({cards:[{front:['x'.repeat(301)],back:null}]},/has 301 letters\. A line holds up to 300 letters\./);
  bad({cards:[{front:'not a list',back:null}]},/must be a list of lines/);
  bad({paper:'Legal'},'Choose Letter or A4 paper.');
  bad({perPage:3},'Cards go 1, 2, 4, 6 or 8 to a page.');
  bad({flip:'diagonal'},'Choose to flip on the long edge or the short edge.');
  bad({type:'poster'},'A deck is for an activity, a toolkit or a facilitator.');
  bad({title:''},/title/);
  bad({title:'x'.repeat(201)},/title/);
  bad({cropMarks:'yes'},'Crop marks are either on or off.');
  assert.throws(()=>checkDeck(null),/no deck/);
  assert.equal(checkDeck(deck({cards:cardsOf(120)})).cards.length,120);
  assert.equal(checkDeck(deck({cards:[{front:Array(8).fill('y'.repeat(300)),back:Array(8).fill('z'.repeat(300))}]})).cards[0].back.length,8);
  assert.equal(checkDeck(deck({title:'x'.repeat(200)})).title.length,200);
});

test('checkDeck returns a cleaned copy and leaves the deck it was given alone',()=>{
  const input=deck({title:'  Tidy\u0000 title  ',cards:[{front:['  one\ttwo\nthree ','x'],back:null}]});
  const frozen=JSON.stringify(input);
  const out=checkDeck(input);
  assert.equal(JSON.stringify(input),frozen);
  assert.equal(out.title,'Tidy title');
  assert.deepEqual(out.cards[0].front,['one two three','x']);
  assert.equal(out.language,'en-US');
  assert.equal(out.cropMarks,false);
  out.cards[0].front.push('changed');
  assert.equal(input.cards[0].front.length,2);
  const lone=checkDeck(deck({cards:[{front:['a\uD800b'],back:null}]}));
  assert.equal(lone.cards[0].front[0],'a b','half a surrogate pair is replaced');
  assert.equal(checkDeck(deck({cards:[{front:['smile \u{1F600} ok'],back:null}]})).cards[0].front[0],'smile \u{1F600} ok','a whole pair is kept');
  assert.deepEqual(cardSheetFiles(input),cardSheetFiles(input),'the same deck gives the same parts');
});

test('deckText is a linear view with a block for each card',()=>{
  const out=deckText(deck({cards:[{front:['Scenario one','a second line'],back:['Answer one']},{front:['Two'],back:null}]}));
  assert.equal(out,'Card 1\nFront: Scenario one\n  a second line\nBack: Answer one\n\nCard 2\nFront: Two\nBack: (blank)\n');
  assert.ok(out.endsWith('\n')&&!out.endsWith('\n\n')&&!out.includes('\r'));
});

test('deckFileName gives a safe name that says which paper it is for',()=>{
  assert.equal(deckFileName('Scenario cards','Letter'),'Scenario cards - print sheet (Letter).docx');
  assert.equal(deckFileName('Café cards','A4'),'Café cards - print sheet (A4).docx');
  assert.equal(deckFileName('a/b\\c:d*e?f"g<h>i|j','A4'),'abcdefghij - print sheet (A4).docx');
  assert.equal(deckFileName('  ','A4'),'Cards - print sheet (A4).docx');
  assert.equal(deckFileName(undefined,'Letter'),'Cards - print sheet (Letter).docx');
  const long=deckFileName('Very long title '.repeat(20),'Letter');
  assert.ok(long.length<=80);
  assert.ok(long.endsWith(' - print sheet (Letter).docx'));
  assert.match(deckFileName('Mixed_name-(1) & more!','A4'),/^[\p{L}\p{N} _()-]+\.docx$/u);
  assert.throws(()=>deckFileName('x','Legal'),/Letter or A4/);
});

// ---- the Word schema validator, when this PC has one: OPENSTAGE_DOCX_VALIDATOR names its validate.py. The sheets are built into a fresh folder
// of their own and checked as data; Python runs isolated (-I) from a launcher that makes the validator's own folder importable.
const VALIDATOR=process.env.OPENSTAGE_DOCX_VALIDATOR;
const LAUNCH='import os,runpy,sys\nv=sys.argv[1]\nsys.path.insert(0,os.path.dirname(v))\nsys.argv=[v]+sys.argv[2:]\nrunpy.run_path(v,run_name="__main__")';
test('the Word schema validator accepts the sheets',{skip:VALIDATOR?false:'OPENSTAGE_DOCX_VALIDATOR is not set, so the Word schema check was not run (set it to the path of the validator\'s validate.py).'},async t=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'card-print-'));
  try{
    const decks={
      'letter-6-up-backs-marks':deck({perPage:6,cropMarks:true,cards:cardsOf(14).map((c,i)=>({front:[`Scenario ${i+1}`,'What do you say <first> & why?'],back:i%3?c.back:null}))}),
      'a4-4-up-short':deck({paper:'A4',perPage:4,flip:'short',cards:cardsOf(9)}),
      'letter-1-up':deck({perPage:1,cropMarks:true,cards:cardsOf(3)})
    };
    for(const [name,m] of Object.entries(decks)){
      const file=path.join(dir,`${name}.docx`);
      fs.writeFileSync(file,await(await zipOf(cardSheetFiles(m))).generateAsync({type:'nodebuffer',compression:'DEFLATE'}));
      const r=spawnSync(process.env.OPENSTAGE_PYTHON||'python',['-I','-c',LAUNCH,VALIDATOR,file],{encoding:'utf8',timeout:180000});
      if(r.error)return t.skip(`Python could not be started (${r.error.message}), so the Word schema check was not run.`);
      const said=`${r.stdout}${r.stderr}`;
      if(r.status!==0&&/ModuleNotFoundError|ImportError/.test(said))return t.skip(`the validator could not start because of a missing Python module, so the Word schema check was not run: ${said.trim().split('\n').pop()}`);
      assert.equal(r.status,0,`${name}: ${said}`);
    }
  }finally{fs.rmSync(dir,{recursive:true,force:true});}
});

// ---- the order of the pages
test('by default each front is followed by its back, which is the order a printer that prints both sides itself takes',()=>{
  const m=deck({cards:cardsOf(6)});delete m.sequence;
  const doc=docOf(m);
  const grids=tablesOf(doc).map(t=>cardsGrid(t).flat().map(s=>s.replace(/Card .*$/,'')));
  assert.equal(grids.length,4);
  assert.match(grids[0].join(' '),/F1/);assert.match(grids[1].join(' '),/B1/);
  assert.match(grids[2].join(' '),/F5/);assert.match(grids[3].join(' '),/B5/);
  const head=partOf(cardSheetFiles(m),'word/header1.xml');
  assert.match(head,/Each front page is followed by its back\./);
  assert.doesNotMatch(head,/are the fronts/);
  assert.equal(checkDeck(m).sequence,'alternate');
});

test('fronts first is a choice, for printing in two passes, and the header says which pages are which',()=>{
  const m=deck({cards:cardsOf(6),sequence:'separate'});
  const grids=tablesOf(docOf(m)).map(t=>cardsGrid(t).flat().join(' '));
  assert.match(grids[0],/F1/);assert.match(grids[1],/F5/);assert.match(grids[2],/B1/);assert.match(grids[3],/B5/);
  assert.match(partOf(cardSheetFiles(m),'word/header1.xml'),/Pages 1 to 2 are the fronts\. Pages 3 to 4 are the backs\./);
  assert.throws(()=>checkDeck(deck({sequence:'sideways'})),/each front with its back/);
});

test('a deck with no backs has the same pages whichever order is chosen',()=>{
  const a=cardSheetFiles(deck({cards:cardsOf(6,false),sequence:'separate'})),b=cardSheetFiles(deck({cards:cardsOf(6,false),sequence:'alternate'}));
  assert.equal(partOf(a,'word/document.xml'),partOf(b,'word/document.xml'));
});