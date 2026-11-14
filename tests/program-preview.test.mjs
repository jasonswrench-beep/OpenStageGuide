// The guide as it reads in the program (program/ui/preview.js): drawn from the project's paragraphs, with text and attributes only.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {describeGuide,renderGuide,drawNodes} from '../program/ui/preview.js';
import {previewStyles} from '../src/guide-styles-edit.mjs';
import {makeContext,para,run,pictureRun} from '../src/guide-ooxml.mjs';
import {formTitleBlock,formText,writingLines,writingBox,fieldRow,checkList,formTable,ratingGrid,sourceLine,printLine,cardSet,citationControl} from '../src/guide-forms-ooxml.mjs';
import {STYLE_IDS} from '../src/guide-styles.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const parseXml=xml=>new DOMParser().parseFromString(xml,'application/xml');
// A page's document as far as the preview needs one: it fails if anything writes markup.
class Fake{
  constructor(tag){this.tag=tag;this.children=[];this.attrs={};this.style={};this.className='';this._text=undefined;}
  setAttribute(k,v){this.attrs[k]=String(v);}
  append(...c){this.children.push(...c);}
  replaceChildren(...c){this.children=[...c];this._text=undefined;}
  set textContent(v){this._text=String(v);this.children=[];}
  get textContent(){return (this._text??'')+this.children.map(c=>c.textContent).join('');}
  set innerHTML(v){throw new Error('markup was written');}
}
const fakeDoc=()=>({createElement:tag=>new Fake(tag)});
const find=(nodes,cls)=>{const out=[];const walk=n=>{if(n.cls&&n.cls.split(' ').includes(cls))out.push(n);(n.children||[]).forEach(walk);};nodes.forEach(walk);return out;};
const text=n=>(n.text??'')+(n.children||[]).map(text).join('');
const tags=(nodes,tag)=>{const out=[];const walk=n=>{if(n.tag===tag)out.push(n);(n.children||[]).forEach(walk);};nodes.forEach(walk);return out;};

async function guide(){
  const a=await started('Forms guide');
  await addModule(a,'Naming it');await addLesson(a,'First','10');await addBlock(a,'ask','What makes it hard?',{notes:'Listen for blame.'});
  const at=a.adapter.snapshot().findIndex(p=>p.text.includes('What makes it hard?'));
  await a.adapter.insert(ctx=>
    formTitleBlock({id:'f1',label:'Handout',number:'1.1',title:'Goals'})+printLine({copies:{n:1,per:'person'},orientation:'landscape'})+formText('Fill this in.')
    +fieldRow([{label:'Name',width:'full'},{label:'Date',width:'short'}])+writingLines({count:2})+writingBox({lines:3})
    +checkList({items:['Read it','Try it'],columns:['When'],ctx})+formTable({columns:[{name:'Goal',weight:3},{name:'Done',kind:'check'}],rows:[['Plan ahead']],ctx,caption:'Goals'})
    +ratingGrid({statements:['I plan.'],evidence:true,ctx})+cardSet({cards:['Role A\nThe manager.']})+sourceLine('Credit.')
    +formTitleBlock({id:'f2',label:'Checklist',number:'1.2',title:'Set-up',facilitator:true})+formText('Facilitator only.')
    +para(STYLE_IDS.body,run('Give out ')+citationControl({id:'f1',text:'Handout 1.1: Goals'},ctx)+run(' now.')),{where:'after',after:at,spare:false});
  return a.adapter.exportState().paragraphs;
}

test('a guide is drawn as its parts: title, module, lesson, time, block with its label, notes, standard pages',async()=>{
  const nodes=describeGuide(await guide(),parseXml);
  assert.equal(text(find(nodes,'g-title')[0]),'Forms guide');
  assert.equal(text(find(nodes,'g-module')[0]),'Module 1: Naming it');
  assert.equal(text(find(nodes,'g-lesson')[0]),'Lesson 1.1: First');
  assert.match(text(find(nodes,'g-time')[0]),/^\s*Time: \d+ minutes?\. Cumulative time: 0 minutes\./);
  const cue=find(nodes,'g-cue')[0];
  assert.ok(find([cue],'g-label').length===1,'the label of a block is drawn as a label');
  assert.match(text(cue),/What makes it hard\?/);
  assert.match(text(find(nodes,'g-notes')[0]),/Listen for blame\./);
  assert.ok(nodes.some(n=>n.tag==='h2'&&text(n)==='Acknowledgements'),'the standard pages are headings');
});

test('forms are drawn with their title, print line, lines, boxes, fields, cards, source and the facilitator\'s form marked',async()=>{
  const nodes=describeGuide(await guide(),parseXml);
  const forms=find(nodes,'g-form');
  assert.deepEqual(forms.map(f=>text(f)),['Handout 1.1: Goals','Checklist 1.2: Set-up']);
  assert.ok(forms[1].cls.includes('facilitator')&&!forms[0].cls.includes('facilitator'));
  assert.equal(text(find(nodes,'g-print')[0]),'Print: 1 per person, landscape.');
  assert.equal(find(nodes,'g-rule').length,2);
  const box=find(nodes,'g-box')[0];assert.equal(box.height,54,'a box of three lines is 54 points tall');
  const field=find(nodes,'g-field')[0];
  assert.deepEqual(field.children.map(c=>[c.cls,c.text]),[['g-fieldlabel','Name'],['g-fieldline',undefined],['g-fieldlabel','Date'],['g-fieldline',undefined]]);
  assert.match(text(find(nodes,'g-card')[0]),/Role A/);
  assert.equal(text(find(nodes,'g-source')[0]),'Credit.');
  assert.equal(text(find(nodes,'g-formtext')[0]),'Fill this in.');
});

test('a table is drawn as a table with header cells where Word marks a header row, a caption, and check boxes in their cells',async()=>{
  const nodes=describeGuide(await guide(),parseXml);
  const tables=find(nodes,'g-table');
  assert.equal(tables.length,3,'a checklist, a table and a rating grid: each once, not once for each cell');
  const goals=tables.find(t=>t.attrs?.['aria-label']==='Goals');
  assert.ok(goals);
  const header=goals.children[0];assert.deepEqual(header.children.map(c=>[c.tag,c.attrs?.scope,text(c)]),[['th','col','Goal'],['th','col','Done']]);
  const body=goals.children[1];assert.deepEqual(body.children.map(c=>c.tag),['td','td']);
  assert.equal(text(body.children[0]),'Plan ahead');
  assert.equal(text(body.children[1]),'☐');
  assert.equal(find(nodes,'g-check').length,2+1+5);
  assert.ok(tags(nodes,'th').length>=6);
});

test('a citation is drawn as a citation with the form\'s name',async()=>{
  const nodes=describeGuide(await guide(),parseXml);
  assert.deepEqual(find(nodes,'g-cite').map(text),['Handout 1.1: Goals']);
  assert.ok(nodes.some(n=>n.cls==='g-body'&&/^Give out Handout 1\.1: Goals now\./.test(text(n))));
});

test('what is not words is not drawn: empty paragraphs, a table\'s later cells, pictures; damaged XML is shown as the words it had',()=>{
  const entries=[
    {xml:'',text:'',style:'Normal'},
    {xml:'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',text:'',style:'Normal'},
    {xml:'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:r><w:drawing><wp:inline/></w:drawing></w:r></w:p>',text:'',style:'OSG Cue'},
    {xml:'this is not xml',text:'Words that were in it',style:'Normal'}
  ];
  const nodes=describeGuide(entries,xml=>{if(!xml.startsWith('<'))throw new Error('damaged');return parseXml(xml);});
  assert.deepEqual(nodes,[{tag:'p',cls:'g-other',text:'Words that were in it',index:3}]);
});

test('the contents lines Update guide writes (plain words, no XML) are drawn as the entries they will be, without the page number only Word can give',async()=>{
  const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const entries=[
    {xml:`<w:p ${W}><w:r><w:t>Contents</w:t></w:r></w:p>`,text:'Contents',style:'heading 1'},
    {xml:'',text:'Module 1: Naming it\t1',style:'TOC 1'},
    {xml:'',text:'Lesson 1.1: First\t1',style:'TOC 2'},
    {xml:'',text:'\t1',style:'TOC 2'},
    {xml:'',text:'Not a contents line',style:'Normal'},
    {xml:`<w:p ${W}><w:r><w:t>Getting Started</w:t></w:r></w:p>`,text:'Getting Started',style:'heading 1'}
  ];
  const nodes=describeGuide(entries,parseXml);
  assert.deepEqual(nodes.filter(n=>/g-toc/.test(n.cls||'')).map(n=>[n.cls,n.text,n.index]),[['g-toc g-toc-1','Module 1: Naming it',1],['g-toc g-toc-2','Lesson 1.1: First',2]],'only lines with words are drawn, at their own level, with no page number');
  assert.equal(nodes.length,4,'the two headings and the two lines; a line with no words and a plain paragraph with no XML are left out');
  // drawn with text only
  const page=new Fake('div');
  drawNodes(fakeDoc(),page,nodes);
  assert.match(page.textContent,/Module 1: Naming itLesson 1\.1: First/);
  assert.doesNotMatch(page.textContent,/\t1/);
  // the same lines are not swallowed by a table that comes before them
  const tbl=`<w:tbl ${W}><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  const after=describeGuide([{xml:tbl,text:'A',style:'OSG Table Head'},{xml:`<w:p ${W}><w:r><w:t>Contents</w:t></w:r></w:p>`,text:'Contents',style:'heading 1'},{xml:'',text:'Module 1\t1',style:'TOC 1'}],parseXml);
  assert.equal(after.length,3,'the table, the heading and the line');
  assert.equal(after.at(-1).cls,'g-toc g-toc-1');
  assert.equal(after.at(-1).text,'Module 1');
});

test('markup in the guide\'s words is drawn as words: nothing is ever written as markup',()=>{
  const evil='<img src=x onerror=alert(1)> & <script>alert(2)</script>';
  const esc=s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const entry={xml:`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="OSGModule"/></w:pPr><w:r><w:t>${esc(evil)}</w:t></w:r></w:p>`,text:evil,style:'OSG Module'};
  const doc=fakeDoc(),container=new Fake('div');
  renderGuide(doc,container,[entry],parseXml);
  const page=container.children[0];
  assert.equal(page.tag,'article');assert.equal(page.attrs['aria-label'],'The guide');
  const h=page.children[0];
  assert.equal(h.tag,'h2');assert.equal(h.textContent,evil,'the words are the words, tags and all');
});

test('an empty guide says how to start, and drawing again replaces what was there',()=>{
  const doc=fakeDoc(),container=new Fake('div');
  renderGuide(doc,container,[],parseXml);
  assert.match(container.textContent,/This guide is empty\. Choose Guide overview/);
  renderGuide(doc,container,[],parseXml);
  assert.equal(container.children.length,1);
  const entries=[{xml:'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>One</w:t></w:r></w:p>',text:'One',style:'Normal'}];
  renderGuide(doc,container,entries,parseXml);
  assert.equal(container.children.length,1);
  assert.equal(container.textContent,'One');
});

test('drawNodes draws a tree with the page\'s document: tags, classes, attributes, a height, text, children',()=>{
  const doc=fakeDoc(),parent=new Fake('div');
  drawNodes(doc,parent,[{tag:'div',cls:'g-box',height:30},{tag:'table',cls:'g-table',attrs:{'aria-label':'T'},children:[{tag:'tr',children:[{tag:'td',text:'x'}]}]}]);
  assert.equal(parent.children[0].className,'g-box');
  assert.equal(parent.children[0].style.height,'40px');
  assert.equal(parent.children[1].attrs['aria-label'],'T');
  assert.equal(parent.children[1].textContent,'x');
});

test('every drawn row says which paragraph it comes from; the editor\'s view also draws paragraphs with nothing in them, and the read view does not',()=>{
  const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const entries=[
    {xml:'',text:'',style:'TOC 1'},
    {xml:`<w:p ${W}><w:r><w:t>One</w:t></w:r></w:p>`,text:'One',style:'Normal'},
    {xml:`<w:p ${W}/>`,text:'',style:'OSG Cue Body'},
    {xml:`<w:p ${W}><w:pPr><w:pStyle w:val="OSGModule"/></w:pPr><w:r><w:t>Module 1: A</w:t></w:r></w:p>`,text:'Module 1: A',style:'OSG Module'},
  ];
  const read=describeGuide(entries,parseXml);
  assert.deepEqual(read.map(n=>n.index),[1,3],'the contents line has no XML and no row; an empty paragraph is left out of the read view');
  const edit=describeGuide(entries,parseXml,{blank:true});
  assert.deepEqual(edit.map(n=>n.index),[1,2,3]);
  const blank=edit[1];
  assert.equal(blank.cls,'g-body g-blank','an empty paragraph keeps the look of its style');
  assert.equal(text(blank),'Empty paragraph','and has words for a screen reader');
  assert.deepEqual(describeGuide([{xml:`<w:p ${W}/>`,text:'',style:'Normal'}],parseXml,{blank:true}).map(n=>n.cls),['g-para g-blank']);
});

test('the author\'s own picture is drawn from the guide\'s copy with its description as alt text; a picture that is not there says so; a cue\'s icon is not drawn',async()=>{
  const a=await started('Pictures');
  await a.adapter.insert(ctx=>para('OSGCueBody',pictureRun({base64:'iVBORw0KGgo=',type:'png',mime:'image/png',size:{cx:1905000,cy:952500},alt:'A cat on a mat',name:'cat'},ctx)),{where:'start',spare:false});
  const entries=a.adapter.exportState().paragraphs;
  const at=entries.findIndex(p=>/r:embed="rIdPic/.test(p.xml));
  assert.ok(at>=0);
  const media=id=>a.adapter.mediaOf(id);
  const nodes=describeGuide(entries,parseXml,{media});
  const row=nodes.find(n=>n.index===at);
  assert.ok(row,'a paragraph that holds only a picture is a row');
  const img=row.children[0];
  assert.equal(img.tag,'img');assert.equal(img.cls,'g-picture');
  assert.match(img.attrs.src,/^data:image\/png;base64,iVBORw0KGgo=$/,'a data address: nothing is fetched');
  assert.equal(img.attrs.alt,'A cat on a mat');
  assert.equal(img.attrs.width,'200','the size the paragraph gives it, in pixels');
  const gone=describeGuide(entries,parseXml,{media:()=>null}).find(n=>n.index===at);
  assert.equal(gone.children[0].cls,'g-picture-missing');assert.equal(text(gone),'[Picture: A cat on a mat]');
  assert.equal(describeGuide(entries,parseXml).find(n=>n.index===at).children[0].cls,'g-picture-missing','with no way to find pictures it still says there is one');
  const icon=started&&await started('Icons');
  await addModule(icon,'M');await addLesson(icon,'L','5');await addBlock(icon,'ask','What?');
  const cue=describeGuide(icon.adapter.exportState().paragraphs,parseXml).find(n=>n.cls==='g-cue');
  assert.ok(cue.children.length>=2&&cue.children.every(c=>c.tag==='span'),'the cue\'s name, a space and its words: no picture of its icon');
  assert.equal(describeGuide([{xml:`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,text:'',style:'Normal'}],parseXml,{blank:true})[0].cls,'g-para g-blank','an empty paragraph is still blank');
});

test('a paragraph\'s own formatting is drawn over its style: alignment, the space around it, line spacing and indents; none when it has none',async()=>{
  const {withParagraphFormat}=await import('../src/guide-format.mjs');
  const io={parse:parseXml,serialize:el=>new XMLSerializer().serializeToString(el)};
  const a=await started('Formats');
  await a.adapter.insert(ctx=>para('OSGCueBody',run('Centered words')),{where:'start',spare:false});
  const p=a.adapter.exportState().paragraphs[0];
  const xml=withParagraphFormat(p.xml,p.style,{alignment:'center',spaceBefore:6,spaceAfter:12,lineSpacing:1.5,leftIndent:36,firstLineIndent:-18},io);
  const [node]=describeGuide([{...p,xml}],parseXml);
  assert.deepEqual(node.format,{alignment:'center',spaceBefore:6,spaceAfter:12,lineSpacing:1.5,leftIndent:36,firstLineIndent:-18});
  const doc=fakeDoc(),host=new Fake('div');
  drawNodes(doc,host,[node]);
  assert.deepEqual(host.children[0].style,{textAlign:'center',marginTop:'8px',marginBottom:'16px',lineHeight:'1.8',marginLeft:'48px',textIndent:'-24px'});
  assert.equal(describeGuide([p],parseXml)[0].format,undefined,'a paragraph that follows its style says nothing');
  const some=drawNodes(doc,new Fake('div'),[{...describeGuide([p],parseXml)[0]}]);
  void some;
});

test('words with their own look are drawn with it: weight, slant, color, raised, and never as markup',()=>{
  const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r=(t,pr='')=>`<w:r>${pr?`<w:rPr>${pr}</w:rPr>`:''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
  const xml=`<w:p xmlns:w="${W}">${r('Say ')}${r('this','<w:b/><w:i/><w:color w:val="C00000"/>')}${r('2','<w:vertAlign w:val="superscript"/>')}</w:p>`;
  const nodes=describeGuide([{xml,text:'Say this2',style:'Normal'}],parseXml);
  const spans=tags(nodes,'span');
  assert.deepEqual(spans.map(s=>s.text),['Say ','this','2']);
  assert.equal(spans[0].css,undefined,'plain words carry no look of their own');
  assert.equal(spans[1].css.fontWeight,'700');
  assert.equal(spans[1].css.fontStyle,'italic');
  assert.equal(spans[1].css.color,'#C00000');
  assert.equal(spans[2].css.verticalAlign,'super');
  // drawn with the page's own document, the look goes on as properties of the element
  const host=new Fake('div');
  drawNodes(fakeDoc(),host,nodes);
  const drawn=[];const walk=n=>{if(n.tag==='span')drawn.push(n);n.children.forEach(walk);};walk(host);
  assert.equal(drawn.find(n=>n.textContent==='this').style.fontWeight,'700');
  assert.equal(drawn.find(n=>n.textContent==='2').style.verticalAlign,'super');
});

test('words in a character style of the author\'s own are drawn with its look, under what was given to the words themselves',()=>{
  const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const spec={font:{bold:true,color:'#C00000'}};
  const edits={overrides:{},custom:[{id:'Warning',name:'Warning words',type:'character',basedOn:'DefaultParagraphFont',spec}]};
  const list=[{id:'DefaultParagraphFont',name:'Default Paragraph Font',type:'character',custom:false,own:{},effective:{}},{id:'Warning',name:'Warning words',type:'character',custom:true,basedOn:'DefaultParagraphFont',own:spec,effective:spec}];
  const styles=previewStyles(edits,list);
  assert.equal(styles.cssForId('Warning').fontWeight,'700');
  assert.equal(styles.cssForId('Missing'),null);
  const xml=`<w:p xmlns:w="${W}"><w:r><w:rPr><w:rStyle w:val="Warning"/></w:rPr><w:t>hot</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Warning"/><w:color w:val="112233"/></w:rPr><w:t>warm</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Gone"/></w:rPr><w:t>cold</w:t></w:r></w:p>`;
  const spans=tags(describeGuide([{xml,text:'hotwarmcold',style:'Normal'}],parseXml,{styles}),'span');
  assert.equal(spans[0].css.fontWeight,'700');assert.equal(spans[0].css.color,'#C00000');
  assert.equal(spans[1].css.color,'#112233','what was given to the words wins over the style');
  assert.equal(spans[1].css.fontWeight,'700');
  assert.equal(spans[2].css,undefined,'a style that is no longer there draws as plain words');
});
