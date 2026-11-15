// The guide as one web page. The page is made from the tree the program draws the guide from (program/ui/preview.js), so these tests make the tree from a real
// guide, write the page, and read the page back as markup: it is well formed, it has the structure a screen reader needs, it holds its pictures inside it, and nothing
// in the tree can put markup, a script or an address into it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';
import {describeGuide} from '../program/ui/preview.js';
import {guideToHtml,nodesToHtml,WEB_CSS} from '../src/guide-html.mjs';
import {glanceBody} from '../src/guide-pages.mjs';
import {readGuide} from '../src/guide-model.mjs';
import {ICON_FILES,iconUrl} from '../src/guide-cues.mjs';

const parseXml=xml=>new DOMParser().parseFromString(xml,'application/xml');
// The page is written as markup that is also well formed, so it can be read back with an XML parser here.
const readPage=html=>{
  const d=parseXml(html.replace(/^<!doctype html>\s*/i,''));
  assert.equal(d.getElementsByTagName('parsererror').length,0,'the page is well formed');
  return d;
};
const byTag=(d,t)=>Array.from(d.getElementsByTagName(t));
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const DATA=`data:image/png;base64,${PNG}`;

async function guideTree(){
  const a=await started('Web guide');
  await addModule(a,'One');await addLesson(a,'First','10');await addBlock(a,'say','Welcome everyone.');
  await a.adapter.insert(()=>glanceBody(readGuide(a.adapter.snapshot()),{withHeading:true}),{where:'after',after:1,spare:false});
  const entries=a.adapter.exportState().paragraphs;
  return {a,nodes:describeGuide(entries,parseXml,{media:id=>a.adapter.mediaOf(id)})};
}

test('the page has what a screen reader needs: a language, a title, a way past the top, one main region, headings in order and a table with its header',async()=>{
  const {nodes}=await guideTree();
  const html=guideToHtml({nodes,title:'Web guide'});
  const d=readPage(html);
  assert.equal(d.documentElement.getAttribute('lang'),'en');
  assert.equal(byTag(d,'title')[0].textContent,'Web guide');
  assert.equal(byTag(d,'main').length,1);
  const skip=byTag(d,'a').find(a=>a.getAttribute('class')==='skip');
  assert.equal(skip.getAttribute('href'),'#main');assert.equal(byTag(d,'main')[0].getAttribute('id'),'main');
  assert.equal(byTag(d,'h1').length,1,'the title of the guide is the one top heading');
  assert.ok(byTag(d,'h2').some(h=>/Module 1: One/.test(h.textContent)));
  assert.ok(byTag(d,'h3').some(h=>/Lesson 1\.1: First/.test(h.textContent)));
  const table=byTag(d,'table')[0];
  assert.ok(table.getAttribute('aria-label'),'the table is named');
  assert.equal(byTag(d,'thead').length,1);assert.ok(byTag(d,'thead')[0].getElementsByTagName('th').length>=3);
  assert.ok(Array.from(byTag(d,'th')).every(th=>th.getAttribute('scope')==='col'));
  assert.equal(byTag(d,'tbody').length,1);
  assert.match(html,/<meta name="viewport"/);
});

test('every heading has an address, and a contents line is a link to the heading of the same words',()=>{
  const nodes=[
    {tag:'h1',cls:'g-title',text:'A guide'},
    {tag:'p',cls:'g-toc g-toc-1',text:'Module 1: One'},{tag:'p',cls:'g-toc g-toc-2',text:'Lesson 1.1: First'},{tag:'p',cls:'g-toc g-toc-1',text:'Not a heading'},
    {tag:'h2',cls:'g-module',text:'Module 1: One'},{tag:'h3',cls:'g-lesson',text:'Lesson 1.1: First'}
  ];
  const d=readPage(guideToHtml({nodes,title:'A guide'}));
  const ids=new Set(['h1','h2','h3'].flatMap(t=>byTag(d,t)).map(h=>h.getAttribute('id')));
  assert.equal(ids.size,3);assert.ok([...ids].every(id=>/^s\d+-[a-z0-9-]+$/.test(id)));
  const links=byTag(d,'a').filter(a=>a.getAttribute('class')!=='skip');
  assert.equal(links.length,2,'two of the three contents lines name a heading');
  assert.ok(links.every(a=>ids.has(a.getAttribute('href').slice(1))),'and each link goes to a heading that is there');
  assert.equal(byTag(d,'p').filter(p=>/Not a heading/.test(p.textContent))[0].getElementsByTagName('a').length,0);
});

test('the pictures are inside the page: the guide\'s own as data, the program\'s icons once they are given, and an icon that cannot be given is left out',async()=>{
  const icon=iconUrl(ICON_FILES[0]);
  const nodes=[
    {tag:'p',cls:'g-cue',children:[{tag:'img',cls:'g-icon',attrs:{src:icon,alt:'','aria-hidden':'true'}},{tag:'span',cls:'g-label',text:'Say'}]},
    {tag:'p',children:[{tag:'img',cls:'g-picture',attrs:{src:DATA,alt:'A group at a flip chart',width:'300'}}]},
    {tag:'p',children:[{tag:'img',attrs:{src:'/assets/icons-doc/missing.png',alt:''}}]}
  ];
  const withIcons=readPage(guideToHtml({nodes,title:'T',picture:src=>src===icon?DATA:null}));
  const imgs=byTag(withIcons,'img');
  assert.equal(imgs.length,2,'the icon and the picture; the third has no data to hold');
  assert.ok(imgs.every(i=>/^data:image\/png;base64,/.test(i.getAttribute('src'))));
  assert.equal(imgs[1].getAttribute('alt'),'A group at a flip chart');
  const without=readPage(guideToHtml({nodes,title:'T'}));
  assert.equal(byTag(without,'img').length,1,'with no icons given, only the guide\'s own picture is in the page');
});

test('nothing from the guide can become markup: tags, attributes, classes, styles and addresses are all checked',()=>{
  const nodes=[
    {tag:'script',text:'alert(1)'},
    {tag:'p',cls:'g-para" onclick="x',text:'<b>not bold</b> & "quoted"'},
    {tag:'p',attrs:{onclick:'steal()',style:'color:red','aria-label':'a "label" <x>'},text:'words'},
    {tag:'p',css:{color:'red',backgroundColor:'url(http://evil/x.png)',fontFamily:'Arial;}</style><script>'},text:'styled'},
    {tag:'p',css:{color:'#112233',fontSize:'12pt',position:'fixed'},text:'allowed'},
    {tag:'img',attrs:{src:'http://evil/x.png',alt:'x'}},
    {tag:'img',attrs:{src:'data:image/svg+xml;base64,PHN2Zz4=',alt:'x'}},
    {tag:'img',attrs:{src:'javascript:alert(1)',alt:'x'}}
  ];
  const html=guideToHtml({nodes,title:'T </title><script>alert(1)</script>'});
  const d=readPage(html);
  assert.equal(byTag(d,'script').length,0,'no script in the page');
  assert.equal(byTag(d,'img').length,0,'no picture that is not a picture of the kinds allowed');
  const markup=html.replace(/<title>[\s\S]*?<\/title>/,'').replace(/<style>[\s\S]*?<\/style>/,'');
  assert.ok(!/onclick|steal|evil|javascript|position:fixed|<b>|<script/.test(markup),'none of it is in the markup');
  assert.ok(markup.includes('<span>alert(1)</span>'),'a tag the view does not use is only words, in a span');
  assert.equal(byTag(d,'title')[0].textContent,'T </title><script>alert(1)</script>','the title is words, escaped');
  const ps=byTag(d,'p');
  assert.equal(ps[0].hasAttribute('class'),false,'a class with a quote in it is dropped');
  assert.equal(ps[0].textContent,'<b>not bold</b> & "quoted"');
  const styled=ps.find(p=>p.textContent==='styled');
  assert.equal(styled.getAttribute('style'),'color:red','only the safe property of the three is kept');
  assert.equal(ps.find(p=>p.textContent==='allowed').getAttribute('style'),'color:#112233;font-size:12pt');
  assert.throws(()=>guideToHtml({nodes:[],title:'T',lang:'en"><script>'}),/language of the page/);
});

test('the page can be read in the dark, in forced colors and on paper, and holds no address of anywhere else',async()=>{
  const {nodes}=await guideTree();
  const html=guideToHtml({nodes,title:'Web guide'});
  assert.match(WEB_CSS,/prefers-color-scheme:dark/);assert.match(WEB_CSS,/forced-colors:active/);assert.match(WEB_CSS,/@media print/);
  assert.ok(!/(?:src|href)="(?:https?:)?\/\//i.test(html),'no link or picture reaches anywhere else');
  assert.ok(!/@import|url\(/.test(WEB_CSS),'the style sheet loads nothing');
  assert.ok(html.length<80000,'one small file: '+html.length);
  assert.match(html,/<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">/,'and it tells the browser to load nothing but its own words and pictures, and to run nothing');
});

test('an empty tree gives a page that says so, and a table of only headings still has a body',()=>{
  assert.match(guideToHtml({nodes:[],title:'Empty'}),/This guide is empty/);
  const html=nodesToHtml([{tag:'table',cls:'g-table',children:[{tag:'tr',children:[{tag:'th',attrs:{scope:'col'},children:[{tag:'span',text:'Only'}]}]}]}]);
  assert.match(html,/<thead><tr><th scope="col"><span>Only<\/span><\/th><\/tr><\/thead><tbody><\/tbody>/);
});
