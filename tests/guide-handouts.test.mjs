// The handout pack: which forms go in it, in what order, with what page breaks and orientation, and what is left out.
import test from 'node:test';
import assert from 'node:assert/strict';
import {packRange,packEntries,withPageBreakBefore,PACK_SCOPES} from '../src/guide-handouts.mjs';
import {bodyWithContents,withSection} from '../src/guide-docx.mjs';
import {STYLE_NAMES} from '../src/guide-styles.mjs';

const P=(style,text,marks)=>({style,text,xml:`<w:p><w:pPr><w:pStyle w:val="X"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`,...(marks?{marks}:{})});
const title=(id,text,facilitator=false)=>P(facilitator?STYLE_NAMES.formTitleFacilitator:STYLE_NAMES.formTitle,text,{bookmarks:[`osg_form_${id}`],refs:[]});
const form=(id,text,print,facilitator=false)=>[title(id,text,facilitator),...(print?[P(STYLE_NAMES.formPrint,print)]:[]),P(STYLE_NAMES.formText,`Body of ${text}`)];
const cue=(text,refs=[])=>P(STYLE_NAMES.cue,text,{bookmarks:[],refs});
// Module 1 (lesson 1.1 with Goals, lesson 1.2 with Plans (landscape) and a facilitator-only key), then the Appendix with Cards.
const paras=[
  P(STYLE_NAMES.module,'Module 1: One'),
  P(STYLE_NAMES.lesson,'Lesson 1.1: First'),cue('Say Hello',[{id:'f9',text:'Handout 2.1: Cards'}]),
  ...form('f1','Handout 1.1: Goals','Print: 1 per person, Letter, portrait.'),
  P(STYLE_NAMES.lesson,'Lesson 1.2: Second'),
  ...form('f2','Handout 1.2: Plans','Print: 1 per person, Letter, landscape.'),
  ...form('f3','Worksheet 1.3: Key','Print: 1 copy.',true),
  P(STYLE_NAMES.module,'Module 2: Two'),
  P(STYLE_NAMES.lesson,'Lesson 2.1: Third'),cue('Say Bye'),
  ...form('f9','Handout 2.1: Cards','Print: 2 per team of 4.')
];
const entries=paras.map(({style,text,xml})=>({style,text,xml}));
const at=text=>paras.findIndex(p=>p.text===text);

test('the scopes are a lesson, a module and the whole guide, and each covers what it names',()=>{
  assert.deepEqual(PACK_SCOPES.map(s=>s.id),['lesson','module','guide']);
  const l=packRange(paras,at('Say Hello'),'lesson'),m=packRange(paras,at('Say Hello'),'module'),g=packRange(paras,0,'guide');
  assert.equal(l.name,'Lesson 1.1: First');assert.ok(l.start===at('Lesson 1.1: First'));assert.equal(l.end,at('Lesson 1.2: Second')-1,'the lesson runs to the next lesson, forms under its blocks included');
  assert.equal(m.name,'Module 1: One');assert.equal(m.end,at('Module 2: Two')-1);
  assert.deepEqual([g.start,g.end],[0,paras.length-1]);
  assert.throws(()=>packRange(paras,at('Module 1: One'),'lesson'),/Put your cursor in a lesson/);
  assert.throws(()=>packRange([P(STYLE_NAMES.cue,'Say Hello')],0,'module'),/Put your cursor in a module/);
  assert.throws(()=>packRange(paras,0,'nonsense'),/Choose a lesson/);
});

test('a lesson\'s pack holds the forms inside it and the forms its blocks cite, for everyone, without the print line',()=>{
  const pack=packEntries(entries,paras,{at:at('Say Hello'),scope:'lesson'});
  assert.deepEqual(pack.forms,['Handout 1.1: Goals','Handout 2.1: Cards'],'Goals is inside the lesson, Cards is cited from it');
  const texts=pack.entries.map(e=>e.text);
  assert.ok(texts.includes('Body of Handout 1.1: Goals'));
  assert.ok(!texts.some(t=>/^Print:/.test(t)),'the print line is for the facilitator');
  assert.equal(pack.scopeName,'Lesson 1.1: First');
});

test('a module\'s pack and the whole guide\'s pack leave out the form for the facilitator only',()=>{
  const m=packEntries(entries,paras,{at:at('Say Hello'),scope:'module'});
  assert.deepEqual(m.forms,['Handout 1.1: Goals','Handout 1.2: Plans','Handout 2.1: Cards']);
  const g=packEntries(entries,paras,{scope:'guide'});
  assert.deepEqual(g.forms,['Handout 1.1: Goals','Handout 1.2: Plans','Handout 2.1: Cards']);
  assert.ok(!g.entries.some(e=>/Key/.test(e.text)),'the form marked for the facilitator is not in it');
});

test('each form starts on a page of its own, except where a change of orientation starts a page already, and a landscape form is flagged',()=>{
  const g=packEntries(entries,paras,{scope:'guide'});
  const titles=g.entries.filter(e=>/^Handout/.test(e.text)&&/pStyle/.test(e.xml));
  assert.equal(titles.length,3);
  assert.ok(!/pageBreakBefore/.test(titles[0].xml),'the first form needs no break');
  assert.ok(!/pageBreakBefore/.test(titles[1].xml),'the landscape form is in a section that starts a page');
  assert.ok(!/pageBreakBefore/.test(titles[2].xml),'and the portrait form after it is in one too');
  assert.deepEqual(g.entries.map(e=>e.landscape),[false,false,true,true,false,false]);
  const body=withSection(bodyWithContents(g.entries));
  assert.equal((body.match(/w:orient="landscape"/g)||[]).length,1,'one landscape section');
  assert.equal((body.match(/<w:sectPr>/g)||[]).length,3);
});

test('two portrait forms in a row get a page break before the second only',()=>{
  const two=[...form('a','Handout 1: A','Print: 1 copy.'),...form('b','Handout 2: B','Print: 1 copy.')];
  const p2=two.map(({style,text,xml,marks})=>({style,text,xml,marks}));
  const pack=packEntries(p2.map(({style,text,xml})=>({style,text,xml})),p2,{scope:'guide'});
  const titles=pack.entries.filter(e=>/^Handout/.test(e.text));
  assert.ok(!/pageBreakBefore/.test(titles[0].xml));
  assert.match(titles[1].xml,/<w:pStyle w:val="X"\/><w:pageBreakBefore\/>/,'after the style, where a paragraph\'s page setup is written');
});

test('with no form for everyone it says what to do, and a page break goes where a paragraph\'s properties expect it',()=>{
  const none=[P(STYLE_NAMES.module,'Module 1: One'),P(STYLE_NAMES.lesson,'Lesson 1.1: First'),...form('f3','Worksheet 1: Key','Print: 1 copy.',true)];
  assert.throws(()=>packEntries(none.map(({style,text,xml})=>({style,text,xml})),none,{scope:'guide'}),/There is no form for everyone in the guide\. Start a form in the Forms tool/);
  assert.throws(()=>packEntries(entries,paras.slice(1),{scope:'guide'}),/changed while the pack was being made/);
  assert.equal(withPageBreakBefore('<w:p><w:r/></w:p>'),'<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r/></w:p>');
  assert.equal(withPageBreakBefore('<w:p><w:pPr><w:pStyle w:val="A"/><w:keepNext/><w:jc w:val="left"/></w:pPr></w:p>'),'<w:p><w:pPr><w:pStyle w:val="A"/><w:keepNext/><w:pageBreakBefore/><w:jc w:val="left"/></w:pPr></w:p>');
  assert.equal(withPageBreakBefore('<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'),'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>','not twice');
});
