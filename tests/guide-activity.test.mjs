// An Activity written as labelled parts: objective, materials, preparation, instructions and case study. Proven against
// stand-ins for Word, not Word itself.
import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {makeContext,buildPackage,cueBlock} from '../src/guide-ooxml.mjs';
import {ACTIVITY_PARTS} from '../src/guide-cues.mjs';
import {readBlock,updatePlan} from '../src/guide-model.mjs';
import {summarize} from '../src/guide-workbook.mjs';
import {textOf} from '../src/guide-check.mjs';
import {parse,paragraphsOf} from './helpers/guide-package.mjs';
import {started,WORD_QUIRKS,click,submit,moveTo,addModule,addLesson,addBlock,startsWith} from './helpers/guide-harness.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const P=(text,style='Normal')=>({text,style});

const PARTS={
  objective:'Practice a neutral question.',
  materials:'Case cards\nTimer',
  preparation:'Print the cards.',
  instructions:'Pair up.\nSwap roles after four minutes.',
  caseStudy:'Sam says the report is late again.'
};
const ACT={cue:'activity',text:'Reframe in pairs',minutes:8,notes:'Walk the room.',parts:PARTS};
const ROWS=[
  'OSG Cue: Activity Reframe in pairs (8 min)',
  'OSG Cue Body: Objective: Practice a neutral question.',
  // a part of several lines is its label on a line of its own and a bulleted list under it (author, 2026-11-07)
  'OSG Facilitator Notes: Materials:',
  'OSG Facilitator Notes: Case cards',
  'OSG Facilitator Notes: Timer',
  'OSG Facilitator Notes: Preparation: Print the cards.',
  'OSG Cue Body: Instructions:',
  'OSG Cue Body: Pair up.',
  'OSG Cue Body: Swap roles after four minutes.',
  'OSG Cue Body: Case study: Sam says the report is late again.',
  'OSG Facilitator Notes: Facilitator notes: Walk the room.'
];
const rowsOf=pkg=>paragraphsOf(pkg).map(p=>`${p.style}: ${p.text.trim()}`.trim());
const make=block=>{const ctx=makeContext(),xml=cueBlock(block,ctx);return buildPackage(xml,ctx,Object.fromEntries([...ctx.used].map(i=>[i,'AAAA'])));};

async function lesson(){
  const a=await started('Reframing conflict',{quirks:WORD_QUIRKS});
  await addModule(a,'Recognize');
  await addLesson(a,'Listen','20');
  return a;
}
// Fills the form the way the author would: pick Activity, tick the box, type the parts, press Insert.
async function typeParts(a,{title='Reframe in pairs',parts=PARTS,notes='Walk the room.',minutes='8'}={}){
  a.pick('activity');
  a.n('guide-activity-parts').checked=true;a.n('guide-activity-parts').onchange();
  a.type('guide-block-text',title);
  for(const p of ACTIVITY_PARTS)a.type(`guide-act-${p.id}`,parts[p.id]||'');
  a.type('guide-block-notes',notes);a.type('guide-block-minutes',minutes);
  await submit(a,'guide-block-form');
}

// ---- the page content
test('an activity written as parts is a cue line with the title and minutes, then each part in its own labelled paragraph',()=>{
  assert.deepEqual(rowsOf(make(ACT)),ROWS);
});

test('materials and preparation are in the facilitator-notes style, the rest in the plain block style, each with its label in bold',()=>{
  const ctx=makeContext(),xml=cueBlock(ACT,ctx);
  const doc=parse(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`);
  const ps=Array.from(doc.getElementsByTagNameNS(W,'p')).slice(1);
  const style=p=>Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val');
  assert.deepEqual(ps.map(style),['OSGCueBody','OSGNotes','OSGNotes','OSGNotes','OSGNotes','OSGCueBody','OSGCueBody','OSGCueBody','OSGCueBody','OSGNotes']);
  const labelled=ps.filter(p=>Array.from(p.getElementsByTagNameNS(W,'rStyle')).some(r=>r.getAttributeNS(W,'val')==='OSGCueLabel')).map(p=>textOf(p).split(':')[0]);
  assert.deepEqual(labelled,['Objective','Materials','Preparation','Instructions','Case study','Facilitator notes']);
});

test('the steps and items of an activity are bulleted lists; one line, the objective and the case study are not',()=>{
  const ctx=makeContext(),xml=cueBlock(ACT,ctx);
  const doc=parse(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`);
  const ps=Array.from(doc.getElementsByTagNameNS(W,'p'));
  const bulleted=ps.filter(p=>p.getElementsByTagNameNS(W,'numPr').length).map(p=>textOf(p));
  assert.deepEqual(bulleted,['Case cards','Timer','Pair up.','Swap roles after four minutes.']);
  assert.equal(ctx.lists,2,'one list for the materials, one for the steps');
  assert.deepEqual([...ctx.bullets],[1,2]);
  const pkg=make(ACT);
  assert.match(pkg,/<w:numFmt w:val="bullet"\/>/,'the package carries a bullet list definition');
  assert.match(pkg,/<w:numId w:val="1"\/>/);
  // the bullets stand under the text of the label line (its indent, plus the bullet's own hang), not out to the left of the block
  assert.match(pkg,/<w:numFmt w:val="bullet"\/>.*?<w:ind w:left="2760" w:hanging="360"\/>/);
  // a label with its list stays with it on the page
  assert.ok(ps.some(p=>textOf(p)==='Instructions:'&&p.getElementsByTagNameNS(W,'keepNext').length===1));
  // one line stays on its label's line, with no bullet
  const one=cueBlock({cue:'activity',text:'Quick',parts:{instructions:'Just one step.'}},makeContext());
  assert.ok(!one.includes('<w:numPr>'));
  // bulleted and lettered lists can share a document: the bullets use their own definition
  const both=makeContext();both.lists=1;cueBlock({cue:'activity',text:'T',parts:{instructions:'A\nB'}},both);
  assert.deepEqual([...both.bullets],[2]);
});

test('parts left empty are left out, and only the instructions are required',()=>{
  const rows=rowsOf(make({cue:'activity',text:'Quick check',parts:{instructions:'Turn to a neighbour and compare answers.'}}));
  assert.deepEqual(rows,['OSG Cue: Activity Quick check','OSG Cue Body: Instructions: Turn to a neighbour and compare answers.']);
  assert.throws(()=>make({cue:'activity',text:'No steps',parts:{objective:'x'}}),/Type the instructions for the activity/);
  assert.throws(()=>make({cue:'activity',text:'One\nTwo',parts:{instructions:'x'}}),/title of an activity written as labelled parts is one line/);
  assert.throws(()=>make({cue:'say',text:'Hello',parts:{instructions:'x'}}),/Only an Activity/);
});

// ---- reading it back
test('an activity written as parts reads back as its title, parts and own notes; any other activity reads as plain text',()=>{
  const paras=[P('Module 1: A','OSG Module'),P('Lesson 1.1: B','OSG Lesson'),...rowsOf(make(ACT)).map(r=>{const i=r.indexOf(': ');return P(r.slice(i+2),r.slice(0,i));})];
  const b=readBlock(paras,2);
  assert.deepEqual(b.parts,PARTS);
  assert.deepEqual([b.text,b.notes,b.minutes,b.audience],['Reframe in pairs','Walk the room.','8','everyone']);
  const plain=readBlock([P('Module 1: A','OSG Module'),P('Activity   Talk it over (5 min)','OSG Cue'),P('A second line.','OSG Cue Body'),P('Facilitator notes: Listen in.','OSG Facilitator Notes')],1);
  assert.equal(plain.parts,undefined);
  assert.deepEqual([plain.text,plain.notes],['Talk it over\nA second line.','Listen in.']);
  // Body text that is not a labelled part makes it a plain block, so nothing is lost on edit.
  const mixed=readBlock([P('Activity   Mixed','OSG Cue'),P('Instructions: Do it.','OSG Cue Body'),P('Some other line that is not a part.','OSG Facilitator Notes')],0);
  assert.equal(mixed.parts,undefined);
});

// ---- the form
test('the parts fields show only for an Activity with the box ticked, and the Text box becomes its title',async()=>{
  const a=await lesson();
  assert.equal(a.n('guide-activity-box').hidden,true);
  a.pick('activity');
  assert.equal(a.n('guide-activity-box').hidden,false);
  assert.equal(a.n('guide-activity-fields').hidden,true);
  assert.equal(a.n('guide-block-text-label').textContent,'Text');
  a.n('guide-activity-parts').checked=true;a.n('guide-activity-parts').onchange();
  assert.equal(a.n('guide-activity-fields').hidden,false);
  assert.equal(a.n('guide-block-text-label').textContent,'Title of the activity');
  a.pick('say');
  assert.equal(a.n('guide-activity-box').hidden,true);
  assert.equal(a.n('guide-activity-fields').hidden,true);
  assert.equal(a.n('guide-activity-parts').checked,false,'another cue puts the form back to plain text');
  assert.equal(a.n('guide-block-text-label').textContent,'Text');
});

test('Insert block writes the parts, with the minutes on the title line, and clears the part fields',async()=>{
  const a=await lesson();
  await typeParts(a);
  assert.equal(a.err(),'');
  const at=a.rows().findIndex(r=>r.startsWith('OSG Cue: Activity'));
  assert.deepEqual(a.rows().slice(at,at+ROWS.length),ROWS);
  assert.match(a.n('guide-status').textContent,/Activity block added/);
  for(const p of ACTIVITY_PARTS)assert.equal(a.n(`guide-act-${p.id}`).value,'');
  assert.equal(a.n('guide-block-text').value,'');
});

test('the form says what is missing: a title, then the instructions',async()=>{
  const a=await lesson();
  const before=a.rows();
  await typeParts(a,{title:'',parts:PARTS});
  assert.match(a.err(),/Type a title for the activity first/);
  await typeParts(a,{parts:{objective:'Only an objective.'}});
  assert.match(a.err(),/Type the instructions for the activity/);
  assert.deepEqual(a.rows(),before);
});

test('the activity\'s minutes feed the lesson\'s time like any block\'s',async()=>{
  const a=await lesson();
  await typeParts(a,{minutes:'8'});
  const plan=updatePlan(await a.adapter.paragraphs());
  assert.equal(plan.totalMinutes,8);
  assert.equal(plan.guide.modules[0].lessons[0].blockMinutes,8);
});

test('Edit this block opens the parts filled in, Update block replaces the whole activity once, and what it wrote stays in order',async()=>{
  const a=await lesson();
  await typeParts(a);
  await moveTo(a,startsWith('OSG Cue: Activity'));
  assert.equal(a.n('guide-edit-selected').textContent,'Edit this block');
  await click(a,'guide-edit-selected');
  assert.equal(a.n('guide-activity-parts').checked,true);
  assert.equal(a.n('guide-activity-fields').hidden,false);
  assert.equal(a.n('guide-block-text').value,'Reframe in pairs');
  for(const p of ACTIVITY_PARTS)assert.equal(a.n(`guide-act-${p.id}`).value,PARTS[p.id],p.id);
  assert.equal(a.n('guide-block-notes').value,'Walk the room.');
  assert.equal(a.n('guide-block-minutes').value,'8');
  a.type('guide-act-instructions','Pair up.\nSwap roles after three minutes.\nDebrief together.');
  a.type('guide-act-materials','');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  const rows=a.rows(),at=rows.findIndex(r=>r.startsWith('OSG Cue: Activity'));
  assert.equal(rows.filter(r=>r.startsWith('OSG Cue: Activity')).length,1);
  assert.deepEqual(rows.slice(at,at+9),[
    'OSG Cue: Activity Reframe in pairs (8 min)',
    'OSG Cue Body: Objective: Practice a neutral question.',
    'OSG Facilitator Notes: Preparation: Print the cards.',
    'OSG Cue Body: Instructions:',
    'OSG Cue Body: Pair up.',
    'OSG Cue Body: Swap roles after three minutes.',
    'OSG Cue Body: Debrief together.',
    'OSG Cue Body: Case study: Sam says the report is late again.',
    'OSG Facilitator Notes: Facilitator notes: Walk the room.'
  ]);
  assert.match(a.n('guide-status').textContent,/Activity block updated/);
});

test('an activity written as plain text still opens as plain text, with the box unticked',async()=>{
  const a=await lesson();
  await addBlock(a,'activity','Talk it over.\nPick one idea to share.',{notes:'Allow five minutes.',minutes:'5'});
  await moveTo(a,startsWith('OSG Cue: Activity'));
  await click(a,'guide-edit-selected');
  assert.equal(a.n('guide-activity-parts').checked,false);
  assert.equal(a.n('guide-activity-fields').hidden,true);
  assert.equal(a.n('guide-block-text').value,'Talk it over.\nPick one idea to share.');
});

// ---- the learner workbook and the review
test('the workbook keeps the objective, instructions and case study, and never the materials, preparation or notes',async()=>{
  const a=await lesson();
  await typeParts(a);
  await click(a,'guide-open-review');
  assert.deepEqual(a.n('guide-review-fix').children.map(li=>li.textContent),['Nothing to fix.']);
  const s=summarize(await a.adapter.paragraphs());
  assert.equal(s.blocks,1);assert.deepEqual(s.leftOut,[]);
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  assert.equal(a.n('guide-workbook-problem').hidden,true,'the leak check passes');
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const text=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p));
  const all=text.join('\n');
  for(const kept of ['Objective: Practice a neutral question.','Instructions:','Pair up.','Swap roles after four minutes.','Case study: Sam says the report is late again.'])assert.ok(all.includes(kept),`${kept} is in the workbook`);
  for(const gone of ['Materials','Case cards','Timer','Preparation','Print the cards','Walk the room','(8 min)'])assert.ok(!all.includes(gone),`${gone} is not in the workbook`);
  assert.ok(text.some(t=>t.includes('Reframe in pairs')),'the title is');
});

test('a facilitator-only activity written as parts is left out of the workbook whole',async()=>{
  const a=await lesson();
  await addBlock(a,'ask','What is hard about it?');
  a.pick('activity');
  a.n('guide-activity-parts').checked=true;a.n('guide-activity-parts').onchange();
  a.n('guide-audience-facilitator').checked=true;a.n('guide-audience-everyone').checked=false;
  a.type('guide-block-text','Demonstration');a.type('guide-act-instructions','Show the two versions side by side.');
  await submit(a,'guide-block-form');
  assert.equal(a.err(),'');
  await click(a,'guide-tab-workbook');await click(a,'guide-workbook-go');
  const opened=a.adapter.openedDocuments()[0];
  assert.ok(opened,a.n('guide-workbook-problem').textContent);
  const zip=await JSZip.loadAsync(Buffer.from(opened.base64,'base64')),doc=parse(await zip.file('word/document.xml').async('string'));
  const all=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>textOf(p)).join('\n');
  assert.ok(!all.includes('Demonstration')&&!all.includes('side by side'));
  assert.ok(all.includes('What is hard about it?'));
});
