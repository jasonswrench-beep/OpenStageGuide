import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {CUES,AUDIENCES,ICON_FILES,cueById,iconUrl,TIME_ICON} from '../src/guide-cues.mjs';
import {STYLE_NAMES,STYLE_IDS} from '../src/guide-styles.mjs';
import {makeContext,moduleHeading,lessonHeading,timeBlock,cueBlock,startGuide,sampleGuide,buildPackage,timeText} from '../src/guide-ooxml.mjs';
import {readStyles} from '../src/guide-check.mjs';
import {checkDocx} from '../scripts/docx-structure-check.mjs';
import {TINY,parse,partXml,paragraphsOf,packageToDocx} from './helpers/guide-package.mjs';

const iconData=Object.fromEntries(ICON_FILES.map(i=>[i,TINY]));
test('every cue has a unique id and label, a known audience and an icon file that exists',async()=>{
  assert.equal(new Set(CUES.map(c=>c.id)).size,CUES.length);
  assert.equal(new Set(CUES.map(c=>c.label)).size,CUES.length);
  for(const c of CUES){
    assert.ok(AUDIENCES.some(a=>a.id===c.audience),c.id);
    assert.ok(c.help.length>5,c.id);
    const bytes=await fs.readFile(new URL(`../public${iconUrl(c.icon)}`,import.meta.url));
    assert.equal(bytes.subarray(1,4).toString(),'PNG',`${c.icon} is a PNG`);
  }
  assert.ok(ICON_FILES.includes(TIME_ICON));
});

test('the fourteen original icons and the forty-eight new ones are all offered, one picture for each cue',()=>{
  assert.deepEqual([...new Set(CUES.map(c=>c.icon))].sort(),['accessibility-note','animation-note','assessment','audio','branch','break','breakout','capture','case-study','chat','check-computer','click-interaction','data-tables','developer-note','discussion','do-this','evaluation','feedback','files','flipchart','game','hand-out','highlight','important','interaction','key-point','knowledge-check','lab','lecture','media','media-asset','microphone','narration','objective','on-screen-text','picture','polling','powerpoint','presenter-1','presenter-2','process','producer','qa','questions','raise-hand','role-play','say','screen-title','slide-notes','steps','time-on-task','tools','trainer-notes','training-instrument','transitions','video','visual-direction','vote','webcam','website','workbook']);
  assert.equal(cueById('say').icon,'say');
  assert.equal(cueById('ask').icon,'questions');
  assert.equal(new Set(CUES.map(c=>c.icon)).size,CUES.length,'no two cues share a picture');
});

test('a cue block has a decorative icon, a label, the text, and notes that start with a text label',()=>{
  const ctx=makeContext();
  const xml=cueBlock({cue:'ask',text:'What makes it hard?\nWhy?',notes:'Listen for blame.\nWrite answers down.',minutes:5},ctx);
  assert.match(xml,/adec:decorative[^>]*val="1"/);
  assert.match(xml,new RegExp(`<w:pStyle w:val="${STYLE_IDS.cue}"/>`));
  assert.match(xml,/<w:t xml:space="preserve"> Ask<\/w:t>/);
  assert.match(xml,/What makes it hard\? \(5 min\)/);
  assert.match(xml,new RegExp(`<w:pStyle w:val="${STYLE_IDS.body}"/>`));
  assert.match(xml,/Facilitator notes: /);
  assert.equal((xml.match(new RegExp(`<w:pStyle w:val="${STYLE_IDS.notes}"/>`,'g'))||[]).length,2);
  assert.deepEqual([...ctx.used],['questions']);
});

test('facilitator-only blocks use the facilitator style, and the author can override a cue\'s default',()=>{
  assert.match(cueBlock({cue:'say',text:'x'},makeContext()),new RegExp(STYLE_IDS.cueFacilitator));
  assert.match(cueBlock({cue:'say',text:'x',audience:'everyone'},makeContext()),new RegExp(`<w:pStyle w:val="${STYLE_IDS.cue}"/>`));
  assert.match(cueBlock({cue:'ask',text:'x',audience:'facilitator'},makeContext()),new RegExp(STYLE_IDS.cueFacilitator));
});

test('text is escaped and an unknown cue or empty title is refused',()=>{
  assert.match(cueBlock({cue:'ask',text:'A & B <c> "d"'},makeContext()),/A &amp; B &lt;c&gt; &quot;d&quot;/);
  assert.throws(()=>cueBlock({cue:'nope',text:'x'},makeContext()),/Choose a cue/);
  assert.throws(()=>startGuide({title:'   '},makeContext()),/Type a title/);
});

test('time wording matches what the author approved, with singular minutes',()=>{
  assert.equal(timeText(20,0),'Time: 20 minutes. Cumulative time: 0 minutes.');
  assert.equal(timeText(1,1),'Time: 1 minute. Cumulative time: 1 minute.');
});

test('a package needs every icon it uses',()=>{
  const ctx=makeContext(),body=cueBlock({cue:'ask',text:'x'},ctx);
  assert.throws(()=>buildPackage(body,ctx,{}),/picture for the questions cue/);
});

test('the sample guide package is well-formed XML and passes the structure check',async()=>{
  const ctx=makeContext(),pkg=buildPackage(sampleGuide(ctx),ctx,iconData);
  parse(pkg);
  const result=await checkDocx(await packageToDocx(pkg,'Sample guide'));
  assert.equal(result.counts.fail,0,JSON.stringify(result.findings.filter(f=>f.level==='fail')));
  assert.deepEqual(result.findings.filter(f=>f.level==='warn').map(f=>f.rule),[]);
  assert.ok(result.findings.some(f=>f.rule==='IMAGE-ALT'&&f.level==='pass'));
});

test('start guide inserts the standard pages with real headings and a contents field',async()=>{
  const ctx=makeContext(),pkg=buildPackage(startGuide({title:'How to Reframe Conflict',standard:true},ctx),ctx,iconData);
  const names=paragraphsOf(pkg).map(p=>p.text.trim()).filter(Boolean);
  assert.equal(names[0],'How to Reframe Conflict');
  for(const h of ['Contents','Getting Started','About This Guide','Program Preparation','Training at a Glance'])assert.ok(names.includes(h),h);
  assert.match(pkg,/ TOC \\o "1-2" \\h \\z \\u /);
  const result=await checkDocx(await packageToDocx(pkg,'How to Reframe Conflict'));
  assert.equal(result.counts.fail,0);
});

test('a guide without the standard pages is just the title',()=>{
  const ctx=makeContext(),pkg=buildPackage(startGuide({title:'Tiny',standard:false},ctx),ctx,iconData);
  assert.deepEqual(paragraphsOf(pkg).map(p=>p.text.trim()),['Tiny']);
});

test('module and lesson headings use their own styles and the agreed wording',()=>{
  const ctx=makeContext(),body=moduleHeading(2,'Reframe')+lessonHeading(2,3,'Neutral questions')+timeBlock({minutes:10,cumulative:30},ctx);
  const rows=paragraphsOf(buildPackage(body,ctx,iconData));
  assert.deepEqual(rows.map(r=>r.style),[STYLE_NAMES.module,STYLE_NAMES.lesson,STYLE_NAMES.time]);
  assert.equal(rows[0].text,'Module 2: Reframe');
  assert.equal(rows[1].text,'Lesson 2.3: Neutral questions');
  assert.match(rows[2].text,/Time: 10 minutes\. Cumulative time: 30 minutes\./);
});

test('module and lesson styles are based on Heading 1 and 2 so Word treats them as headings',()=>{
  const ctx=makeContext(),pkg=buildPackage(moduleHeading(1,'x'),ctx,iconData);
  const styles=readStyles(parse(partXml(pkg,'/word/styles.xml')));
  assert.equal(styles.get(STYLE_IDS.module).basedOn,'Heading1');
  assert.equal(styles.get(STYLE_IDS.lesson).basedOn,'Heading2');
  assert.equal(styles.get('Heading1').outline,1);
});
