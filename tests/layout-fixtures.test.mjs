import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {buildFixture,VARIANTS,LESSON,TINY_PNG} from '../scripts/layout-fixtures.mjs';
import {checkDocx,readOut,contrast} from '../scripts/docx-structure-check.mjs';

const icons=()=>TINY_PNG;
const mutate=async(bytes,file,fn)=>{
  const zip=await JSZip.loadAsync(bytes);
  zip.file(file,fn(await zip.file(file).async('string')));
  return zip.generateAsync({type:'nodebuffer'});
};
const rules=(result,level)=>result.findings.filter(f=>f.level===level).map(f=>f.rule);

test('each layout variant builds and has no structure failures',async()=>{
  for(const v of VARIANTS){
    const result=await checkDocx(await buildFixture(v,icons));
    assert.equal(result.counts.fail,0,`${v}: ${JSON.stringify(result.findings.filter(f=>f.level==='fail'))}`);
    assert.ok(result.findings.some(f=>f.rule==='TITLE'&&f.level==='pass'),v);
    assert.ok(result.findings.some(f=>f.rule==='LANGUAGE'&&f.level==='pass'),v);
    assert.ok(result.findings.some(f=>f.rule==='HEADINGS'&&f.level==='pass'),v);
    assert.deepEqual(rules(result,'warn'),[],`${v} should have no warnings`);
  }
});

test('only the comparison variant uses a table, and it has a three-column header row',async()=>{
  for(const v of VARIANTS){
    const result=await checkDocx(await buildFixture(v,icons));
    const tables=result.findings.filter(f=>f.rule==='TABLE');
    if(v==='table'){
      assert.equal(tables.length,1);
      assert.match(tables[0].message,/3 columns, 7 rows/);      assert.ok(result.findings.some(f=>f.rule==='TABLE-HEADER'&&f.level==='pass'));
    }else assert.equal(tables.length,0,`${v} must not use a layout table`);
  }
});

test('all variants carry the same lesson content so the comparison is fair',async()=>{
  for(const v of VARIANTS){
    const text=(await readOut(await buildFixture(v,icons))).join('\n');
    for(const b of LESSON.blocks){
      assert.ok(text.includes(b.text),`${v} is missing: ${b.text}`);
      if(b.notes)assert.ok(text.includes(b.notes),`${v} is missing notes: ${b.notes}`);
      if(b.prompt)assert.ok(text.includes(b.prompt),`${v} is missing prompt: ${b.prompt}`);
    }
  }
});

test('notes always start with a text label, never with shading alone',async()=>{
  for(const v of ['cue-labels','narrative']){
    const lines=await readOut(await buildFixture(v,icons));
    const notes=lines.filter(l=>/Facilitator Notes/i.test(l.split(':')[0]));
    assert.equal(notes.length,LESSON.blocks.filter(b=>b.notes).length);
    for(const l of notes)assert.match(l,/: Facilitator notes: /);
  }
});

test('decorative icons are skipped in the reading order',async()=>{
  for(const v of VARIANTS){
    const lines=await readOut(await buildFixture(v,icons));
    assert.ok(!lines.some(l=>l.includes('[picture')),v);
  }
});

test('the table variant reads row by row with its header row first',async()=>{
  const lines=await readOut(await buildFixture('table',icons));
  const at=lines.findIndex(l=>l.startsWith('Table with 3 columns and 7 rows.'));
  assert.ok(at>=0);
  assert.match(lines[at+1],/Row 1 \(header row\): Cue \| Content \| Notes/);
  assert.match(lines[at+4],/Ask \| What makes a disagreement hard to reframe\?/);
  assert.match(lines.join('\n'),/\| No notes\./);
});

test('a table without a header row fails',async()=>{
  const bytes=await mutate(await buildFixture('table',icons),'word/document.xml',x=>x.replace('<w:trPr><w:tblHeader/></w:trPr>',''));
  assert.ok(rules(await checkDocx(bytes),'fail').includes('TABLE-HEADER'));
});

test('merged and blank table cells are reported',async()=>{
  const bytes=await mutate(await buildFixture('table',icons),'word/document.xml',x=>x.replace('<w:tcW w:w="2500" w:type="dxa"/>','<w:tcW w:w="2500" w:type="dxa"/><w:gridSpan w:val="2"/>').replace('<w:t xml:space="preserve">No notes.</w:t>','<w:t xml:space="preserve"></w:t>'));
  const warn=rules(await checkDocx(bytes),'warn');
  assert.ok(warn.includes('TABLE-MERGED'));
  assert.ok(warn.includes('TABLE-BLANK'));
});

test('a picture with no alt text and no decorative mark fails',async()=>{
  const bytes=await mutate(await buildFixture('cue-labels',icons),'word/document.xml',x=>x.replace(/<a:extLst>.*?<\/a:extLst>/,''));
  assert.ok(rules(await checkDocx(bytes),'fail').includes('IMAGE-ALT'));
});

test('a skipped heading level fails',async()=>{
  let bytes=await buildFixture('narrative',icons);
  bytes=await mutate(bytes,'word/styles.xml',x=>x.replace('</w:styles>','<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style></w:styles>'));
  bytes=await mutate(bytes,'word/document.xml',x=>x.replace('<w:pStyle w:val="Heading2"/>','<w:pStyle w:val="Heading3"/>'));
  const result=await checkDocx(bytes);
  assert.match(result.findings.find(f=>f.rule==='HEADINGS'&&f.level==='fail')?.message||'',/Heading 3 follows heading 1/);
});

test('a missing title fails',async()=>{
  const bytes=await mutate(await buildFixture('narrative',icons),'docProps/core.xml',x=>x.replace(/<dc:title>[^<]*<\/dc:title>/,'<dc:title></dc:title>'));
  assert.ok(rules(await checkDocx(bytes),'fail').includes('TITLE'));
});

test('low contrast in a shaded style fails',async()=>{
  const bytes=await mutate(await buildFixture('cue-labels',icons),'word/styles.xml',x=>x.replace('<w:rPr><w:color w:val="1B2430"/></w:rPr>','<w:rPr><w:color w:val="C8C0AE"/></w:rPr>'));
  assert.ok(rules(await checkDocx(bytes),'fail').includes('CONTRAST'));
});

test('contrast ratio is computed correctly',()=>{
  assert.ok(Math.abs(contrast('000000','FFFFFF')-21)<0.01);
  assert.ok(Math.abs(contrast('FFFFFF','FFFFFF')-1)<0.01);
});

test('icons are half an inch square in every layout',async()=>{
  for(const v of VARIANTS){
    const xml=await (await JSZip.loadAsync(await buildFixture(v,icons))).file('word/document.xml').async('string');
    assert.equal((xml.match(/<wp:extent cx="457200" cy="457200"\/>/g)||[]).length,7,v);
    assert.ok(!xml.includes('cx="274320"'),v);
  }
});

test('an unknown layout name is refused',async()=>{
  await assert.rejects(buildFixture('grid',icons),/Unknown layout/);
});
