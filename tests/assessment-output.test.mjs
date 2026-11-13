import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {prepareAssessmentExport,assessmentExportBlob,assessmentExportName} from '../src/assessment-output.mjs';
import {FAMILY_WORKBOOK_PARTS,FAMILY_WORKBOOK_SETTINGS_XML} from '../src/workbook-family-output.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');
const stamp='2026-10-28T14:00:00.000Z';
const options=(route='respondent')=>({route,instructions:'Read this\r\nThen answer\tcarefully & honestly \u4E2D\u6587 \u{1F642}',baseLanguage:'en-US',stamp});
function definition(count=5){return {schemaVersion:1,id:'public-assessment',version:2,title:'Synthetic measure & context',evidenceStatus:'not-established',items:[{id:'PRIVATE_ITEM_ID',type:'likert',prompt:'Rate the explanation\r\nwith a choice \u4E2D\u6587 \u{1F642}',reverse:true,options:Array.from({length:count},(_,i)=>({id:`PRIVATE_OPTION_${i}`,label:`Readable anchor ${i+1} ${'long '.repeat(20)}`,code:81001+i}))},{id:'PRIVATE_SEMANTIC_ID',type:'semantic-differential',prompt:'Rate clarity',reverse:false,target:'The synthetic target',leftEndpoint:'Unclear',rightEndpoint:'Clear',highScoreMeaning:'PRIVATE_HIGH_MEANING',options:Array.from({length:count},(_,i)=>({id:`PRIVATE_SEMANTIC_OPTION_${i}`,label:`Position label ${i+1}`,code:81001+i}))}],subscales:[{id:'PRIVATE_SUBSCALE_ID',name:'PRIVATE_SUBSCALE_NAME',itemIds:['PRIVATE_ITEM_ID','PRIVATE_SEMANTIC_ID'],aggregation:'mean',missing:{mode:'minimum-answered',minimumAnswered:1}}]};}

test('respondent projection excludes all private keys and both routes are definition-only immutable copies',()=>{
  const d=definition(),before=JSON.stringify(d),p=prepareAssessmentExport(d,options());
  assert.equal(JSON.stringify(d),before);assert.equal(Object.isFrozen(d),false);assert.equal(Object.isFrozen(p),true);assert.equal(Object.isFrozen(p.preview.items[0]),true);
  const whole=JSON.stringify(p);for(const secret of ['PRIVATE_','81001','81002','highScoreMeaning','subscales','optionId','reverse'])assert.ok(!whole.includes(secret),secret);
  assert.ok(whole.includes('The synthetic target'));assert.ok(whole.includes('Unclear'));assert.ok(whole.includes('Clear'));assert.ok(whole.includes('Read this\\nThen answer\\tcarefully'));
  d.items[0].prompt='MUTATED CALLER';assert.ok(!JSON.stringify(p).includes('MUTATED CALLER'));
  const a=prepareAssessmentExport(d,options('author'));assert.match(JSON.stringify(a),/PRIVATE author scoring guide/u);assert.match(JSON.stringify(a),/PRIVATE_SUBSCALE_NAME/u);assert.match(JSON.stringify(a),/81001/u);assert.match(JSON.stringify(a),/planned N/u);
  assert.throws(()=>prepareAssessmentExport(d,options(),{responses:[]}));assert.throws(()=>prepareAssessmentExport(d,{...options(),responses:[]}));
});

test('strict export options refuse hidden inherited getter symbol unsupported version and malformed text without getter reads',()=>{
  let hits=0;const good=options();
  const mutants=[null,undefined,{},Object.assign(Object.create(good),good),{...good,stamp:undefined},{...good,baseLanguage:'auto'},{...good,route:null},{...good,instructions:'bad\u0001'}, {...good,instructions:'bad\ud800'}];
  const getter={...good};Object.defineProperty(getter,'route',{enumerable:true,get(){hits++;return 'respondent';}});mutants.push(getter);
  const hidden={...good};Object.defineProperty(hidden,'extra',{value:1});mutants.push(hidden,{...good,[Symbol('hidden')]:true});
  for(const m of mutants)assert.throws(()=>prepareAssessmentExport(definition(),m));assert.equal(hits,0);
  for(const schemaVersion of [0,2,null])assert.throws(()=>prepareAssessmentExport({...definition(),schemaVersion},good));
});

test('two through eleven anchors retain ordered full labels and semantic real headings language compatibility in exact packages',async()=>{
  for(const count of [2,11])for(const route of ['respondent','author']){
    const d=definition(count),p=prepareAssessmentExport(d,options(route)),blob=await assessmentExportBlob(d,options(route),Zip),zip=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});
    assert.deepEqual(Object.keys(zip.files).sort(),[...FAMILY_WORKBOOK_PARTS].sort());assert.equal(await zip.file('word/settings.xml').async('string'),FAMILY_WORKBOOK_SETTINGS_XML);
    for(const name of FAMILY_WORKBOOK_PARTS){const raw=await zip.file(name).async('uint8array');assert.deepEqual(raw,new TextEncoder().encode(p.files[name]));let errors=[];const parser=new DOMParser({errorHandler:{warning:m=>errors.push(m),error:m=>errors.push(m),fatalError:m=>errors.push(m)}});assert.ok(parser.parseFromString(p.files[name],'application/xml').documentElement);assert.deepEqual(errors,[]);}
    const body=p.files['word/document.xml'];assert.match(body,/w:pStyle w:val="Heading1"/u);assert.match(body,/w:pStyle w:val="Heading2"/u);assert.match(p.files['word/styles.xml'],/w:lang w:val="en-US"/u);assert.match(p.files['docProps/core.xml'],/<dc:language>en-US/u);assert.match(body,/w:br/u);assert.match(body,/w:tab/u);assert.ok(body.includes('\u4E2D\u6587 \u{1F642}'));assert.ok(p.preview.instructions.includes('Read this\nThen answer\t'));assert.ok(!p.preview.instructions.includes('\r'));
    if(route==='respondent'){assert.equal((body.match(/\[ \] Position /gu)||[]).length,count*2);assert.equal((body.match(/\[ \] Not applicable/gu)||[]).length,2);assert.doesNotMatch(Object.values(p.files).join(''),/PRIVATE_|81001/u);}else{assert.match(body,/Reverse keyed: 81001/u);assert.match(body,/Minimum answered mean: k = 1/u);}
    assert.ok(!body.includes('<w:tbl'));assert.ok(!body.includes('<w:sdt'));assert.ok(!body.includes('<w:checkBox'));
  }
});

test('new route refuses altered returned ZIP parts extra parts and malformed UTF8 despite valid recomputed CRCs',async()=>{
  for(const route of ['respondent','author'])for(const mutate of [z=>z.file('word/document.xml','<changed>PRIVATE_RESPONSE_CANARY</changed>'),z=>z.file('word/styles.xml',new Uint8Array([0xc3,0x28])),z=>z.file('responses.json','PRIVATE_RESPONSE_CANARY')]){
    function AlteredZip(){const z=new Zip(),generate=z.generateAsync.bind(z);z.generateAsync=async config=>{const original=await generate(config),changed=await Zip.loadAsync(original);mutate(changed);return changed.generateAsync(config);};return z;}
    AlteredZip.loadAsync=Zip.loadAsync.bind(Zip);
    await assert.rejects(()=>assessmentExportBlob(definition(),options(route),AlteredZip),/changed|unsupported|ZIP|parts|inventory/u);
  }
});

test('private scoring guide states key direction missing NA partial and no aggregate policy without scoring responses',async()=>{
  const p=prepareAssessmentExport(definition(),options('author')),body=p.preview.paragraphs.map(x=>x.text).join('\n');
  for(const text of ['PRIVATE','L + U - x','exactly once','n/N','minimum','No overall score','Never treat missing','No test responses','no norms'])assert.ok(body.includes(text),text);
  const d=definition();d.subscales=[];assert.match(prepareAssessmentExport(d,options('author')).files['word/document.xml'],/No aggregate requested/u);
  await assert.rejects(()=>assessmentExportBlob(d,options(),Zip,{responses:[]}));
  assert.notEqual(assessmentExportName('respondent'),assessmentExportName('respondent'));assert.match(assessmentExportName('author'),/PRIVATE-Author/u);assert.throws(()=>assessmentExportName('../unsafe'));
});
