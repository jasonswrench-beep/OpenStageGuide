import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {WORKBOOK_FAMILIES} from '../src/workbook-presentation.mjs';
import {prepareFamilyWorkbook,familyReviewText,FAMILY_BRAND_FIELDS,familyWorkbookBlob} from '../src/workbook-family-flow.mjs';
import {rawControl} from '../src/activity-draft.mjs';
import {EMPTY_WORKBOOK_FRONT_MATTER,WORKBOOK_COVER_FIELDS,WORKBOOK_COPYRIGHT_FIELDS,validateWorkbookFrontMatter,workbookFrontMatterRequest,workbookCopyrightNoticeStarter,workbookFrontMatterRows} from '../src/workbook-public-frontmatter.mjs';
const require=createRequire(import.meta.url),Zip=require('jszip'),{DOMParser}=require('@xmldom/xmldom');
const stamp='2026-10-28T14:30:00.000Z';
const metadata=()=>({schemaVersion:1,includeCover:true,includeCopyright:true,cover:{title:'Public cover < & >',subtitle:'Subtitle\nSecond line\tcontinued \u4e2d\u6587 \u{1f642}',author:'Explicit Author',edition:'Second edition',year:'2024\u20132026',audience:'Public learners'},copyright:{rightsHolder:'Explicit Rights Holder',year:'2024-2026',noticeChoice:'custom-text',noticeText:'Exact custom notice < & >\nSecond line\tcontinued.',permissionsContact:'Ask the authored contact',credits:'Exact credits',licenseName:'INACTIVE_PRIVATE_LICENSE',licenseVersion:'INACTIVE_PRIVATE_VERSION',licenseUrl:'INACTIVE_PRIVATE_URL'}});
const brand=()=>Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,rawControl({value:k==='organizationName'?'Public Organization':''})]));
const raw=m=>({schemaVersion:1,includeCover:{checked:m.includeCover,...rawControl({value:'on'})},includeCopyright:{checked:m.includeCopyright,...rawControl({value:'on'})},cover:Object.fromEntries(WORKBOOK_COVER_FIELDS.map(k=>[k,rawControl({value:m.cover[k]})])),copyright:Object.fromEntries(WORKBOOK_COPYRIGHT_FIELDS.map(k=>[k,rawControl({value:m.copyright[k]})]))});
function fixture(wrench=false){
 let p=wrench?wrenchCourse('Course learner title',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_RIGHTS'},['background']):course('Course learner title');
 p=addModule(p,'Selected module','Objective text').next;
 p=addActivity(p,p.modules[0].id,{title:'Omitted',minutes:5,participant:'OMITTED_LEARNER',facilitator:'PRIVATE_TEACHER',answers:'PRIVATE_ANSWERS',lines:0}).next;
 p=addActivity(p,p.modules[0].id,{title:'Selected activity',minutes:8,participant:'Public learner instructions',facilitator:'PRIVATE_TEACHER',answers:'PRIVATE_ANSWERS',lines:3}).next;
 const result=editDeliveryPlan(p,{kind:'create',name:'Public plan',targetMinutes:8,groups:[{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}]});return {p:result.next,id:result.id};
}
const config=(familyId,planId='',frontMatter)=>({planId,familyId,rawBrand:brand(),stamp,...(frontMatter===undefined?{}:{frontMatter})});

test('public metadata is explicit defensive versioned data; inactive text stays raw but not published',()=>{
 const m=metadata(),before=JSON.stringify(m),canonical=validateWorkbookFrontMatter(m);
 assert.equal(JSON.stringify(m),before);assert.equal(Object.isFrozen(m),false);assert.equal(Object.isFrozen(canonical.cover),true);
 assert.equal(validateWorkbookFrontMatter({...metadata(),cover:{...metadata().cover,subtitle:String.raw`Literal \r\n \t text`}}).cover.subtitle,String.raw`Literal \r\n \t text`);
 assert.equal(canonical.cover.subtitle,m.cover.subtitle);m.cover.title='CHANGED';assert.equal(canonical.cover.title,'Public cover < & >');
 const r=workbookFrontMatterRows(canonical,'Public Organization');assert.equal(r.schemaVersion,1);assert.equal(r.cover[0].text,'Public Organization');assert.equal(r.cover.at(-1).stopKeep,true);assert.equal(r.copyright[0].pageBreak,true);assert.doesNotMatch(JSON.stringify(r),/INACTIVE_PRIVATE/);
 const off=workbookFrontMatterRows({...metadata(),includeCover:false,includeCopyright:false},'Public Organization');assert.deepEqual(off.cover,[]);assert.deepEqual(off.copyright,[]);
});

test('strict own data null version type hidden symbol inherited getter and extra records refuse without reads',()=>{
 const good=metadata();let hits=0;
 const getter={...good};Object.defineProperty(getter,'cover',{enumerable:true,get(){hits++;return good.cover;}});
 const hidden={...good};Object.defineProperty(hidden,'includeCover',{value:true,enumerable:false});
 for(const bad of [null,undefined,{},getter,hidden,{...good,schemaVersion:2},{...good,includeCover:'yes'},{...good,unknown:'x'},{...good,[Symbol('x')]:1},Object.assign(Object.create({x:1}),good),{...good,cover:{...good.cover,title:null}}])assert.throws(()=>validateWorkbookFrontMatter(bad));
 assert.equal(hits,0);assert.throws(()=>validateWorkbookFrontMatter(good,{}));assert.throws(()=>workbookFrontMatterRows(good,'Org',{}));
 const r=raw(good);Object.defineProperty(r.copyright.noticeChoice,'value',{enumerable:true,get(){hits++;return 'custom-text';}});assert.throws(()=>workbookFrontMatterRequest(r));assert.equal(hits,0);
 const checkbox=raw(good);checkbox.includeCover.checked='true';assert.throws(()=>workbookFrontMatterRequest(checkbox));const invalid=raw(good);invalid.includeCover.invalid.customError=true;assert.throws(()=>workbookFrontMatterRequest(invalid));
 const starterBad={...good,includeCover:'yes'};assert.throws(()=>workbookCopyrightNoticeStarter(starterBad));
});

test('required authored fields year ranges and explicit named license validate with no inferred rights',()=>{
 for(const year of ['2026','2024-2026',' 2024 \u2013 2026 '])assert.doesNotThrow(()=>validateWorkbookFrontMatter({...metadata(),copyright:{...metadata().copyright,year}}));
 for(const year of ['2026-2024','0000','2026/2027','202','today',' '])assert.throws(()=>validateWorkbookFrontMatter({...metadata(),copyright:{...metadata().copyright,year}}));
 for(const [part,key]of [['cover','title'],['copyright','rightsHolder'],['copyright','noticeText'],['copyright','noticeChoice']])assert.throws(()=>validateWorkbookFrontMatter({...metadata(),[part]:{...metadata()[part],[key]:' '}}));
 const named={...metadata(),copyright:{...metadata().copyright,noticeChoice:'named-license',licenseName:'An explicitly authored license',licenseVersion:'Version 2',licenseUrl:'https://example.test/license'}};
 assert.doesNotThrow(()=>validateWorkbookFrontMatter(named));assert.match(workbookCopyrightNoticeStarter(named),/An explicitly authored license Version 2/u);
 for(const licenseUrl of ['javascript:alert(1)','file:///x','https://user:secret@example.test/license',' https://example.test'])assert.throws(()=>validateWorkbookFrontMatter({...named,copyright:{...named.copyright,licenseUrl}}));
 const noChoice={...EMPTY_WORKBOOK_FRONT_MATTER};assert.throws(()=>workbookCopyrightNoticeStarter(noChoice));
 const permissions={...metadata(),copyright:{...metadata().copyright,noticeChoice:'custom-permissions'}};const starter=workbookCopyrightNoticeStarter(permissions);assert.match(starter,/Specify the uses/u);assert.doesNotMatch(starter,/Permission is granted/u);
});

test('inactive fields preserve invalid buffers and do not publish or enforce active-only rules',()=>{
 const m=metadata();m.includeCover=false;m.cover.year='NOT A YEAR';m.copyright.licenseUrl='NOT A URL';const r=raw(m);r.cover.year.invalid.patternMismatch=true;r.copyright.licenseUrl.invalid.customError=true;
 const before=JSON.stringify(r),v=workbookFrontMatterRequest(r);assert.equal(JSON.stringify(r),before);assert.equal(v.cover.year,'NOT A YEAR');assert.equal(v.copyright.licenseUrl,'NOT A URL');
 r.includeCover.checked=true;assert.throws(()=>workbookFrontMatterRequest(r));r.includeCover.checked=false;r.copyright.noticeChoice.value='named-license';assert.throws(()=>workbookFrontMatterRequest(r));
 r.copyright.noticeChoice.value='custom-text';assert.doesNotThrow(()=>workbookFrontMatterRequest(r));
 const white=metadata();for(const k of ['subtitle','author','edition','audience'])white.cover[k]='  ';white.cover.year='';white.copyright.permissionsContact='  ';white.copyright.credits='\n';const rows=workbookFrontMatterRows(white,'   ');
 assert.equal(rows.cover.length,1);assert.deepEqual(rows.cover.map(x=>x.text),[white.cover.title]);assert.doesNotMatch(JSON.stringify(rows),/Author:|Edition:|Audience:|Permissions contact:|Credits:/u);
});

test('all five families full and selected Minimal/Wrench preserve disabled bytes and public learner scope',()=>{
 for(const wrench of [false,true]){const {p,id}=fixture(wrench),before=structuredClone(p);
  for(const planId of ['',id])for(const family of WORKBOOK_FAMILIES){
   const old=prepareFamilyWorkbook(p,config(family.id,planId));
   const off=prepareFamilyWorkbook(p,config(family.id,planId,{...metadata(),includeCover:false,includeCopyright:false}));
   assert.deepEqual(off.output,old.output);assert.equal(familyReviewText(off),familyReviewText(old));
   for(const includeCover of [false,true])for(const includeCopyright of [false,true]){
    const m={...metadata(),includeCover,includeCopyright},on=prepareFamilyWorkbook(p,config(family.id,planId,m));
    assert.equal(on.eligible,true);assert.equal(on.output.profileVersion,3);assert.deepEqual(on.output.preview.sections,old.output.preview.sections);assert.equal(on.output.preview.responseLineTwips,old.output.preview.responseLineTwips);
    const joined=Object.values(on.output.files).join('');assert.doesNotMatch(joined,/PRIVATE_AUDIENCE|PRIVATE_PURPOSE|PRIVATE_AUTHOR|PRIVATE_RIGHTS|PRIVATE_TEACHER|PRIVATE_ANSWERS|INACTIVE_PRIVATE_LICENSE|INACTIVE_PRIVATE_VERSION|INACTIVE_PRIVATE_URL/u);
    if(planId)assert.doesNotMatch(joined,/OMITTED_LEARNER/u);
    assert.equal((joined.match(/Public cover &lt; &amp; &gt;/gu)||[]).length,includeCover?2:0);
    assert.equal((on.output.files['word/document.xml'].match(/>Copyright<\/w:t>/gu)||[]).length,includeCopyright?1:0);
    assert.equal(on.output.preview.title,includeCover?m.cover.title:p.title);
    for(const name of Object.keys(old.output.files).filter(k=>!['word/document.xml','docProps/core.xml'].includes(k)))assert.equal(on.output.files[name],old.output.files[name],name);
   }
  }
  assert.deepEqual(p,before);
 }
});

test('normal-flow rows XML-escape tabs and newlines, keep long notices breakable and align core title/author',async()=>{
 const {p}=fixture(),m=metadata();m.copyright.noticeText='Long notice < & >\r\n\t'+('Readable notice. '.repeat(900));
 const before=JSON.stringify(m),result=prepareFamilyWorkbook(p,config('clearline','',m));assert.equal(JSON.stringify(m),before);
 const files=result.output.files,body=files['word/document.xml'];assert.match(body,/<w:tab\/>/u);assert.match(body,/<w:br\/>/u);assert.match(body,/<w:keepNext w:val="0"\/>/u);assert.doesNotMatch(body,/<w:tbl|<w:drawing|<w:txbxContent/u);
 assert.match(files['docProps/core.xml'],/<dc:creator>Explicit Author<\/dc:creator>/u);assert.match(files['docProps/core.xml'],/<dc:title>Public cover &lt; &amp; &gt; Subtitle/u);
 const errors=[];for(const xml of Object.values(files))new DOMParser({errorHandler:{warning:x=>errors.push(x),error:x=>errors.push(x),fatalError:x=>errors.push(x)}}).parseFromString(xml,'application/xml');assert.deepEqual(errors,[]);
 const blob=await familyWorkbookBlob(files,Zip),zip=await Zip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});assert.equal(Object.keys(zip.files).length,7);assert.equal(await zip.file('word/document.xml').async('string'),body);
 const review=familyReviewText(result);assert.match(review,/Public front matter extension: version 1/u);assert.match(review,/Rights holder: Explicit Rights Holder/u);assert.doesNotMatch(review,/INACTIVE_PRIVATE/u);
 for(const bad of ['bad\u0001','bad\ud800','x'.repeat(16001)])assert.throws(()=>validateWorkbookFrontMatter({...metadata(),copyright:{...metadata().copyright,noticeText:bad}}));
});
