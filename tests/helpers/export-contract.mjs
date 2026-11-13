import assert from 'node:assert/strict';
import path from 'node:path';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {course,addModule,addActivity,ICONS} from '../../src/model.mjs';

export const PARTS=['[Content_Types].xml','_rels/.rels','docProps/core.xml','word/_rels/document.xml.rels','word/document.xml','word/styles.xml'].sort();
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/package/2006/relationships';
const O='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export function exportCorpus({optionalBlanks=false}={}){
  let project=course('Contract course & café');
  const privateCanaries=['PRIVATE_COURSE_DETAILS','PRIVATE_UNKNOWN_ROOT','PRIVATE_SLIDE_BASELINE','PRIVATE_SLIDE_REVIEW','PRIVATE_SLIDE_UNSAVED','PRIVATE_SLIDE_COMPARISON'];
  project.details={audience:privateCanaries[0],purpose:'',authors:'',attribution:'',unknown:'PRIVATE_UNKNOWN_DETAILS'};
  privateCanaries.push(project.details.unknown);
  project.unknown={nested:privateCanaries[1],slideSource:privateCanaries[2],review:privateCanaries[3],unsaved:privateCanaries[4],comparison:privateCanaries[5]};
  for(let m=0;m<2;m++){
    project=addModule(project,`Module ${m+1}`,optionalBlanks&&m===0?'':`Objective ${m+1} — learner & action`).next;
    const module=project.modules[m];module.unknown=`PRIVATE_MODULE_${m}`;privateCanaries.push(module.unknown);
    for(let i=0;i<7;i++){
      const index=m*7+i,secret=`PRIVATE_ACTIVITY_${index}`,material=`PRIVATE_MATERIAL_${index}`,provenance=`PRIVATE_LIBRARY_${index}`;
      project=addActivity(project,module.id,{title:`Activity ${index+1}`,icon:ICONS[index][0],minutes:index===0?0:1.5,participant:`Learner ${index+1} <&>\nSecond line Ω`,facilitator:optionalBlanks?'':secret,answers:optionalBlanks?'':`${secret}_ANSWER`,lines:index%3,
        materials:[{item:material,quantity:2.5,unit:'copies',use:'consumable',notes:`${material}_NOTES`}]}).next;
      const activity=project.modules[m].activities[i];activity.librarySource={entryId:crypto.randomUUID(),name:provenance,format:'openstage-activity-library'};activity.unknown={secret};
      privateCanaries.push(secret,material,provenance,activity.librarySource.entryId);
    }
  }
  return {project,privateCanaries};
}

function parse(xml,name){
  const errors=[];const doc=new DOMParser({errorHandler:{warning:m=>errors.push(m),error:m=>errors.push(m),fatalError:m=>errors.push(m)}}).parseFromString(xml,'application/xml');
  assert.deepEqual(errors,[],`Malformed XML in ${name}`);return doc;
}
function decoded(node){
  let result=node.nodeValue||'';
  for(let i=0;i<(node.attributes?.length||0);i++)result+=' '+node.attributes.item(i).value;
  for(let child=node.firstChild;child;child=child.nextSibling)result+=' '+decoded(child);
  return result;
}
export async function auditExport(blob,{project,privateCanaries,kind,start='09:00'}){
  assert.ok(['workbook','agenda','materials'].includes(kind));
  const zip=await JSZip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true});
  const files=Object.values(zip.files).filter(f=>!f.dir);
  assert.deepEqual(files.map(f=>f.name).sort(),PARTS,'Unexpected package part / possible hidden attachment');
  const docs=new Map();
  for(const file of files){
    const raw=await file.async('string'),doc=parse(raw,file.name),text=raw+' '+decoded(doc)+' '+doc.documentElement.textContent;
    for(const secret of privateCanaries)assert.ok(!text.includes(secret),`Private canary leaked in ${file.name}: ${secret}`);
    assert.doesNotMatch(raw,/openstage\.guide\.project|osg:|customXml|oleObject|altChunk|attachedTemplate/,'Master/attachment metadata leaked');
    docs.set(file.name,doc);
  }
  const types=docs.get('[Content_Types].xml'),T='http://schemas.openxmlformats.org/package/2006/content-types';
  const rows=tag=>{const nodes=types.getElementsByTagNameNS(T,tag);return Array.from({length:nodes.length},(_,i)=>nodes.item(i));};
  assert.deepEqual(rows('Default').map(n=>[n.getAttribute('Extension'),n.getAttribute('ContentType')]),[['rels','application/vnd.openxmlformats-package.relationships+xml'],['xml','application/xml']],'Unexpected default content type');
  assert.deepEqual(rows('Override').map(n=>[n.getAttribute('PartName'),n.getAttribute('ContentType')]).sort(),[
    ['/word/document.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'],
    ['/word/styles.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'],
    ['/docProps/core.xml','application/vnd.openxmlformats-package.core-properties+xml']
  ].sort(),'Unexpected override content type');
  const contracts={
    '_rels/.rels':[['rId1',O+'officeDocument','word/document.xml'],['rId2','http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','docProps/core.xml']],
    'word/_rels/document.xml.rels':[['rId1',O+'styles','styles.xml']]
  };
  for(const [name,expected] of Object.entries(contracts)){
    const rels=[...Array(docs.get(name).getElementsByTagNameNS(R,'Relationship').length)].map((_,i)=>docs.get(name).getElementsByTagNameNS(R,'Relationship').item(i));
    assert.deepEqual(rels.map(r=>[r.getAttribute('Id'),r.getAttribute('Type'),r.getAttribute('Target')]),expected,'Unexpected relationship');
    for(const rel of rels){assert.ok(!rel.hasAttribute('TargetMode'),'External relationship forbidden');const base=name==='_rels/.rels'?'':path.posix.dirname(path.posix.dirname(name));assert.ok(docs.has(path.posix.join(base,rel.getAttribute('Target'))),'Dangling target');}
  }
  const core=docs.get('docProps/core.xml').documentElement;
  const coreNames=[];for(let n=core.firstChild;n;n=n.nextSibling)if(n.nodeType===1)coreNames.push(n.nodeName);
  assert.deepEqual(coreNames,['dc:title','dc:creator','dcterms:created'],'Unexpected core property');
  assert.equal(core.getElementsByTagName('dc:title').item(0).textContent,kind==='materials'?'Materials checklist':project.title+(kind==='workbook'?' Participant workbook':' Planned agenda'));
  assert.equal(core.getElementsByTagName('dc:creator').item(0).textContent,'OpenStage Guide');
  const doc=docs.get('word/document.xml');
  for(const tag of ['vanish','webHidden','commentRangeStart','commentReference','del','instrText','object','altChunk'])assert.equal(doc.getElementsByTagNameNS(W,tag).length,0,`Forbidden hidden/revision/attachment node ${tag}`);
  if(kind==='materials')return auditMaterials(docs,{project,privateCanaries});
  const text=[...Array(doc.getElementsByTagNameNS(W,'t').length)].map((_,i)=>doc.getElementsByTagNameNS(W,'t').item(i).textContent).join('\n');
  assert.ok(text.includes(project.title));assert.ok(text.includes(project.id));assert.ok(text.includes(`Generation revision: ${project.revision}`));assert.ok(text.includes('Full course · current planned order'));
  const pnodes=doc.getElementsByTagNameNS(W,'p'),paragraphs=Array.from({length:pnodes.length},(_,i)=>{
    const node=pnodes.item(i),texts=node.getElementsByTagNameNS(W,'t'),styles=node.getElementsByTagNameNS(W,'pStyle');
    return {node,text:Array.from({length:texts.length},(_,j)=>texts.item(j).textContent).join('\n'),style:styles.item(0)?.getAttributeNS(W,'val')};
  });
  const clock=seconds=>{const days=Math.floor(seconds/86400),rest=seconds%86400,h=Math.floor(rest/3600),m=Math.floor(rest%3600/60),s=rest%60;return (days?`Day ${days+1} · `:'')+String(h).padStart(2,'0')+':'+String(m).padStart(2,'0')+(s?':'+String(s).padStart(2,'0'):'');};
  const [hours,minutes]=start.split(':').map(Number);let elapsed=hours*3600+minutes*60;
  const expectedSections=[],entry=(style,text,borders=0)=>({style,text,borders});
  let previous=-1,paragraphIndex=-1;
  for(const module of project.modules){
    expectedSections.push(entry('Heading1',module.title));
    if(kind==='workbook'&&module.objective)expectedSections.push(entry('Heading2','Learning objective'),entry('Normal',module.objective));
    const at=text.indexOf(module.title);assert.ok(at>previous,'Module order');previous=at;
    for(const activity of module.activities){const next=text.indexOf(activity.title+'\n',previous);assert.ok(next>previous,'Activity order');previous=next;
      expectedSections.push(entry('Heading2',activity.title));
      const at=paragraphs.findIndex((p,i)=>i>paragraphIndex&&p.style==='Heading2'&&p.text===activity.title);assert.ok(at>paragraphIndex,'Activity heading missing');paragraphIndex=at;
      const following=paragraphs.findIndex((p,i)=>i>at&&['Heading1','Heading2'].includes(p.style));const section=paragraphs.slice(at+1,following<0?undefined:following);
      if(kind==='workbook'){
        assert.equal(section.filter(p=>p.text===activity.participant).length,1,`Exactly one learner paragraph for ${activity.title}`);
        expectedSections.push(entry('Normal',activity.participant),...Array.from({length:activity.lines},()=>entry('',' ',1)));
        assert.equal(section.reduce((sum,p)=>sum+p.node.getElementsByTagNameNS(W,'bottom').length,0),activity.lines,`Response space for ${activity.title}`);
      }else{
        const end=elapsed+Math.round(activity.minutes*60);
        expectedSections.push(entry('Normal',`${clock(elapsed)}–${clock(end)} · ${activity.minutes} minutes`));
        assert.equal(section[0]?.text,`${clock(elapsed)}–${clock(end)} · ${activity.minutes} minutes`,`Agenda timing for ${activity.title}`);elapsed=end;
      }
      for(const line of activity.participant.split('\n'))assert.equal(text.includes(line),kind==='workbook','Learner text contract');
    }
    if(module.objective)assert.equal(text.includes(module.objective),kind==='workbook','Objective contract');
  }
  assert.match(paragraphs[4]?.text||'',/^Generated: \d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?Z$/);
  const expectedPrefix=[entry('Title',project.title),entry('Subtitle',kind==='workbook'?'Participant workbook':'Planned workshop agenda'),entry('Normal','Full course · current planned order'),entry('Normal',`Project ID: ${project.id} · Generation revision: ${project.revision}`),entry('Normal',paragraphs[4].text)];
  if(kind==='agenda'){const total=project.modules.flatMap(m=>m.activities).reduce((sum,a)=>sum+a.minutes,0);expectedPrefix.push(entry('Normal',`${total} minutes · ${start}–${clock(elapsed)}`));}
  assert.deepEqual(paragraphs.map(p=>entry(p.style||'',p.text,p.node.getElementsByTagNameNS(W,'bottom').length)),[...expectedPrefix,...expectedSections],'Exact output paragraph/heading/activity/timing sequence');
  if(kind==='agenda'){const total=project.modules.flatMap(m=>m.activities).reduce((sum,a)=>sum+a.minutes,0);assert.ok(paragraphs.some(p=>p.text===`${total} minutes · ${start}–${clock(elapsed)}`),'Agenda total/end time');}
  assert.equal(doc.getElementsByTagNameNS(W,'bottom').length,kind==='workbook'?project.modules.flatMap(m=>m.activities).reduce((sum,a)=>sum+a.lines,0):0,'Response lines');
  assert.equal(paragraphs.filter(p=>p.style==='Heading2'&&p.text==='Learning objective').length,kind==='workbook'?project.modules.filter(m=>m.objective).length:0,'Optional objective headings');
  return {kind,parts:files.length,privateCanariesChecked:privateCanaries.length,activities:project.modules.flatMap(m=>m.activities).length};
}

// Independent public projection and arithmetic; never imports the producer/planner.
function auditMaterials(docs,{project,privateCanaries}){
  const groups=new Map();let configured=0,total=0;
  project.modules.forEach((module,mi)=>module.activities.forEach((activity,ai)=>{
    total++;const list=activity.materials||[];if(list.length)configured++;
    for(const entry of list){const clean=s=>s.normalize('NFKC').trim().replace(/\s+/gu,' '),key=JSON.stringify([clean(entry.item).toLowerCase(),clean(entry.unit),entry.use]);if(!groups.has(key))groups.set(key,{item:entry.item.replace(/\r\n?/g,'\n').trim(),unit:entry.unit.replace(/\r\n?/g,'\n').trim(),use:entry.use,values:[],labels:[]});const group=groups.get(key);group.values.push(Math.round(Number(entry.quantity)*1000));group.labels.push(`${mi+1}.${ai+1} ${activity.title.replace(/\r\n?/g,'\n').trim()||'(Untitled activity)'}`);}
  }));
  const doc=docs.get('word/document.xml'),paragraphs=Array.from(doc.getElementsByTagNameNS(W,'p')).map(n=>({style:n.getElementsByTagNameNS(W,'pStyle')[0]?.getAttributeNS(W,'val'),text:Array.from(n.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join('\n')})),created=docs.get('docProps/core.xml').getElementsByTagName('dcterms:created')[0].textContent;
  assert.match(created,/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/);
  const row=(style,text)=>({style,text}),expected=[row('Title','Materials checklist'),row('Normal','Full course · current saved course order · sequential planning'),row('Normal',`Project ID: ${project.id} · Generation revision: ${project.revision}`),row('Normal',`Generated: ${created}`),row('Normal',`${configured} of ${total} activities have recorded materials · ${groups.size} material groups`),row('Normal','Recorded requirements only. Completeness is not certified. Activities without entries may still need materials.'),row('Normal','Full course in sequential planned order. Consumable or retained supplies add. Reusable means a shared pool available again for the next activity; use the largest single-activity requirement. Quantities are explicit activity totals, not per-learner counts. Unit case stays distinct; no conversion, parallel-work allowance or stock subtraction.')];
  if(!groups.size)expected.push(row('Normal','No materials recorded'));
  for(const group of groups.values()){const quantity=(group.use==='consumable'?group.values.reduce((a,b)=>a+b,0):Math.max(...group.values))/1000;expected.push(row('Heading1',`☐ ${group.item}`),row('Normal',`Quantity: ${quantity} · Unit: ${group.unit} · ${group.use==='consumable'?'Consumed or retained':'Reusable shared pool'}`),row('Normal',`Activities: ${group.labels.join('; ')}`));}
  assert.deepEqual(paragraphs,expected,'Exact checklist scope/coverage/group totals/units/activity order');
  for(const name of ['bottom','sdt','drawing'])assert.equal(doc.getElementsByTagNameNS(W,name).length,0);
  const styles=docs.get('word/styles.xml');assert.deepEqual(Array.from(styles.getElementsByTagNameNS(W,'style')).map(n=>n.getAttributeNS(W,'styleId')),['Normal','Title','Heading1']);assert.equal(styles.getElementsByTagNameNS(W,'outlineLvl')[0].getAttributeNS(W,'val'),'0');
  return {kind:'materials',parts:docs.size,privateCanariesChecked:privateCanaries.length,activities:total};
}
