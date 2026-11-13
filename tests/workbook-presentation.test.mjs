import test from 'node:test';
import assert from 'node:assert/strict';
import {course,wrenchCourse,addModule,addActivity} from '../src/model.mjs';
import {editDeliveryPlan} from '../src/delivery-plans.mjs';
import {workbookPreview,worksheetFiles,WORKBOOK_SCOPE} from '../src/worksheet.mjs';
import {planWorkbookPreview,planWorkbookFiles,SELECTED_OUTPUT_SCOPE} from '../src/selected-plan-outputs.mjs';
import {resolveSelectedPlanScope} from '../src/selected-plan-scope.mjs';
import {WORKBOOK_FAMILIES,WORKBOOK_SEMANTIC_VERSION,WORKBOOK_MIN_TEXT_CONTRAST,WORKBOOK_MAX_SEMANTIC_NODES,WORKBOOK_MAX_SEMANTIC_TEXT_UNITS,workbookFamily,semanticWorkbook,semanticWorkbookFromCourse,semanticWorkbookFromSelected,resolveWorkbookBrand,textContrast} from '../src/workbook-presentation.mjs';

const stamp='2026-10-26T10:00:00.000Z';
function fixture(){
 let p=course('Learning 🌱');
 for(let i=0;i<3;i++)p=addModule(p,`Module ${i}`,`Objective ${i}`).next;
 const add=(mi,name,lines)=>{p=addActivity(p,p.modules[mi].id,{title:name,minutes:10,participant:`Learner ${name}\nSecond line`,facilitator:`PRIVATE_FAC_${name}`,answers:`PRIVATE_ANSWER_${name}`,lines}).next;};
 add(0,'OMITTED',0);add(0,'A',1);add(1,'B',20);add(1,'C',0);
 const groups=[{moduleId:p.modules[2].id,activityIds:[]},{moduleId:p.modules[1].id,activityIds:[p.modules[1].activities[1].id,p.modules[1].activities[0].id]},{moduleId:p.modules[0].id,activityIds:[p.modules[0].activities[1].id]}];
 const created=editDeliveryPlan(p,{kind:'create',name:'Reviewed',targetMinutes:30,groups});p=created.next;
 p.opaque={secret:'PRIVATE_OPAQUE'};p.details={audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_ATTRIBUTION'};
 return {p,planId:created.id};
}
const fullIdentity=p=>({projectId:p.id,revision:p.revision,modules:p.modules.map(m=>({id:m.id,activityIds:m.activities.map(a=>a.id)}))});
const selectedIdentity=(p,planId)=>({projectId:p.id,revision:p.revision,modules:resolveSelectedPlanScope(p,planId).groups.map(g=>({id:g.moduleId,activityIds:g.activityIds}))});
const full=(p)=>semanticWorkbook(workbookPreview(p),fullIdentity(p),{kind:'full',label:WORKBOOK_SCOPE});
const selected=(p,planId)=>semanticWorkbook(planWorkbookPreview(p,planId,stamp),selectedIdentity(p,planId),{kind:'selected',planId,label:SELECTED_OUTPUT_SCOPE});

test('five stable, distinct original registry definitions have guarded defaults and no unsupported claims',()=>{
 assert.equal(WORKBOOK_FAMILIES.length,5);
 assert.deepEqual(WORKBOOK_FAMILIES.map(f=>f.id),['clearline','editorial-study','practice-lab','field-guide','reflection-journal']);
 assert.equal(new Set(WORKBOOK_FAMILIES.map(f=>f.direction)).size,5);
 for(const family of WORKBOOK_FAMILIES){
  assert.equal(workbookFamily(family.id),family);assert.equal(Object.isFrozen(family),true);
  assert.equal(family.semanticVersion,WORKBOOK_SEMANTIC_VERSION);
  assert.equal(family.accentUse,'decorative-only');
  assert.deepEqual(Object.values(family.capabilities),[false,false,false,false,false]);
  const brand=resolveWorkbookBrand(family.id);
  assert.equal(brand.eligible,true,`${family.id}: ${brand.issues}`);
  assert.ok(textContrast(brand.effective.text,brand.effective.background)>=WORKBOOK_MIN_TEXT_CONTRAST);
  assert.ok(textContrast(brand.effective.accentText,brand.effective.accentBackground)>=WORKBOOK_MIN_TEXT_CONTRAST);
 }
 assert.throws(()=>workbookFamily('unlisted'),/unavailable/);
});

test('full and selected semantics retain exact learner order, IDs, response counts and policy text',()=>{
 const {p,planId}=fixture(),before=structuredClone(p),f=full(p),s=selected(p,planId);
 assert.equal(f.version,1);assert.equal(f.scope.kind,'full');assert.equal(s.scope.kind,'selected');assert.equal(s.scope.planId,planId);
 assert.deepEqual(f.nodes.map(n=>n.id),p.modules.flatMap(m=>[m.id,...m.activities.map(a=>a.id)]));
 assert.deepEqual(s.nodes.map(n=>n.title),['Module 1','C','B','Module 0','A']);
 assert.deepEqual(s.nodes.filter(n=>n.kind==='activity').map(n=>n.responseLines),[0,20,1]);
 assert.ok(s.identityLines.some(x=>/does not confirm objective coverage/i.test(x)));
 const serialized=JSON.stringify(s);
 for(const absent of ['OMITTED','Objective 2','PRIVATE_FAC','PRIVATE_ANSWER','PRIVATE_OPAQUE','PRIVATE_AUDIENCE','PRIVATE_PURPOSE','PRIVATE_AUTHOR','PRIVATE_ATTRIBUTION'])assert.ok(!serialized.includes(absent),absent);
 assert.equal(Object.isFrozen(s.nodes),true);assert.equal(Object.isFrozen(p),false);assert.equal(Object.isFrozen(p.modules[0]),false);
 assert.deepEqual(p,before);
 assert.deepEqual(semanticWorkbookFromCourse(p),f);
 assert.deepEqual(semanticWorkbookFromSelected(p,planId,stamp),s);
});

test('semantic mapping is pure, stable, and independent of caller objects',()=>{
 const {p,planId}=fixture(),a=selected(p,planId),b=selected(p,planId);
 assert.deepEqual(a,b);assert.notEqual(a,b);
 const preview=planWorkbookPreview(p,planId,stamp),identity=structuredClone(selectedIdentity(p,planId));
 const c=semanticWorkbook(preview,identity,{kind:'selected',planId,label:SELECTED_OUTPUT_SCOPE});
 assert.deepEqual(a,c);assert.equal(Object.isFrozen(preview),false);assert.equal(Object.isFrozen(identity),false);
 preview.modules[0].activities[0].participant='Changed';identity.modules[0].activityIds[0]='changed';
 assert.deepEqual(c,a);
 for(const family of WORKBOOK_FAMILIES)assert.deepEqual(selected(p,planId).nodes,a.nodes,`family ${family.id} must not rewrite semantics`);
});

test('semantic contract refuses mismatched roster, unsupported nodes, inherited/accessor and null boundaries',()=>{
 const {p,planId}=fixture(),preview=planWorkbookPreview(p,planId,stamp),identity=selectedIdentity(p,planId),scope={kind:'selected',planId,label:SELECTED_OUTPUT_SCOPE};
 const rejects=(v=preview,i=identity,s=scope)=>assert.throws(()=>semanticWorkbook(v,i,s),/Unsupported workbook presentation/);
 rejects(null);rejects(Object.create(preview));
 rejects({...preview,modules:[...preview.modules,{title:'extra',objective:'',activities:[]}]});
 rejects({...preview,planId:'wrong'});
 rejects(preview,identity,{...scope,version:2});
 rejects(preview,identity,{...scope,version:null});
 rejects(preview,identity,{...scope,label:'Unreviewed scope'});
 rejects(preview,{...identity,modules:[...identity.modules].reverse()});
 rejects(preview,{...identity,modules:identity.modules.map((m,index)=>index?m:{...m,id:'same-length-foreign-id'})});
 rejects(preview,{...identity,modules:identity.modules.map((m,index)=>index?m:{...m,activityIds:['same-length-foreign-id',...m.activityIds.slice(1)]})});
 rejects({...preview,modules:preview.modules.map((m,index)=>index?m:{...m,answers:'PRIVATE'})});
 rejects({...preview,modules:preview.modules.map((m,index)=>index?m:{...m,activities:[{...m.activities[0],lines:21},...m.activities.slice(1)]})});
 const accessor={...preview};Object.defineProperty(accessor,'title',{get(){throw Error('getter called');},enumerable:true});rejects(accessor);
 const hidden={...preview};Object.defineProperty(hidden,'title',{value:preview.title,enumerable:false});rejects(hidden);
 const inherited=Object.create({kind:'selected'});inherited.planId=planId;inherited.label=SELECTED_OUTPUT_SCOPE;rejects(preview,identity,inherited);
 const hole={...preview,modules:[,preview.modules[0]]};rejects(hole);
 const extra=[...preview.modules];Object.defineProperty(extra,'secret',{value:'PRIVATE',enumerable:false});rejects({...preview,modules:extra});
 const v2={...preview,modules:preview.modules.map((m,index)=>index?m:{...m,activities:[{kind:'unknown',...m.activities[0]},...m.activities.slice(1)]})};rejects(v2);
});

test('brand rejects unsafe contrast without silently changing requested colors',()=>{
 const request={organizationName:'Example College',text:'#DDDDDD',background:'#FFFFFF',accentText:'#EEEEEE',accentBackground:'#FFFFFF'};
 const before=structuredClone(request),result=resolveWorkbookBrand('clearline',request);
 assert.equal(result.eligible,false);assert.equal(result.effective,null);assert.deepEqual(result.rawRequest,request);assert.deepEqual(request,before);
 assert.equal(result.proposedSafeAlternative.text,'#000000');assert.equal(result.proposedSafeAlternative.accentText,'#000000');
 const corrected=resolveWorkbookBrand('clearline',{...request,...result.proposedSafeAlternative});
 assert.equal(corrected.eligible,true);assert.equal(corrected.effective.text,'#000000');assert.equal(request.text,'#DDDDDD');
 assert.ok(textContrast('#777777','#FFFFFF')<4.5);assert.ok(textContrast('#767676','#FFFFFF')>=4.5);
});

test('brand bounds, explicit unsupported capabilities, own-data refusal, and pure repeated resolution',()=>{
 for(const bad of [{bodyPt:8},{headingPt:29},{marginIn:0.2},{lineSpacing:0},{pageSize:'Tabloid'},{printMode:'invisible'},{headingFont:'Comic Sans MS'},{fallbackFont:'Unknown'},{background:'transparent'},{organizationName:'x'.repeat(121)},{organizationName:'bad\ud800'},{organizationName:'bad\ufffe'},{logo:'data:image/png'}]){
  if('logo' in bad)assert.throws(()=>resolveWorkbookBrand('clearline',bad),/unsupported/);
  else assert.equal(resolveWorkbookBrand('clearline',bad).eligible,false,JSON.stringify(bad));
 }
 for(const bad of [null,[],Object.create({text:'#000000'}),{text:{value:'#000000'}}])assert.throws(()=>resolveWorkbookBrand('clearline',bad),/Unsupported workbook presentation/);
 const accessor={};Object.defineProperty(accessor,'text',{get(){throw Error('getter called');},enumerable:true});assert.throws(()=>resolveWorkbookBrand('clearline',accessor),/data field/);
 const hidden={};Object.defineProperty(hidden,'bodyFont',{get(){throw Error('hidden getter called');},enumerable:false});assert.throws(()=>resolveWorkbookBrand('clearline',hidden),/enumerable data field/);
 const hiddenExtra={};Object.defineProperty(hiddenExtra,'logo',{value:'PRIVATE',enumerable:false});assert.throws(()=>resolveWorkbookBrand('clearline',hiddenExtra),/unsupported/);
 const explicitUndefined={};Object.defineProperty(explicitUndefined,'text',{value:undefined,enumerable:true});assert.throws(()=>resolveWorkbookBrand('clearline',explicitUndefined),/JSON scalar/);
 for(const bad of [()=>{},Symbol('bad'),1n,NaN,Infinity])assert.throws(()=>resolveWorkbookBrand('clearline',{text:bad}),/JSON scalar|finite/);
 const raw={organizationName:'College',bodyPt:12};const a=resolveWorkbookBrand('editorial-study',raw),b=resolveWorkbookBrand('editorial-study',raw);
 assert.deepEqual(a,b);assert.equal(Object.isFrozen(raw),false);assert.equal(Object.isFrozen(a),true);
 assert.equal(a.effective.organizationName,'College');assert.equal(a.effective.bodyPt,12);
});

test('new pure contracts leave fixed-stamp legacy full and selected part maps byte-identical',()=>{
 const {p,planId}=fixture(),beforeFull=worksheetFiles(p,stamp),beforeSelected=planWorkbookFiles(p,planId,stamp),before=structuredClone(p);
 full(p);selected(p,planId);for(const family of WORKBOOK_FAMILIES)resolveWorkbookBrand(family.id,{organizationName:'College'});
 assert.deepEqual(worksheetFiles(p,stamp),beforeFull);assert.deepEqual(planWorkbookFiles(p,planId,stamp),beforeSelected);assert.deepEqual(p,before);
});

test('Wrench front matter stays outside learner semantics and legacy workbook output',()=>{
 let p=wrenchCourse('Workshop',{audience:'PRIVATE_AUDIENCE',purpose:'PRIVATE_PURPOSE',authors:'PRIVATE_AUTHOR',attribution:'PRIVATE_RIGHTS'},['background']);
 p=addModule(p,'Module','Objective').next;
 p=addActivity(p,p.modules[0].id,{title:'Practice',minutes:15,participant:'Use this page.',facilitator:'PRIVATE_TEACH',answers:'PRIVATE_KEY',lines:2}).next;
 const before=worksheetFiles(p,stamp),source=structuredClone(p),semantic=semanticWorkbookFromCourse(p);
 assert.deepEqual(semantic.nodes.map(n=>n.kind),['module','activity']);
 assert.deepEqual(semantic.nodes.filter(n=>n.kind==='activity').map(n=>n.responseLines),[2]);
 for(const forbidden of ['PRIVATE_AUDIENCE','PRIVATE_PURPOSE','PRIVATE_AUTHOR','PRIVATE_RIGHTS','PRIVATE_TEACH','PRIVATE_KEY'])assert.ok(!JSON.stringify(semantic).includes(forbidden));
 assert.deepEqual(worksheetFiles(p,stamp),before);assert.deepEqual(p,source);
});

test('aggregate node and published-text budgets refuse only the new route without altering caller data',()=>{
 const base={projectId:'course',revision:1,title:'Course',modules:[]};
 const scope={kind:'full',label:WORKBOOK_SCOPE};
 const identity={projectId:'course',revision:1,modules:[]};
 assert.equal(WORKBOOK_MAX_SEMANTIC_NODES,2000);assert.equal(WORKBOOK_MAX_SEMANTIC_TEXT_UNITS,250000);
 for(let i=0;i<WORKBOOK_MAX_SEMANTIC_NODES;i++){
  base.modules.push({title:'M',objective:'',activities:[]});identity.modules.push({id:`m${i}`,activityIds:[]});
 }
 assert.equal(semanticWorkbook(base,identity,scope).nodes.length,2000);
 base.modules.push({title:'M',objective:'',activities:[]});identity.modules.push({id:'m2000',activityIds:[]});
 assert.throws(()=>semanticWorkbook(base,identity,scope),/node budget/);
 assert.equal(base.modules.length,2001);assert.equal(Object.isFrozen(base),false);
 const six=n=>({projectId:'course',revision:1,title:'Course',modules:[{title:'M',objective:'',activities:Array.from({length:6},(_,i)=>({title:`A${i}`,participant:'x'.repeat(n),lines:0}))}]});
 const sixIds={projectId:'course',revision:1,modules:[{id:'m',activityIds:Array.from({length:6},(_,i)=>`a${i}`)}]};
 assert.equal(semanticWorkbook(six(40000),sixIds,scope).nodes.length,7);
 const tooLong=six(42000);assert.throws(()=>semanticWorkbook(tooLong,sixIds,scope),/text budget/);
 assert.equal(Object.isFrozen(tooLong.modules[0].activities[0]),false);
 assert.throws(()=>semanticWorkbook({...six(1),title:'bad\ufffe'},sixIds,scope),/supported text/);
 assert.throws(()=>semanticWorkbook({...six(1),title:'bad\ud800'},sixIds,scope),/supported text/);
});
