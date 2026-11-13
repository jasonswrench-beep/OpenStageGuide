// Semantic partial starter, not original-template formatting or learning modules.
export const FRONT_PROFILE='wrench-starter',FRONT_VERSION=1,FRONT_SCHEMA=2,FRONT_JOURNAL=5;
export const FRONT_SECTIONS=Object.freeze([['opening','Opening'],['getting-started','Getting started'],['background','Background'],['team','Development team / acknowledgements'],['overview','Program overview']].map(Object.freeze));
export const hasFrontMatter=p=>p?.schema===FRONT_SCHEMA;
const fail=()=>{throw new Error('Unsupported or incomplete front-matter profile. Preserve this guide; use its matching add-in version.');};
const uuid=x=>typeof x==='string'&&/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/.test(x);
export function createFrontMatter(selected=[]){
  if(!Array.isArray(selected)||new Set(selected).size!==selected.length||selected.some(k=>typeof k!=='string'||!FRONT_SECTIONS.slice(1).some(([id])=>id===k)))fail();
  return {profile:FRONT_PROFILE,version:FRONT_VERSION,sections:FRONT_SECTIONS.filter(([kind])=>kind==='opening'||selected.includes(kind)).map(([kind,heading])=>({id:crypto.randomUUID(),kind,heading,body:''}))};
}
export function validateFrontMatter(project,ids=new Set([project.id])){
  if(!hasFrontMatter(project))return;
  const f=project.frontMatter;if(!f||Array.isArray(f)||f.profile!==FRONT_PROFILE||f.version!==FRONT_VERSION||!Array.isArray(f.sections)||f.sections.length<1||f.sections.length>5)fail();let previous=-1;
  for(const [i,s] of f.sections.entries()){
    const at=FRONT_SECTIONS.findIndex(([kind])=>kind===s?.kind);
    if(!uuid(s?.id)||ids.has(s.id)||at<=previous||i===0&&s.kind!=='opening'||typeof s.heading!=='string'||!s.heading.trim()||s.heading.length>50000||typeof s.body!=='string'||s.body.length>50000)fail();ids.add(s.id);previous=at;
  }
}
export const frontMatterFields=p=>hasFrontMatter(p)?p.frontMatter.sections.flatMap(s=>[{tag:`osg:section-heading:${s.id}`,text:s.heading,type:'section-heading',id:s.id,style:'Heading1'},{tag:`osg:section-body:${s.id}`,text:s.body,type:'section-body',id:s.id,style:'Normal'}]):[];
export const frontMatterEvidence=p=>({profile:FRONT_PROFILE,version:FRONT_VERSION,courseId:p.id,sectionIds:p.frontMatter.sections.map(s=>s.id)});
export function validFrontJournal(j,p){
  try{
    const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b),cues=p.modules.flatMap(m=>m.activities.map(a=>a.id));
    return hasFrontMatter(p)&&j?.schema===FRONT_JOURNAL&&j.placement===true&&['course','module','activity','metadata','activity-settings'].includes(j.kind)&&typeof j.txId==='string'&&['prepared','committed'].includes(j.phase)&&same(j.frontMatter,frontMatterEvidence(p))&&same(j.cueChecks,cues)&&(j.phase==='committed'||(j.kind==='course'?j.base===null:hasFrontMatter(j.base)&&same(frontMatterEvidence(j.base),frontMatterEvidence(p))));
  }catch{return false;}
}
