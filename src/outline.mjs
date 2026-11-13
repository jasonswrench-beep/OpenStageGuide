import {ICONS,validate} from './model.mjs';
import {reviewCourse} from './planning.mjs';
import {objectiveText,unsupportedObjective} from './objective-text.mjs';
const normal=s=>s.normalize('NFKD').replace(/\p{M}/gu,'').toLowerCase();
const unfinished=texts=>texts.some(s=>typeof s==='string'&&/\b(?:TODO|TBD)\b|\[(?:insert|add|replace)\b[^\]]*\]/i.test(s));
export function searchOutline(project,{query='',icon='',status='all'}={}){
  validate(project);
  if(typeof query!=='string'||query.length>200||typeof icon!=='string'||icon&&!ICONS.some(([label])=>label===icon)||!['all','unfinished','needs-attention','missing-objective','no-recorded-materials'].includes(status))throw new Error('Choose supported outline filters and search text up to 200 characters.');
  const terms=normal(query.trim()).split(/\s+/u).filter(Boolean),matches=texts=>{const text=normal(texts.filter(s=>typeof s==='string').join('\n'));return terms.every(t=>text.includes(t));};
  const errors=new Set(reviewCourse(project).filter(x=>x.severity==='error').map(x=>x.id)),rows=[];
  const accept=(id,texts,missingObjective,noRecordedMaterials=false)=>status==='all'||status==='unfinished'&&unfinished(texts)||status==='needs-attention'&&errors.has(id)||status==='missing-objective'&&missingObjective||status==='no-recorded-materials'&&noRecordedMaterials;
  project.modules.forEach((m,i)=>{
    const moduleText=[m.title,objectiveText(m)],missingObjective=!unsupportedObjective(m)&&!objectiveText(m).trim(),activities=[];
    m.activities.forEach((a,j)=>{
      const own=[a.title,a.facilitator,a.participant,a.answers];
      const noRecordedMaterials=a.materials===undefined||Array.isArray(a.materials)&&a.materials.length===0;
      if((!icon||a.icon===icon)&&matches([...moduleText,...own,a.icon])&&accept(a.id,own,missingObjective,noRecordedMaterials))activities.push({type:'activity',id:a.id,moduleId:m.id,moduleIndex:i,activityIndex:j,title:a.title,icon:a.icon,minutes:a.minutes});
    });
    const ownMatch=!icon&&matches(moduleText)&&accept(m.id,moduleText,missingObjective);
    if(ownMatch||activities.length){rows.push({type:'module',id:m.id,moduleId:m.id,moduleIndex:i,title:m.title,activityCount:m.activities.length,minutes:m.activities.reduce((n,a)=>n+a.minutes,0),context:!ownMatch});rows.push(...activities);}
  });
  return {rows,moduleCount:rows.filter(r=>r.type==='module').length,activityCount:rows.filter(r=>r.type==='activity').length,totalModules:project.modules.length,totalActivities:project.modules.reduce((n,m)=>n+m.activities.length,0),filtered:!!terms.length||!!icon||status!=='all'};
}
