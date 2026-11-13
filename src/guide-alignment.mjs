// Learning alignment on the guide: which objectives have an activity and an assessment item. Nothing is stored. The guide is read from the
// document, and things are tied together only by what they say:
//   - an objective is an Objective block in a lesson;
//   - an Activity is tied to it when the Activity's own Objective part says the same thing;
//   - an assessment item is tied to it when the item's Learning objective says the same thing.
// "The same thing" is exact words (capitals, spaces and punctuation aside), or "similar" when most of the words match. A similar match
// counts, and is marked, so a small rewording does not break the link. The older tool kept links by hand against a course record.
import {readGuide,readBlock,readAssessments,readItem,kindOf,cueFromText} from './guide-model.mjs';

export const SIMILAR_AT=0.6;
const norm=t=>String(t??'').toLowerCase().replace(/[^\p{L}\p{N}\s]/gu,' ').replace(/\s+/g,' ').trim();
const words=t=>new Set(norm(t).split(' ').filter(w=>w.length>2));
const firstLine=t=>String(t??'').split(/\r?\n/)[0].trim();

// 'same' for equal words, 'similar' for most words shared, null for neither.
export function likeness(a,b){
  const x=norm(a),y=norm(b);
  if(!x||!y)return null;
  if(x===y)return 'same';
  const wa=words(a),wb=words(b);
  if(!wa.size||!wb.size)return null;
  let shared=0;for(const w of wa)if(wb.has(w))shared++;
  return shared/(wa.size+wb.size-shared)>=SIMILAR_AT?'similar':null;
}

export function alignment(paras){
  const guide=readGuide(paras);
  const objectives=[],activities=[],items=[];
  for(const mod of guide.modules)for(const lesson of mod.lessons){
    for(let i=lesson.index+1;i<=lesson.end;i++){
      if(kindOf(paras[i])!=='cue')continue;
      const cue=cueFromText(paras[i].text);
      if(!cue)continue;
      if(cue.id==='objective'){
        const b=readBlock(paras,i);
        const text=firstLine(b?b.text:'');
        if(text)objectives.push({index:i,text,lesson:lesson.text,module:mod.text||mod.title,activities:[],items:[]});
      }else if(cue.id==='activity'){
        const b=readBlock(paras,i);
        if(b)activities.push({index:i,title:firstLine(b.text),objective:String((b.parts&&b.parts.objective)||'').split(/\r?\n/)[0].trim(),lesson:lesson.text});
      }
    }
  }
  for(const section of readAssessments(paras).sections)for(const it of section.items){
    // A survey item has no right answer and no objective to check against: only items with a key are measured against objectives.
    if(!it.hasKey)continue;
    const full=readItem(paras,it.index);
    items.push({index:it.index,label:`${it.label} in ${section.title||'an assessment'}`,objective:String(full&&full.objective||'').split(/\r?\n/)[0].trim()});
  }
  const tied=new Set();
  for(const o of objectives){
    for(const a of activities){const l=likeness(o.text,a.objective);if(l){o.activities.push({...a,similar:l==='similar'});tied.add(a);}}
    for(const it of items){const l=likeness(o.text,it.objective);if(l){o.items.push({...it,similar:l==='similar'});tied.add(it);}}
  }
  const loose=[];
  for(const a of activities)if(!tied.has(a))loose.push({kind:'activity',index:a.index,text:a.title||'An activity',why:a.objective?`its objective, "${a.objective}", matches none of your lesson objectives.`:'it has no Objective part.'});
  for(const it of items)if(!tied.has(it))loose.push({kind:'item',index:it.index,text:it.label,why:it.objective?`its learning objective, "${it.objective}", matches none of your lesson objectives.`:'it has no learning objective.'});
  const practised=objectives.filter(o=>o.activities.length).length,checked=objectives.filter(o=>o.items.length).length,both=objectives.filter(o=>o.activities.length&&o.items.length).length;
  return {objectives,loose,summary:{objectives:objectives.length,practised,checked,both}};
}

export function alignmentSummary(a){
  const s=a.summary;
  if(!s.objectives)return 'There are no objectives yet. Add an Objective block to a lesson, then write the same words in an Activity\'s Objective part and in an assessment item\'s Learning objective.';
  const n=(c,w)=>`${c} ${c===1?w:w+'s'}`;
  return `${n(s.objectives,'objective')}: ${s.practised} with an activity, ${s.checked} with an assessment item, ${s.both} with both.`;
}
