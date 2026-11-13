import {KEY,PENDING} from '../../src/document-store.mjs';
import {course,addModule,addActivity} from '../../src/model.mjs';
import {frontMatterFields} from '../../src/frontmatter.mjs';
const copy=x=>x==null?x:structuredClone(x);
let sequence=0;
export function fixture(){
  let p=addModule(course('Recovery sample'),'Original module','Practice a response.').next;
  return addActivity(p,p.modules[0].id,{title:'Original activity',minutes:10,participant:'Learner task.',facilitator:'PRIVATE DIRECTIONS',answers:'SECRET ANSWER',lines:2}).next;
}
export function records(p){return p?[{tag:`osg:course:${p.id}`,text:p.title},...frontMatterFields(p).map(({tag,text})=>({tag,text})),...p.modules.flatMap(m=>[
  {tag:`osg:module:${m.id}`,text:m.title},{tag:`osg:objective:${m.id}`,text:m.objective},
  ...m.activities.flatMap(a=>['activity','facilitator','participant','answers'].map(type=>({tag:`osg:${type}:${a.id}`,text:type==='activity'?a.title:a[type]})))
])]:[];}
export function makeHost(project){
  const h={durable:project?{[KEY]:copy(project)}:{},memory:{},records:records(project),paragraphs:[],pictures:[],selection:null,
    saves:0,refreshes:0,syncs:0,runs:0,saveFailures:new Map(),refreshFailures:new Set(),syncFailures:new Map(),hooks:{},
    cues:Object.fromEntries((project?.modules.flatMap(m=>m.activities)||[]).map(a=>[a.id,`${a.icon} · ${a.minutes} min`])),cueWrites:0};
  h.memory=copy(h.durable);
  h.settings={
    get:key=>copy(h.memory[key]),set:(key,value)=>{h.memory[key]=copy(value);},remove:key=>{delete h.memory[key];},
    saveAsync:callback=>{
      const n=++h.saves,fail=h.saveFailures.get(n);h.hooks.beforeSave?.(n,h);
      if(!fail||fail.persist)h.durable=copy(h.memory);
      queueMicrotask(()=>callback(fail?{status:'failed',error:{message:'Injected settings failure'}}:{status:'succeeded'}));
    },
    refreshAsync:callback=>{
      const n=++h.refreshes;h.hooks.beforeRefresh?.(n,h);
      if(!h.refreshFailures.has(n))h.memory=copy(h.durable);
      queueMicrotask(()=>callback(h.refreshFailures.has(n)?{status:'failed',error:{message:'Injected refresh failure'}}:{status:'succeeded'}));
    }
  };
  h.Office={context:{document:{settings:h.settings},host:'Word',requirements:{isSetSupported:()=>true}},HostType:{Word:'Word'},AsyncResultStatus:{Succeeded:'succeeded'},
    onReady:async()=>{},actions:{associate:()=>{}},addin:{showAsTaskpane:async()=>{}}};
  h.Word={run:async callback=>{
    h.runs++;let queue=[];
    const collection=items=>({load(){},get items(){return items();}});
    const paragraph=(text,beforeTag=null)=>{
      const para={text,beforeTag,font:{},getRange:()=>({
        insertContentControl(){const c={tag:'',text};queue.push(()=>{if(h.dropField?.(c))return;if(beforeTag){const at=h.records.findIndex(r=>r.tag===beforeTag);if(at<0)throw new Error('Mock insertion boundary missing');h.records.splice(at,0,c);}else h.records.push(c);if(c.tag.startsWith('osg:activity:')){const previous=h.paragraphs.at(-2)?.text;if(previous?.includes(' · '))h.cues[c.tag.split(':').at(-1)]=previous;}});return c;},
        insertInlinePictureFromBase64(){const p={};queue.push(()=>h.pictures.push(p));return p;}
      })};
      queue.push(()=>h.paragraphs.push(para));return para;
    };
    const nullObject=isNullObject=>({isNullObject,load(){}});
    const cue=id=>({load(){},get text(){return h.cues[id]||'';},isNullObject:!Object.hasOwn(h.cues,id),
      parentContentControlOrNullObject:nullObject(!h.cueNested),parentTableOrNullObject:nullObject(!h.cueTable),
      inlinePictures:collection(()=>h.cuePictures?[{}]:[]),getRange:()=>({contentControls:collection(()=>h.cueChildren?[{}]:[])}),
      search:text=>collection(()=>h.cues[id]?.includes(text)?[{insertText(replacement,where){if(where!=='Replace')throw new Error('Mock duration replacement only');queue.push(()=>{h.cues[id]=h.cues[id].replace(text,replacement);h.cueWrites++;});}}]:[])
    });
    const range=r=>({tag:r.tag,
      select:()=>{queue.push(()=>h.selection=r.tag);},
      compareLocationWith(other){const result={};queue.push(()=>{const order=h.physicalOrder||h.records.map(x=>x.tag);result.value=h.rangeRelations?.[r.tag]||(order.indexOf(r.tag)<order.indexOf(other.tag)?'Before':'After');});return result;},
      parentTableOrNullObject:nullObject(!h.tableTags?.has(r.tag)),
      paragraphs:{load(){},get items(){return [{text:h.paragraphText?.[r.tag]??r.text,getPreviousOrNullObject:()=>cue(r.tag.split(':').at(-1))}];},getFirst:()=>({insertParagraph:(text,location)=>{if(location!=='Before')throw new Error('Unsupported mock boundary');return paragraph(text,r.tag);}})}
    });
    const controls=collection(()=>h.records.map(r=>({...r,parentContentControlOrNullObject:nullObject(!h.nestedTags?.has(r.tag)),getRange:()=>range(r),insertParagraph:(text,location)=>{if(location!=='Before')throw new Error('Unsupported mock control boundary');return paragraph(text,r.tag);}})));
    controls.getByTag=tag=>collection(()=>controls.items.filter(x=>x.tag===tag));
    const ctx={document:{contentControls:controls,body:{getOoxml(){const result={};queue.push(()=>{h.hooks.beforeBodyOoxml?.(h);result.value=h.bodyOoxml();});return result;},insertParagraph:text=>paragraph(text),getRange:()=>({compareLocationWith(other){const result={};queue.push(()=>{result.value=h.unrelatedTags?.has(other.tag)?'Unrelated':'Contains';});return result;}})}},sync:async()=>{
      const n=++h.syncs;h.hooks.beforeSync?.(n,h);
      const fail=h.syncFailures.get(n);
      if(!fail||fail.apply){const ops=queue;queue=[];for(const op of ops)op();}
      if(fail)throw new Error('Injected Word synchronization failure');
    }};
    return callback(ctx);
  }};
  h.reopen=()=>{h.memory=copy(h.durable);return h;};
  return h;
}
export async function adapter(h){
  globalThis.Office=h.Office;globalThis.Word=h.Word;
  globalThis.fetch=async()=>({ok:true,blob:async()=>new Blob(['mock icon'])});
  globalThis.FileReader=class{readAsDataURL(){this.result='data:image/png;base64,AAAA';queueMicrotask(()=>this.onload());}};
  return import(`../../src/word.mjs?mock=${++sequence}`);
}
export function journal(h){return h.durable[PENDING];}
