// Hand-off to OpenStage Author: the screens of a storyboard, written as an Author project (a folder of JSON files, zipped), so the
// course can be built from them. Author's own validator reads what this writes. The format is Author's (its docs/format-spec.md and
// schema folder); this file writes the smallest part of it that a storyboard fills in:
//   one scene for each lesson, one slide for each screen, a heading, the on-screen text, and the narration as the voice-over script.
// What a storyboard says that has no place in a slide yet (visual direction, interactions, knowledge checks, feedback, branches,
// animation, media, accessibility and developer notes) goes into a plain "hand-off notes" file beside the project.
import {esc} from './guide-ooxml.mjs';

// 0.49.0 is Author's current format (its tools/validate.py refuses 0.48.0 and asks for the project to be updated). 0.49 added two optional
// words to a project's `meta`, `version` and `courseCode`; they are written only when they hold something (an empty one is an error).
export const AUTHOR_SCHEMA_VERSION='0.49.0';
export const META_LIMITS=Object.freeze({version:20,courseCode:40});
// The theme Author's own sample projects start from (its "Classic" theme).
const THEME=Object.freeze({
  schemaVersion:'0.2.0',id:'classic',name:'Classic',
  colors:{background:'#ffffff',surface:'#dde3ea',text:'#141414',muted:'#3d4654',accent:'#0b5fff',accentStrong:'#0b3fb3',onAccentStrong:'#ffffff',dark:'#0b2a6f',onDark:'#ffffff',good:'#0f5132',bad:'#842029',line:'#cfd6e0'},
  fonts:{heading:'system-ui',body:'system-ui'},
  backgrounds:{default:{type:'solid',color:'background'},dark:{type:'solid',color:'dark'},accent:{type:'solid',color:'accentStrong'}}
});
const VOICE='af_heart'; // Author's default narration voice (Kokoro, on the author's computer)

const pad=(n,w=3)=>String(n).padStart(w,'0');
const json=o=>`${JSON.stringify(o,null,2)}\n`;
const plain=t=>String(t??'').replace(/\s+/g,' ').trim();
const slug=s=>String(s||'course').replace(/\s*:\s*/g,' - ').replace(/[\\/*?"<>|]+/g,' ').replace(/\s+/g,' ').trim()||'course';

const BASE={states:{},triggers:[],timing:{appear:'time',start:0,visibleFor:null},animation:{enter:{effect:'none',duration:500},exit:{effect:'none',duration:500}},provenance:null,simulation:null};
// Object ids are unique across the whole project, so each carries its slide's number.
function textObject({slide,n,name,html,altText,role,tabOrder,geometry,style,narration}){
  return {id:`obj-${pad(slide)}-${n}`,type:'text',name,geometry:{...geometry,rotation:0,z:n*10},content:{html},style,
    accessibility:{altText,role,tabOrder,decorative:false},...BASE,narration:narration||null};
}
const narrationOf=(script,lang)=>plain(script)?{script:String(script).replace(/\s*\n\s*/g,' ').trim(),voice:VOICE,speed:1,lang,pronunciations:[],audioCache:null}:null;

// The optional words of `meta`: the version of the course ("1.0") and its course code ("CMM 360"). Left out when blank; refused when too long.
export function metaWords({version='',courseCode=''}={}){
  const out={};
  for(const [key,value] of [['version',version],['courseCode',courseCode]]){
    const text=plain(value);
    if(!text)continue;
    if(text.length>META_LIMITS[key])throw new Error(`The ${key==='version'?'version':'course code'} can be ${META_LIMITS[key]} characters at most.`);
    out[key]=text;
  }
  return out;
}

// One slide for one screen. The narration is attached to the on-screen text, or to the heading when the screen has no on-screen text.
export function slideFor(screen,index,lang){
  const title=plain(screen.title)||`Screen ${screen.number}`,id=`slide-${pad(index)}`;
  const lines=String(screen.onScreen||'').split('\n').map(plain).filter(Boolean);
  const narration=narrationOf(screen.narration,lang);
  const heading=textObject({slide:index,n:1,name:'Heading',html:`<h1>${esc(title)}</h1>`,altText:title,role:'heading',tabOrder:1,
    geometry:{x:80,y:60,width:1120,height:96},style:{fontFamily:'system-ui',fontSize:40,color:'#141414'},narration:lines.length?null:narration});
  const objects=[heading];
  if(lines.length)objects.push(textObject({slide:index,n:2,name:'On-screen text',html:lines.map(l=>`<p>${esc(l)}</p>`).join(''),altText:lines.join(' '),role:'paragraph',tabOrder:2,
    geometry:{x:80,y:200,width:1120,height:360},style:{fontFamily:'system-ui',fontSize:24,color:'#333333'},narration}));
  return {id,name:title,duration:null,layout:null,background:null,baseLayer:{id:'layer-base',name:'Base',objects},layers:[]};
}

// The files of the project, with paths inside the folder `${name}.osproj`, plus the hand-off notes beside it.
export function authorProject(screens,{title,language='en-US',author='',version='',courseCode='',today=new Date().toISOString().slice(0,10)}){
  if(!screens.length)throw new Error('There are no screens to send yet. Add a screen first.');
  const name=slug(title),root=`${name}.osproj`,slides=screens.map((s,i)=>slideFor(s,i+1,language));
  // One scene for each lesson, in order; screens before any lesson are one scene called "Course".
  const scenes=[];
  screens.forEach((s,i)=>{
    const lesson=plain(s.lesson)||'Course';
    let scene=scenes.at(-1);
    if(!scene||scene.name!==lesson){scene={id:`scene-${scenes.length+1}`,name:lesson,slides:[]};scenes.push(scene);}
    scene.slides.push(slides[i].id);
  });
  const project={
    schemaVersion:AUTHOR_SCHEMA_VERSION,
    meta:{title:plain(title)||'Course',author:plain(author),language,created:today,modified:today,...metaWords({version,courseCode})},
    stage:{width:1280,height:720},variables:[],theme:{ref:'theme.json'},credits:true,completion:{slides:'all'},scenes
  };
  return [
    {name:`${root}/project.json`,text:json(project)},
    {name:`${root}/theme.json`,text:json(THEME)},
    ...slides.map(s=>({name:`${root}/slides/${s.id}.json`,text:json(s)})),
    {name:`${slug(title)} - hand-off notes.txt`,text:handoffNotes(screens,title)}
  ];
}

const NOTE_ORDER=[['visual-direction','Visual direction'],['interaction','Interaction'],['knowledge-check','Knowledge check'],['feedback','Feedback'],['branch','Branch'],['animation-note','Animation note'],['media-asset','Media asset'],['accessibility-note','Accessibility note'],['developer-note','Developer note']];
// What the storyboard says about each screen that the project does not hold yet, for whoever builds the screens in Author.
export function handoffNotes(screens,title){
  const out=[`${plain(title)||'Course'} - notes for building in OpenStage Author`,'',
    'The project in this folder has one slide for each screen, with its heading, its on-screen text and its narration as the slide\'s voice-over script.',
    'Everything else the storyboard says about a screen is below. Build it in the editor.',''];
  let lesson=null;
  for(const s of screens){
    if(s.lesson!==lesson){lesson=s.lesson;if(lesson)out.push(lesson,'='.repeat(Math.min(60,lesson.length)),'');}
    out.push(`Screen ${s.number}: ${plain(s.title)||'(no title)'}`);
    let any=false;
    for(const [id,label] of NOTE_ORDER){
      for(const b of s.blocks.filter(x=>x.cue===id&&plain(x.text))){any=true;out.push(`  ${label}: ${b.text.replace(/\n/g,'\n    ')}`);}
    }
    if(!any)out.push('  (nothing more)');
    out.push('');
  }
  return out.join('\n').trimEnd()+'\n';
}
export const authorZipName=title=>`${slug(title)} - OpenStage Author project.zip`;
