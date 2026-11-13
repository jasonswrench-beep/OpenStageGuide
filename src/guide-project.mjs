// The project file of OpenStage Guide when it is kept as a program of its own, not inside a Word document.
// A project holds what the Word document holds for the pane: the paragraphs (their words, styles and the Word-format XML
// that made them), the cursor, the guide's own cue names, the chosen look, what the guide remembers about its slides and its
// delivery versions, and what the look asked of the styles and the header and footer. Nothing else: no course record is kept
// apart from the paragraphs, because the guide is whatever the paragraphs say it is.
//
// These functions are pure. They read and write text; they never touch a disk or a document. Anything that arrives from a file
// is checked here before the pane sees it: a wrong kind of file, a project from a newer program, or a shape that is not a guide
// is refused with a plain sentence, and nothing is changed.

export const PROJECT_FORMAT='openstage-guide-project';
export const PROJECT_VERSION=1;
export const LIMITS=Object.freeze({
  fileCharacters:64*1024*1024,
  paragraphs:200000,
  paragraphXml:2*1024*1024,
  words:2*1024*1024,
  styleName:200,
  title:2000,
  cues:50,
  cueName:200,
  looseItems:10000,
  pictures:200,
  pictureBase64:16*1024*1024
});

export class ProjectError extends Error{
  constructor(message,code='invalid'){super(message);this.name='ProjectError';this.code=code;}
}

const isObject=v=>v!==null&&typeof v==='object'&&!Array.isArray(v);
const text=(v,max,what)=>{
  if(typeof v!=='string')throw new ProjectError(`The project's ${what} is not text, so it was not opened.`);
  if(v.length>max)throw new ProjectError(`The project's ${what} is too long, so it was not opened.`);
  return v;
};
const orNull=(v,max,what)=>v===null||v===undefined?null:text(v,max,what);
// Settings the look wrote are plain data. Take a copy with only JSON values, and none of the keys that could reach an object's prototype.
const FORBIDDEN=new Set(['__proto__','constructor','prototype']);
const plain=(v,depth=0)=>{
  if(depth>20)throw new ProjectError('The project holds settings nested too deeply, so it was not opened.');
  if(v===null||['string','boolean'].includes(typeof v))return v;
  if(typeof v==='number'){if(!Number.isFinite(v))throw new ProjectError('The project holds a number that is not a number, so it was not opened.');return v;}
  if(Array.isArray(v))return v.map(x=>plain(x,depth+1));
  if(isObject(v)){
    const out={};
    for(const [k,x] of Object.entries(v)){if(FORBIDDEN.has(k))continue;out[k]=plain(x,depth+1);}
    return out;
  }
  throw new ProjectError('The project holds a setting of a kind this program does not know, so it was not opened.');
};
const list=(v,what)=>{
  if(v===undefined||v===null)return [];
  if(!Array.isArray(v)||v.length>LIMITS.looseItems)throw new ProjectError(`The project's ${what} are not in the right form, so it was not opened.`);
  return plain(v);
};

// The author's pictures: id -> {mime, base64}. The id is the one the paragraphs' XML names (r:embed="..."). Only the kinds of picture a Word
// document can hold safely, as base64 text of a sane size.
export const PICTURE_MIME=Object.freeze(['image/png','image/jpeg','image/gif']);
const PICTURE_ID=/^rIdPic[A-Za-z0-9_-]{1,60}$/;
const BASE64=/^[A-Za-z0-9+/]*={0,2}$/;
const cleanMedia=m=>{
  if(m===undefined||m===null)return {};
  if(!isObject(m))throw new ProjectError('The project\'s pictures are not in the right form, so it was not opened.');
  const entries=Object.entries(m);
  if(entries.length>LIMITS.pictures)throw new ProjectError('The project has more pictures than this program can hold, so it was not opened.');
  const out=Object.create(null);
  let total=0;
  for(const [id,v] of entries){
    if(!PICTURE_ID.test(id)||!isObject(v)||!PICTURE_MIME.includes(v.mime)||typeof v.base64!=='string'||!v.base64||!BASE64.test(v.base64)||v.base64.length%4!==0)
      throw new ProjectError('A picture in the project is not in the right form, so it was not opened.');
    if(v.base64.length>LIMITS.pictureBase64||(total+=v.base64.length)>LIMITS.fileCharacters)throw new ProjectError('A picture in the project is too large, so it was not opened.');
    out[id]={mime:v.mime,base64:v.base64};
  }
  return {...out};
};

// Checks the guide part of a project and returns a clean copy of it. Unknown fields are dropped; a missing optional field takes its empty value.
export function cleanGuide(g){
  if(!isObject(g))throw new ProjectError('The project does not hold a guide, so it was not opened.');
  if(!Array.isArray(g.paragraphs))throw new ProjectError('The project has no list of paragraphs, so it was not opened.');
  if(g.paragraphs.length>LIMITS.paragraphs)throw new ProjectError('The project has more paragraphs than this program can hold, so it was not opened.');
  const paragraphs=g.paragraphs.map((p,i)=>{
    if(!isObject(p))throw new ProjectError(`Paragraph ${i+1} of the project is not in the right form, so it was not opened.`);
    const out={xml:text(p.xml??'',LIMITS.paragraphXml,`paragraph ${i+1}`),text:text(p.text??'',LIMITS.words,`paragraph ${i+1}`),style:text(p.style??'Normal',LIMITS.styleName,`paragraph ${i+1} style`)};
    if(p.builtIn!==undefined&&p.builtIn!==null)out.builtIn=text(p.builtIn,LIMITS.styleName,`paragraph ${i+1} style`);
    return out;
  });
  const cursor=Number.isInteger(g.cursor)?g.cursor:-1;
  if(!Array.isArray(g.ownCues??[])||(g.ownCues??[]).length>LIMITS.cues)throw new ProjectError('The project\'s own cue names are not in the right form, so it was not opened.');
  return {
    title:text(g.title??'',LIMITS.title,'title'),
    paragraphs,
    cursor:Math.max(-1,Math.min(cursor,paragraphs.length-1)),
    ownCues:(g.ownCues??[]).map(c=>text(c,LIMITS.cueName,'own cue name')),
    look:orNull(g.look,LIMITS.words,'look'),
    slideMemory:orNull(g.slideMemory,LIMITS.words,'slide memory'),
    versionMemory:orNull(g.versionMemory,LIMITS.words,'delivery versions'),
    forms:orNull(g.forms,LIMITS.words,'form settings'),
    applied:list(g.applied,'look settings'),
    headerFooter:list(g.headerFooter,'header and footer settings'),
    media:cleanMedia(g.media)
  };
}

// The text of a project file for a guide (as the document's exportState gives it).
export function serializeProject(guide,{savedAt=new Date().toISOString()}={}){
  const clean=cleanGuide(guide);
  const out=JSON.stringify({format:PROJECT_FORMAT,version:PROJECT_VERSION,savedAt,guide:clean},null,1)+'\n';
  if(out.length>LIMITS.fileCharacters)throw new ProjectError('This guide is larger than a project file can be.','too-large');
  return out;
}

// Reads the text of a project file. Returns {version, savedAt, guide} or throws a ProjectError whose message is fit to show.
export function parseProject(fileText){
  if(typeof fileText!=='string')throw new ProjectError('That is not an OpenStage Guide project.');
  if(fileText.length>LIMITS.fileCharacters)throw new ProjectError('That file is larger than an OpenStage Guide project can be.','too-large');
  let data;
  try{data=JSON.parse(fileText);}catch{throw new ProjectError('That is not an OpenStage Guide project: the file is not readable in the form a project is saved in.');}
  if(!isObject(data)||data.format!==PROJECT_FORMAT)throw new ProjectError('That is not an OpenStage Guide project.');
  if(!Number.isInteger(data.version)||data.version<1)throw new ProjectError('That project has no usable version, so it was not opened.');
  if(data.version>PROJECT_VERSION)throw new ProjectError('That project was saved by a newer OpenStage Guide. Update the program, then open it again.','newer');
  return {version:data.version,savedAt:typeof data.savedAt==='string'?data.savedAt.slice(0,64):'',guide:cleanGuide(data.guide)};
}
