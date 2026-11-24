// What the program remembers besides the guides themselves: the activities saved in the activity library, the look chosen for the workbook, and the
// agenda's start time. In a browser these were kept in the page's local storage, but the program's window starts its server on a port that is new
// every time, and local storage belongs to one address and port, so they were gone the next time the program started. In the program they are kept
// in one small file beside the guides (OpenStage Guide settings.json), which the program's own server reads and writes: the page asks for a setting by
// name and never sees a file name or a folder, and only the names below are ever kept (the activities, the session templates, the look, the start time and the zoom of the guide's page).
import {STORE_KEY as LIBRARY_KEY} from './activity-library.mjs';
import {TEMPLATES_KEY} from './session-text.mjs';
import {ZOOM_KEY} from './zoom.mjs';

export const AGENDA_START_KEY='osg.agendaStart.v1';
/** When the author last chose "Clear previous guides" on the welcome page: guides not changed since then are left off that list (the files are never touched). */
export const GUIDES_CLEARED_KEY='osg.guidesClearedAt.v1';
export const LOOK_KEY='openstage-guide.workbook-look';
/** The only settings the program keeps (the zoom of the window is kept by the server itself: src/zoom.mjs). */
export const SETTING_KEYS=Object.freeze([LIBRARY_KEY,TEMPLATES_KEY,AGENDA_START_KEY,LOOK_KEY,ZOOM_KEY,GUIDES_CLEARED_KEY]);
export const SETTINGS_FORMAT='openstage-guide-settings';
export const SETTINGS_FILE='OpenStage Guide settings.json';
export const SETTINGS_PATH='/api/settings';
/** The most one setting may hold, in characters (50 activities of 4,000 letters a part come to well under this), and the most the file may hold. */
export const SETTING_MAX_CHARS=1024*1024;
export const SETTINGS_MAX_CHARS=2*1024*1024;

/**
 * The time a setting holds, as the text the program writes (an ISO time), or null when it is empty, not a time, or a day or more in the future.
 * @param {unknown} value
 * @returns {string | null}
 */
export function cleanClearedAt(value){
  if(typeof value!=='string'||value.length>40)return null;
  const t=Date.parse(value);
  if(!Number.isFinite(t)||t>Date.now()+24*3600*1000)return null;
  return new Date(t).toISOString();
}

/**
 * The guides the welcome page lists: the ones changed after the list was last cleared (or all of them, if it never was), and how many are left off.
 * A guide that is changed or made after the clearing is listed again by itself. The Open guide window always lists every guide.
 * @param {{ name: string, when: string }[]} guides @param {string | null} clearedAt
 * @returns {{ shown: { name: string, when: string }[], hidden: number }}
 */
export function splitGuides(guides,clearedAt){
  if(!clearedAt)return {shown:guides,hidden:0};
  const shown=guides.filter(g=>String(g.when)>clearedAt);
  return {shown,hidden:guides.length-shown.length};
}

/**
 * The settings in the text of the file: only the known names, only text values, and nothing at all when the file is not readable (a
 * file that is not ours is the same as no settings).
 * @param {string} text
 * @returns {Record<string, string>}
 */
export function parseSettings(text){
  /** @type {Record<string, string>} */
  const out={};
  let data;
  try{data=JSON.parse(String(text??''));}catch{return out;}
  if(!data||typeof data!=='object'||data.format!==SETTINGS_FORMAT||!data.values||typeof data.values!=='object'||Array.isArray(data.values))return out;
  let total=0;
  for(const key of SETTING_KEYS){
    const value=Object.hasOwn(data.values,key)?data.values[key]:undefined;
    if(typeof value!=='string'||value.length>SETTING_MAX_CHARS||(total+=value.length)>SETTINGS_MAX_CHARS)continue;
    out[key]=value;
  }
  return out;
}

/** The text of the file for these settings. @param {Record<string, string>} values */
export function serializeSettings(values){
  /** @type {Record<string, string>} */
  const clean={};
  for(const key of SETTING_KEYS)if(typeof values?.[key]==='string')clean[key]=values[key];
  return JSON.stringify({format:SETTINGS_FORMAT,version:1,values:clean},null,1)+'\n';
}

const browserStorage=()=>{try{return globalThis.localStorage||null;}catch{return null;}};

/**
 * The page's side: storage with the same three calls as the browser's local storage (getItem, setItem, removeItem), kept by the program's server.
 * The calls are answered at once (a small request to this program, on this computer), because the pane asks for a setting while it is being drawn.
 * A setting the server does not have yet, but this browser's own local storage does (kept there before the program had a file), is carried over the
 * first time it is asked for. A refusal is thrown, as the browser's storage does when it is full, so the pane says what it could not keep.
 * @param {{ xhr?: () => XMLHttpRequest, legacy?: Storage | null }} [options]
 */
export function createProgramStorage({xhr=()=>new XMLHttpRequest(),legacy=browserStorage()}={}){
  /** @param {'GET' | 'POST'} method @param {string} [body] */
  function call(method,body){
    const request=xhr();
    request.open(method,SETTINGS_PATH,false);
    if(body!==undefined)request.setRequestHeader('Content-Type','application/json');
    request.send(body);
    if(request.status<200||request.status>=300){
      let why='';
      try{why=String(JSON.parse(request.responseText).error||'');}catch{why='';}
      throw new Error(why||`The program answered ${request.status}.`);
    }
    return request.responseText;
  }
  const values=()=>{
    try{const v=JSON.parse(call('GET')).values;return v&&typeof v==='object'?v:{};}catch{return {};}
  };
  return {
    /** @param {string} key */
    getItem(key){
      const have=values()[key];
      if(typeof have==='string')return have;
      let old=null;
      try{old=legacy?.getItem(key)??null;}catch{old=null;}
      if(old!==null){try{call('POST',JSON.stringify({key,value:old}));}catch{/* it is still given back, and tried again next time */}}
      return old;
    },
    /** @param {string} key @param {unknown} value */
    setItem(key,value){call('POST',JSON.stringify({key,value:String(value)}));},
    /** @param {string} key */
    removeItem(key){
      call('POST',JSON.stringify({key,value:null}));
      try{legacy?.removeItem(key);}catch{/* the old copy is only a way in */}
    },
  };
}
