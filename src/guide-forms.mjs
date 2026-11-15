// Forms and handouts: the part of the guide model that knows what a form is, how forms are numbered, how the script cites them, how they
// are printed, and what is wrong with them. Pure functions over the guide's paragraphs, like the rest of guide-model.mjs: nothing here
// touches a document. The design is in docs/FORMS-AND-HANDOUTS-SPEC-2026-11-13.md.
//
// A form is a paragraph in the style "OSG Form Title" (or "OSG Form Title Facilitator", for a form the participant workbook leaves out)
// that reads "Handout 5.4: Standards in depth" (a label, a number, a colon, a title) and carries a bookmark named osg_form_<id>. The id
// never changes when the form is renumbered or renamed. A citation in the script is a content control tagged osg-ref:<id> that holds the
// form's name as words. "Update guide" rewrites numbers and citations from the ids, the way it already rewrites heading numbers and time
// lines. Nothing is hidden: the bookmark and the tag are ordinary Word-format constructs that a project file or a Word export keeps.
//
// A form may have a print line directly under its title, in the style "OSG Form Print": "Print: 1 per person, Letter, landscape." It is for
// the facilitator and never reaches the workbook. With the number of people, teams or tables the guide records, the print lines add up to
// a list of copies to make.
//
// Each paragraph the adapters report may carry `marks`: {bookmarks:[name], refs:[{id,text}]}, read from the paragraph's own XML.
// A document that reports no marks (Word through the add-in) simply has no forms.
import {STYLE_NAMES} from './guide-styles.mjs';

export const BOOKMARK_PREFIX='osg_form_';
export const REF_TAG_PREFIX='osg-ref:';
export const SCHEMES=Object.freeze(['auto','unit','letter','sequence','none']);
// The words the Forms tool will offer in front of a number. Any other label an author types is read and kept.
export const LABELS=Object.freeze(['Handout','Worksheet','Checklist','Form','Template','Cards']);
export const DEFAULT_LABEL='Handout';
const ID=/^[a-z][a-z0-9]{0,11}$/;
export const isFormId=id=>ID.test(String(id));
const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').trim();

// ---- the guide's settings for forms (kept with the document, like the look and the slide memory)
// scheme: how forms are numbered. group: how many people, teams and tables there will be, so that print lines can be added up.
const count=v=>Number.isInteger(v)&&v>=1&&v<=9999?v:null;
// The stated length of the programme, in minutes (kit slice K2). The settings are kept with the guide under the name the form settings always had.
const length=v=>Number.isInteger(v)&&v>=1&&v<=99999?v:null;
export function readSettings(text){
  let d=null;
  try{d=text?JSON.parse(text):null;}catch{d=null;}
  const group={people:count(d?.group?.people),teams:count(d?.group?.teams),tables:count(d?.group?.tables)};
  const minutes=length(d?.time?.minutes);
  return {scheme:SCHEMES.includes(d?.scheme)?d.scheme:'auto',...(Object.values(group).some(v=>v!==null)?{group}:{}),...(minutes!==null?{time:{minutes}}:{})};
}
export const writeSettings=s=>JSON.stringify({scheme:SCHEMES.includes(s?.scheme)?s.scheme:'auto',...(s?.group?{group:{people:count(s.group.people),teams:count(s.group.teams),tables:count(s.group.tables)}}:{}),...(length(s?.time?.minutes)!==null?{time:{minutes:length(s.time.minutes)}}:{})});

// ---- the words of a form's title
const TITLE=/^([A-Za-z][A-Za-z ]{0,30}?)(?:\s+(\d+(?:\.\d+)?|[A-Za-z]{1,2}))?\s*:\s*(\S.*)$/s;
// "Handout 5.4: Standards in depth" -> {label:'Handout',number:'5.4',title:'Standards in depth'}; null when it does not read that way.
export function parseTitle(text){
  const m=TITLE.exec(clean(text));
  return m?{label:m[1].trim(),number:m[2]||'',title:m[3].trim()}:null;
}
export const titleText=({label=DEFAULT_LABEL,number='',title=''})=>`${label}${number?` ${number}`:''}: ${title}`;

// ---- the print line: "Print: 1 per person, Letter, landscape."
const PAPERS=['Letter','A4'],ORIENTATIONS=['portrait','landscape'],SIDES=['one-sided','two-sided'];
const PER=['person','team','table','room'];
// Reads the words after "Print:" into {copies:{n,per,size}|null, paper, orientation, sides, editable, notes}. null when the line does not start with "Print:".
export function parsePrint(text){
  const m=/^Print:\s*(.*?)\.?\s*$/i.exec(clean(text));
  if(!m)return null;
  const spec={copies:null,paper:null,orientation:null,sides:null,editable:false,notes:[]};
  for(const token of m[1].split(',').map(t=>t.trim()).filter(Boolean)){
    const lower=token.toLowerCase();
    let x;
    if((x=/^(\d{1,4}) per (person|room|table|team)(?: of (\d{1,3}))?$/i.exec(token)))spec.copies={n:Number(x[1]),per:x[2].toLowerCase(),size:x[3]?Number(x[3]):null};
    else if((x=/^(\d{1,5}) cop(?:y|ies)$/i.exec(token)))spec.copies={n:Number(x[1]),per:'fixed',size:null};
    else if(PAPERS.some(p=>p.toLowerCase()===lower))spec.paper=PAPERS.find(p=>p.toLowerCase()===lower);
    else if(ORIENTATIONS.includes(lower))spec.orientation=lower;
    else if(SIDES.includes(lower))spec.sides=lower;
    else if(/^editable( download)?$/i.test(token))spec.editable=true;
    else spec.notes.push(token);
  }
  return spec;
}
// The canonical words for a print line. Throws on something that cannot be printed.
export function printText(spec){
  const parts=[];
  const c=spec?.copies;
  if(c){
    const n=c.n;
    if(!Number.isInteger(n)||n<1||n>(c.per==='fixed'?99999:9999))throw new Error('Copies must be a whole number from 1 up.');
    if(c.per==='fixed')parts.push(`${n} ${n===1?'copy':'copies'}`);
    else{
      if(!PER.includes(c.per))throw new Error('Copies are per person, team, table or room.');
      parts.push(`${n} per ${c.per}${c.size&&['team','table'].includes(c.per)?` of ${c.size}`:''}`);
    }
  }
  if(spec?.paper){if(!PAPERS.includes(spec.paper))throw new Error('Paper is Letter or A4.');parts.push(spec.paper);}
  if(spec?.orientation){if(!ORIENTATIONS.includes(spec.orientation))throw new Error('A page is portrait or landscape.');parts.push(spec.orientation);}
  if(spec?.sides){if(!SIDES.includes(spec.sides))throw new Error('A page is one-sided or two-sided.');parts.push(spec.sides);}
  if(spec?.editable)parts.push('editable download');
  for(const note of spec?.notes||[])parts.push(String(note).replace(/[,\r\n]+/g,' ').trim());
  if(!parts.length)throw new Error('Say how many copies, or the paper, or the page.');
  return `Print: ${parts.join(', ')}.`;
}

// ---- numbers
// A, B, ... Z, AA, AB, ... for the 0-based position i.
export function lettersFor(i){
  let n=i+1,s='';
  while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}
  return s;
}
// "auto" means by unit when the guide has modules and letters when it has none.
export const resolveScheme=(scheme,hasModules)=>!scheme||scheme==='auto'?(hasModules?'unit':'letter'):scheme;

// The number each form should have, in the order the forms are given. By unit a form takes its module's number and its place in the
// module (5.4); a form outside any module (the front pages, the Appendix) takes a letter instead.
export function numberForms(forms,scheme,hasModules){
  const mode=resolveScheme(scheme,hasModules),perModule=new Map();
  let seq=0,loose=0;
  return forms.map(f=>{
    if(mode==='none')return '';
    if(mode==='sequence')return String(++seq);
    if(mode==='letter')return lettersFor(seq++);
    if(f.module>0){const k=(perModule.get(f.module)||0)+1;perModule.set(f.module,k);return `${f.module}.${k}`;}
    return lettersFor(loose++);
  });
}

// ---- reading
const FORM_TEXT_STYLES=new Set([STYLE_NAMES.formText,STYLE_NAMES.formLine,STYLE_NAMES.formField,STYLE_NAMES.formBox,STYLE_NAMES.formSource,STYLE_NAMES.formCard]);
const TABLE_STYLES=new Set([STYLE_NAMES.tableHead,STYLE_NAMES.tableCell,STYLE_NAMES.ratingBox]);
const TITLE_STYLES=new Set([STYLE_NAMES.formTitle,STYLE_NAMES.formTitleFacilitator]);
// What ends a form: a module, lesson, assessment or heading, the guide title, a block of the script, or the next form.
const endsForm=style=>[STYLE_NAMES.module,STYLE_NAMES.lesson,STYLE_NAMES.assessment,STYLE_NAMES.guideTitle,STYLE_NAMES.cue,STYLE_NAMES.cueFacilitator,STYLE_NAMES.formTitle,STYLE_NAMES.formTitleFacilitator].includes(style)||/^(heading \d|title)$/i.test(style);

// Where the forms and the citations are. A form is any paragraph in a form title style; its id is the one in its bookmark.
export function readForms(paras){
  const forms=[],cites=[],byId=new Map(),duplicates=[],noId=[];
  let module=0,appendix=false,hasModules=false,current=null,table=null,blockAudience=null;
  const closeTable=()=>{if(table&&current&&table.cells>0&&!table.head)current.headerless.push(table.first);table=null;};
  paras.forEach((p,index)=>{
    const style=String(p.style||'');
    if(style===STYLE_NAMES.module){module++;hasModules=true;appendix=false;}
    else if(/^heading [12]$/i.test(style)&&/^appendix\b/i.test(clean(p.text)))appendix=true;
    // The script block a paragraph is in decides who a citation in it is told to.
    if(style===STYLE_NAMES.cue)blockAudience='everyone';
    else if(style===STYLE_NAMES.cueFacilitator)blockAudience='facilitator';
    else if(endsForm(style))blockAudience=null;
    const isTitle=TITLE_STYLES.has(style);
    if(endsForm(style)){closeTable();if(!isTitle)current=null;}
    const marks=p.marks||{};
    if(isTitle){
      const id=(marks.bookmarks||[]).map(b=>String(b).startsWith(BOOKMARK_PREFIX)?String(b).slice(BOOKMARK_PREFIX.length):null).find(x=>x&&ID.test(x))||null;
      const parsed=parseTitle(p.text);
      const f={index,id,audience:style===STYLE_NAMES.formTitleFacilitator?'facilitator':'everyone',module:appendix?0:module,appendix,text:clean(p.text),numbered:!!parsed,
        label:parsed?.label||DEFAULT_LABEL,number:parsed?.number||'',title:parsed?.title||clean(p.text),print:null,printIndex:-1,printUnreadable:false,
        parts:0,emptyFields:[],headerless:[]};
      const next=paras[index+1];
      if(next&&String(next.style||'')===STYLE_NAMES.formPrint){
        const spec=parsePrint(next.text);
        f.printIndex=index+1;
        if(spec)f.print=spec;else f.printUnreadable=true;
      }
      forms.push(f);current=f;
      if(!id)noId.push(f);
      else if(byId.has(id))duplicates.push(f);
      else byId.set(id,f);
    }else if(current){
      if(FORM_TEXT_STYLES.has(style)){
        closeTable();current.parts++;
        if(style===STYLE_NAMES.formField&&!clean(p.text))current.emptyFields.push(index);
      }else if(TABLE_STYLES.has(style)){
        current.parts++;
        if(!table)table={head:false,cells:0,first:index};
        if(style===STYLE_NAMES.tableHead)table.head=true;else table.cells++;
      }else if(style!==STYLE_NAMES.formPrint)closeTable();
    }
    for(const r of marks.refs||[])cites.push({index,id:String(r.id),text:clean(r.text),audience:blockAudience});
  });
  closeTable();
  return {forms,cites,byId,duplicates,noId,hasModules};
}

// Where each form is: from its title to the paragraph before whatever ends it (the same rule readForms reads by). `orientation` is what its print line says,
// or null. The outputs read this to give a landscape form its own page setup and to pick the forms of a lesson, a module or the guide.
export function formSpans(paras){
  const {forms}=readForms(paras);
  return forms.map(f=>{
    let end=f.index;
    while(end+1<paras.length&&!endsForm(String(paras[end+1].style||'')))end++;
    return {index:f.index,end,id:f.id,audience:f.audience,text:f.text,numbered:f.numbered,label:f.label,number:f.number,orientation:f.print?.orientation||null,printIndex:f.printIndex};
  });
}
// For each paragraph, whether it is in a form that is printed landscape. Used by the Word outputs to put such a form in a section of its own.
export function landscapeFlags(paras){
  const flags=paras.map(()=>false);
  for(const s of formSpans(paras))if(s.orientation==='landscape')for(let i=s.index;i<=s.end;i++)flags[i]=true;
  return flags;
}

// What "Update guide" would change for forms, and what to tell the author. `settings` is readSettings' result.
export function formPlan(paras,settings={}){
  const read=readForms(paras),numbers=numberForms(read.forms,settings.scheme||'auto',read.hasModules);
  const edits=[],issues=[];
  const group=settings.group||{};
  const citedIds=new Set(read.cites.map(c=>c.id));
  read.forms.forEach((f,i)=>{
    f.want=numbers[i];
    f.fullText=titleText({label:f.label,number:f.want,title:f.title});
    if(f.numbered&&f.number!==f.want)edits.push({kind:'form',index:f.index,text:f.fullText});
    if(!clean(f.text))issues.push({level:'fail',rule:'FORM-EMPTY-TITLE',message:'A form has no title. Type its title, such as "Handout: Goals".',paragraph:f.index});
    else if(!f.numbered)issues.push({level:'warn',rule:'FORM-NAME',message:`"${f.text}" is not written as "Handout N: Title", so it is not renumbered.`,paragraph:f.index});
    if(!f.id)issues.push({level:'warn',rule:'FORM-NO-ID',message:`"${f.text}" has no identity, so the script cannot cite it and it is not renumbered in citations.`,paragraph:f.index});
    if(clean(f.text)&&f.parts===0)issues.push({level:'warn',rule:'FORM-EMPTY',message:`${f.text} has nothing in it yet.`,paragraph:f.index});
    for(const at of f.emptyFields)issues.push({level:'warn',rule:'FORM-FIELD-LABEL',message:`A field in ${f.text} has no label, so no one knows what to write on its line.`,paragraph:at});
    for(const at of f.headerless)issues.push({level:'fail',rule:'FORM-TABLE-HEADER',message:`A table in ${f.text} has no header row, so a screen reader cannot say what its columns hold.`,paragraph:at});
    if(f.printUnreadable)issues.push({level:'warn',rule:'FORM-PRINT-NAME',message:`The print line of ${f.text} does not read as "Print: 1 per person, Letter, landscape."`,paragraph:f.printIndex});
    const c=f.print?.copies;
    if(c&&c.per==='person'&&!group.people)issues.push({level:'warn',rule:'FORM-PRINT-GROUP',message:`${f.text} is printed for each person, but the guide does not say how many people there will be. Say how many in Forms, in the Activities menu.`,paragraph:f.index});
    if(c&&c.per==='team'&&!group.teams&&!(c.size&&group.people))issues.push({level:'warn',rule:'FORM-PRINT-GROUP',message:`${f.text} is printed for each team, but the guide does not say how many teams there will be. Say how many in Forms, in the Activities menu.`,paragraph:f.index});
    if(c&&c.per==='table'&&!group.tables)issues.push({level:'warn',rule:'FORM-PRINT-GROUP',message:`${f.text} is printed for each table, but the guide does not say how many tables there will be. Say how many in Forms, in the Activities menu.`,paragraph:f.index});
    if(f.id&&f.audience==='everyone'&&citedIds.has(f.id)&&!f.print&&!f.printUnreadable)issues.push({level:'warn',rule:'FORM-PRINT-NONE',message:`${f.fullText} is given out but has no print line (how many copies, which paper).`,paragraph:f.index});
  });
  for(const f of read.duplicates)issues.push({level:'warn',rule:'FORM-DUPLICATE',message:`"${f.text}" shares its identity with an earlier form, so citations point to the earlier one.`,paragraph:f.index});
  // Several citations of one form in one paragraph are told apart by their order inside it.
  const seen=new Map(),cited=new Set(),missing=new Set(),told=new Set();
  for(const c of read.cites){
    const key=`${c.index}|${c.id}`,nth=seen.get(key)||0;
    seen.set(key,nth+1);
    const f=read.byId.get(c.id);
    if(!f){
      if(!missing.has(key)){missing.add(key);issues.push({level:'fail',rule:'FORM-MISSING',message:'A handout or form is cited here, but it is no longer in the guide.',paragraph:c.index});}
      continue;
    }
    cited.add(c.id);
    if(c.text!==f.fullText)edits.push({kind:'citation',index:c.index,id:c.id,nth,from:c.text,to:f.fullText});
    if(f.audience==='facilitator'&&c.audience==='everyone'&&!told.has(key)){told.add(key);issues.push({level:'warn',rule:'FORM-AUDIENCE',message:`${f.fullText} is for the facilitator only, but a block for everyone cites it. Learners will not have it.`,paragraph:c.index});}
  }
  // A guide that cites forms at all should cite each of them; a guide that lists its forms without citing any is left alone.
  if(read.cites.length)for(const f of read.byId.values())if(!cited.has(f.id))issues.push({level:'warn',rule:'FORM-UNCITED',message:`${f.fullText} is not cited from the script.`,paragraph:f.index});
  return {read,numbers,edits,issues,scheme:resolveScheme(settings.scheme,read.hasModules)};
}

// ---- copies to make
// Adds up the print lines. A line says how many copies for each unit (person, team, table, room); the guide's group says how many units.
// Returns {rows:[{index,id,text,module,per,n,units,total}], total, missing:[{index,text,reason}]}: a form whose units are unknown has no
// total and is listed in `missing`.
export function copiesToMake(paras,settings={}){
  const {forms}=readForms(paras),g=settings.group||{};
  const rows=[],missing=[];
  for(const f of forms){
    const c=f.print?.copies;
    if(!c)continue;
    let units=null,why='';
    if(c.per==='person'){units=g.people??null;why='how many people';}
    else if(c.per==='room'||c.per==='fixed')units=1;
    else if(c.per==='table'){units=g.tables??null;why='how many tables';}
    else if(c.per==='team'){units=g.teams??(c.size&&g.people?Math.ceil(g.people/c.size):null);why='how many teams';}
    const total=units===null?null:c.n*units;
    rows.push({index:f.index,id:f.id,text:f.text,module:f.module,per:c.per,n:c.n,units,total});
    if(units===null)missing.push({index:f.index,text:f.text,reason:`The guide does not say ${why}.`});
  }
  return {rows,total:rows.reduce((a,r)=>a+(r.total||0),0),missing};
}

// A new id that no form and no citation in the guide uses: f1, f2, ... (a citation whose form was deleted keeps its id, so the id is not reused).
export function newFormId(paras){
  const read=readForms(paras),used=new Set([...read.forms.map(f=>f.id).filter(Boolean),...read.cites.map(c=>c.id)]);
  for(let n=1;;n++)if(!used.has(`f${n}`))return `f${n}`;
}
