// The rating grids and self-scored instruments of a guide, sent to PrismStatistics as a .pstat file (format 1, the shared core's PSTAT-1 specification
// in github.com/jasonswrench-beep/openstage-core): one variable for each statement, in order, with the statement as its label and the scale's points as value
// labels, and as many empty rows as there are people to enter answers for. When the answers come back (typed from the paper forms, or pasted), the data is ready to analyse.
//
// What the file cannot hold, it says in plain sentences in `source.report.notes`, which PrismStatistics shows when it opens the file: how each scale is added up and what each
// range of totals means. There is no profile in the core for an instrument and its scoring guide yet, so there is no part to carry them as data; that is asked for in
// CORE-REQUESTS.md (request G1). This file follows format 1 only (no `profile`), the way any writer other than PrismStatistics may write it: `operations` and
// `results` are empty, and nothing is written that PrismStatistics owns. The writer is this file's own, as the core's notes advise for a tool that keeps no runtime
// dependency on the package; the tests check what it writes with the core's `checkPstat`.
import {readTable} from './guide-table.mjs';
import {formSpans} from './guide-forms.mjs';

export const PSTAT_FORMAT='PrismStatistics project';
export const PSTAT_VERSION=1;
export const MAX_ROWS=999;
export const MAX_STATEMENTS=200;
const clean=t=>String(t??'').replace(/[\r\n\u0007]+/g,' ').replace(/\s+/g,' ').trim();
const plural=(n,w)=>`${n} ${w}${n===1?'':'s'}`;

const isTable=e=>/^<w:tbl[\s>]/.test(String(e?.xml??''));

// A rating grid: the header row starts with "Statement", then one cell for each point of the scale (and a last cell for the evidence when the grid has that column).
const RATING_DESCRIPTION=/^Statements to rate on a scale from (.*) to (.*)\. Each statement has a check box under each point(, and a column to write the evidence in)?\.$/s;
// A self-scored instrument is three tables that share a caption: the statements, the adding up (caption + ": adding up") and, when it has them, what the scores mean.
const INSTRUMENT_DESCRIPTION='The statements to rate. Write your score for each one in the last column.';

/**
 * The scale of an instrument as its score column's header gives it: "Score (1 (never) to 5 (always))", "Score (1 to 5)" or "Score".
 * @returns {{text:string, low:number|null, high:number|null, lowLabel:string, highLabel:string}}
 */
export function scoreScale(headerText){
  const m=/^Score\s*\((.*)\)\s*$/s.exec(clean(headerText));
  const text=m?clean(m[1]):'';
  const r=/^(-?\d+)\s*(?:\(([^)]*)\))?\s*to\s*(-?\d+)\s*(?:\(([^)]*)\))?$/i.exec(text);
  if(!r)return {text,low:null,high:null,lowLabel:'',highLabel:''};
  return {text,low:Number(r[1]),high:Number(r[3]),lowLabel:clean(r[2]),highLabel:clean(r[4])};
}

/**
 * Every rating grid and self-scored instrument in the guide, in order.
 * @param {{xml?:string,text?:string,style?:string,marks?:object}[]} paras the guide's paragraphs, tables as entries with their XML
 * @param {{parse:Function,serialize:Function}} io
 * @returns {{kind:'rating'|'instrument', title:string, formTitle:string, index:number, statements:string[], marks:string[], evidence:boolean, scale:ReturnType<typeof scoreScale>, scales:{name:string,items:number[],divide:number|null}[], bands:{group:string,from:number,to:number,meaning:string}[]}[]}
 */
export function readInstruments(paras,io){
  const spans=formSpans(paras);
  const formAt=i=>spans.find(s=>i>=s.index&&i<=s.end)||null;
  const out=[],used=new Set();
  const read=i=>{try{return readTable(paras[i].xml,io);}catch{return null;}};
  for(let i=0;i<paras.length;i++){
    if(!isTable(paras[i])||used.has(i))continue;
    const t=read(i);
    if(!t||!t.rows.length||!t.rows[0].header)continue;
    const head=t.rows[0].cells.map(c=>clean(c.text)),form=formAt(i);
    const base={formTitle:form?clean(form.text):'',index:i};
    const rating=RATING_DESCRIPTION.exec(t.description||'');
    if(head[0]==='Statement'&&rating){
      const evidence=!!rating[3],marks=head.slice(1,evidence?-1:undefined);
      const statements=t.rows.slice(1).map(r=>clean(r.cells[0]?.text)).filter(Boolean);
      if(statements.length&&marks.length>=2)out.push({kind:'rating',title:clean(t.caption)||base.formTitle||'Rating grid',...base,statements,marks,evidence,scale:{text:'',low:null,high:null,lowLabel:'',highLabel:''},scales:[],bands:[]});
      used.add(i);
      continue;
    }
    if(t.description===INSTRUMENT_DESCRIPTION&&head[0]==='No.'&&head[1]==='Statement'){
      const caption=clean(t.caption);
      const statements=t.rows.slice(1).map(r=>clean(r.cells[1]?.text)).filter(Boolean);
      // The tables that go with it come right after it.
      // (a table's cells are entries of their own after it, so they are stepped over; the search stops at the next form's title)
      const following=name=>{
        for(let j=i+1;j<paras.length;j++){
          if(/^OSG Form Title/.test(String(paras[j].style||'')))return null;
          if(isTable(paras[j])){const u=read(j);if(u&&clean(u.caption)===`${caption}: ${name}`)return {j,u};}
        }
        return null;
      };
      const adding=following('adding up'),meaning=following('what the scores mean');
      const scales=[];
      if(adding){
        used.add(adding.j);
        const h=adding.u.rows[0].cells.map(c=>clean(c.text));
        const divide=h.includes('Divide by');
        for(const r of adding.u.rows.slice(1)){
          const cells=r.cells.map(c=>clean(c.text));
          const items=cells[1].split(/[,\s]+/).map(Number).filter(n=>Number.isInteger(n)&&n>0);
          if(cells[0]&&items.length)scales.push({name:cells[0],items,divide:divide?(Number(cells[3])||items.length):null});
        }
      }
      const bands=[];
      if(meaning){
        used.add(meaning.j);
        const h=meaning.u.rows[0].cells.map(c=>clean(c.text)),withScale=h[0]==='Scale';
        for(const r of meaning.u.rows.slice(1)){
          const cells=r.cells.map(c=>clean(c.text));
          const [lo,hi]=(withScale?cells[1]:cells[0]).split(/\s+to\s+/).map(Number);
          const text=withScale?cells[2]:cells[1];
          if(Number.isFinite(lo)&&text)bands.push({group:withScale&&cells[0]!=='Every scale'?cells[0]:'',from:lo,to:Number.isFinite(hi)?hi:lo,meaning:text});
        }
      }
      if(statements.length)out.push({kind:'instrument',title:caption||base.formTitle||'Self-scored instrument',...base,statements,marks:[],evidence:false,scale:scoreScale(head[2]??''),scales,bands});
      used.add(i);
    }
  }
  return out;
}

// ---- the dataset

const pad=(n,width)=>String(n).padStart(width,'0');
/** The scale of a rating grid as value labels: points named by words get 1, 2, 3...; points that are numbers are those numbers and need no labels. */
function ratingValues(marks){
  const nums=marks.map(m=>/^-?\d+(\.\d+)?$/.test(m)?Number(m):null);
  if(nums.every(n=>n!==null))return {labels:[],low:Math.min(...nums),high:Math.max(...nums)};
  return {labels:marks.map((m,i)=>({value:i+1,label:m})),low:1,high:marks.length};
}
function instrumentValues(scale){
  const labels=[];
  if(scale.low!==null&&scale.lowLabel)labels.push({value:scale.low,label:scale.lowLabel});
  if(scale.high!==null&&scale.highLabel&&scale.high!==scale.low)labels.push({value:scale.high,label:scale.highLabel});
  return {labels,low:scale.low,high:scale.high};
}
const numberText=n=>Number.isInteger(n)?String(n):String(Math.round(n*1e6)/1e6);

/**
 * The project.json of the file, for the chosen instruments.
 * @param {ReturnType<typeof readInstruments>} instruments @param {{guideTitle?:string, rows?:number, appVersion?:string}} [options]
 * `rows` is how many people there are to enter answers for (0 to MAX_ROWS); the first variable, `respondent`, numbers them.
 */
export function pstatProject(instruments,{guideTitle='',rows=0,appVersion='OpenStage Guide'}={}){
  if(!instruments.length)throw new Error('Choose at least one rating grid or self-scored instrument to send.');
  const n=Number(rows);
  if(!Number.isInteger(n)||n<0||n>MAX_ROWS)throw new Error(`Give the number of people as a whole number from 0 to ${MAX_ROWS.toLocaleString('en-US')}.`);
  const total=instruments.reduce((a,i)=>a+i.statements.length,0);
  if(total>MAX_STATEMENTS)throw new Error(`The file holds up to ${MAX_STATEMENTS} statements; these have ${total}. Send fewer of them.`);
  const many=instruments.length>1,variables=[],columns=[],notes=[];
  let nextId=1;
  const add=(v,col)=>{variables.push({id:nextId++,...v});columns.push(col);};
  add({name:'respondent',label:'Respondent (a number for each person who answers)',measure:'scale',value_labels:[],missing_values:[]},{type:'numeric',values:Array.from({length:n},(_,i)=>i+1)});
  notes.push(`Made from the guide${guideTitle?` "${clean(guideTitle)}"`:''} by OpenStage Guide. The answers are empty: type or paste them in, one row for each person (${n?plural(n,'row'):'no rows yet'}).`);
  instruments.forEach((ins,k)=>{
    const prefix=many?`f${k+1}_`:'',width=Math.max(2,String(ins.statements.length).length);
    const values=ins.kind==='rating'?ratingValues(ins.marks):instrumentValues(ins.scale);
    const names=ins.statements.map((_,i)=>`${prefix}q${pad(i+1,width)}`);
    const scaleOf=i=>ins.scales.find(s=>s.items.includes(i+1))?.name||'';
    ins.statements.forEach((s,i)=>{
      const label=`${many?`${ins.title}: `:''}${s}${scaleOf(i)?` (scale: ${scaleOf(i)})`:''}`;
      add({name:names[i],label,measure:'ordinal',value_labels:values.labels,missing_values:[]},{type:'numeric',values:Array.from({length:n},()=>null)});
    });
    const range=values.low!==null&&values.high!==null?`, answered ${numberText(values.low)} to ${numberText(values.high)}`:'';
    notes.push(`${ins.kind==='rating'?'Rating grid':'Self-scored instrument'} "${ins.title}"${ins.formTitle&&ins.formTitle!==ins.title?` (in ${ins.formTitle})`:''}: ${plural(ins.statements.length,'statement')} in ${names[0]}${names.length>1?` to ${names.at(-1)}`:''}${range}${ins.kind==='rating'&&ins.marks.length?`; the points are ${ins.marks.join(', ')}`:''}.`);
    if(ins.evidence)notes.push(`"${ins.title}" also has a column to write the evidence in; the words are not data and are not in the file.`);
    // The scoring guide, in words: how each scale is added up and what each range of totals means.
    for(const sc of ins.scales){
      const vars=sc.items.slice().sort((a,b)=>a-b).map(x=>names[x-1]).filter(Boolean);
      notes.push(`Scale "${sc.name}" of "${ins.title}": add ${vars.join(', ')}${sc.divide?`, then divide the total by ${sc.divide} for the average`:''}.`);
      for(const b of ins.bands.filter(b=>!b.group||b.group.toLowerCase()===sc.name.toLowerCase()))
        notes.push(`${sc.name}, a ${sc.divide?'score':'total'} of ${b.from===b.to?numberText(b.from):`${numberText(b.from)} to ${numberText(b.to)}`}: ${b.meaning}`);
    }
    for(const b of ins.bands.filter(b=>b.group&&!ins.scales.some(s=>s.name.toLowerCase()===b.group.toLowerCase())))
      notes.push(`${b.group}, a total of ${b.from===b.to?numberText(b.from):`${numberText(b.from)} to ${numberText(b.to)}`}: ${b.meaning}`);
  });
  return {
    format:PSTAT_FORMAT,version:PSTAT_VERSION,app_version:appVersion,
    dataset:{name:clean(guideTitle)||'Instruments',variables,columns,next_id:nextId},
    source:{file_name:`${clean(guideTitle)||'Guide'} (OpenStage Guide)`,report:{rows:n,variables:variables.length,renamed:[],short_rows:0,notes}},
    operations:[],results:[]
  };
}

/**
 * The bytes of the .pstat file: a zip with `project.json` only, deflated.
 * @param {object} project @param {new()=>any} Zip JSZip
 */
export async function pstatBytes(project,Zip){
  const zip=new Zip();
  zip.file('project.json',JSON.stringify(project),{date:new Date(Date.UTC(2026,0,1)),compression:'DEFLATE'});
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}

/** A file name for the file: the guide's name and "for PrismStatistics". @param {string} guideName */
export const pstatFileName=guideName=>`${String(guideName||'Guide').replace(/[\\/:*?"<>|\u0000-\u001f]/g,'-').replace(/\s+/g,' ').trim().replace(/[. ]+$/,'').slice(0,80)||'Guide'} - instruments.pstat`;
