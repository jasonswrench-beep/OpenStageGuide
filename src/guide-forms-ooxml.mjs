// The Word-format pieces of forms: a form's title, a citation of a form in the script, and the elements a form is made of (spec slice F2).
// Pure functions, text in and XML out, like the rest of guide-ooxml.mjs. What forms mean is in guide-forms.mjs.
//
// Rules the elements follow, because a form is for pens on paper and for screen readers:
//  * A table is used only where the content is a table. It has a header row that Word knows is a header row, a caption and a
//    description, and no merged cells. A form never uses a table to place things on the page.
//  * A line to write on is a paragraph with a bottom border, a field's line is a tab leader, and a box to write in is a paragraph with a
//    border and a set height. Nothing is made of typed underscores.
//  * A check box is a real check box content control with a name that says what it is for.
import {para,run,lines,esc} from './guide-ooxml.mjs';
import {STYLE_IDS} from './guide-styles.mjs';
import {table,row,cell,head,checkBox,PAGE} from './guide-assess-ooxml.mjs';
import {BOOKMARK_PREFIX,REF_TAG_PREFIX,isFormId,titleText,printText} from './guide-forms.mjs';

export const FORM_LIMITS=Object.freeze({lines:12,boxLines:12,fields:4,columns:8,rows:40,items:60,statements:30,label:200,cell:300,extraColumns:3,cards:60,cardWords:300,groups:8,bands:24});

const need=id=>{if(!isFormId(id))throw new Error('A form needs an identity: a few lower-case letters and digits, such as f1.');return String(id);};
// A bookmark's number only has to differ from the others in the document; the id's own digits are a stable choice.
const bookmarkNumber=id=>9000+(parseInt(String(id).replace(/\D+/g,''),10)||0);
const clip=(text,max,what)=>{
  const s=String(text??'').replace(/\s+/g,' ').trim();
  if(s.length>max)throw new Error(`${what} is longer than ${max} letters.`);
  return s;
};
const whole=(n,min,max,what)=>{
  if(!Number.isInteger(n)||n<min||n>max)throw new Error(`${what} must be a whole number from ${min} to ${max}.`);
  return n;
};
const short=(text,n=40)=>{const t=String(text).replace(/\s+/g,' ').trim();return t.length>n?`${t.slice(0,n-1)}…`:t;};
const emptyCell=width=>cell(width,para(STYLE_IDS.tableCell,''));
const split=(total,weights)=>{
  const sum=weights.reduce((a,b)=>a+b,0),w=weights.map(x=>Math.floor(total*x/sum));
  w[w.length-1]+=total-w.reduce((a,b)=>a+b,0);
  return w;
};

// ---- the title and the citation (slice F1)
// The title paragraph of a form: "Handout 5.4: Standards in depth" in the form title style, carrying the bookmark that names the form.
// `facilitator` makes it a form for the facilitator only, which the participant workbook leaves out.
export function formTitleBlock({id,label,number,title,facilitator=false}){
  const n=bookmarkNumber(need(id));
  const words=clip(title,FORM_LIMITS.label,'The title');
  if(!words)throw new Error('Type a title for the form.');
  return para(facilitator?STYLE_IDS.formTitleFacilitator:STYLE_IDS.formTitle,`<w:bookmarkStart w:id="${n}" w:name="${BOOKMARK_PREFIX}${id}"/>${run(titleText({label,number,title:words}))}<w:bookmarkEnd w:id="${n}"/>`);
}

// A citation of a form, to sit inside a paragraph of the script: a content control tagged with the form's id that holds the form's
// name as words. `ctx` is the insert's context (guide-ooxml.mjs makeContext), which hands out the control's number.
export function citationControl({id,text},ctx){
  need(id);
  return `<w:sdt><w:sdtPr><w:alias w:val="${esc('Cited handout or form')}"/><w:tag w:val="${REF_TAG_PREFIX}${id}"/><w:id w:val="${800000+(++ctx.sdt)}"/></w:sdtPr><w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`;
}

// The print line that goes directly under a form's title: how many copies, which paper, which page. It is for the facilitator: the
// participant workbook leaves it out. spec: {copies:{n,per,size},paper,orientation,sides,editable,notes} (see guide-forms.mjs parsePrint).
export const printLine=spec=>para(STYLE_IDS.formPrint,run(printText(spec)));

// ---- text, lines and boxes
// Reading or instruction text: one paragraph for each line typed.
export function formText(text){
  const ls=lines(text);
  if(!ls.length)throw new Error('Type the words first.');
  return ls.map(l=>para(STYLE_IDS.formText,run(l))).join('');
}
// A label (optional) and 1 to 12 ruled lines to write on.
export function writingLines({label='',count=3}={}){
  whole(count,1,FORM_LIMITS.lines,'Lines to write on');
  const words=clip(label,FORM_LIMITS.label,'The label');
  return (words?para(STYLE_IDS.formText,run(words)):'')+Array.from({length:count},()=>para(STYLE_IDS.formLine,'')).join('');
}
// A label (optional) and an empty box about `lines` lines tall (1 to 12), for a drawing or a longer answer.
export function writingBox({label='',lines:height=4}={}){
  whole(height,1,FORM_LIMITS.boxLines,'The height of the box, in lines,');
  const words=clip(label,FORM_LIMITS.label,'The label');
  return (words?para(STYLE_IDS.formText,run(words)):'')+para(STYLE_IDS.formBox,'',`<w:spacing w:before="120" w:after="120" w:line="${height*360}" w:lineRule="exact"/>`);
}
// One to four labelled fields on a line (Name, Date, Signature), each followed by a line to write on. width: short, medium or full.
const FIELD_WEIGHT={short:1,medium:2,full:4};
const GAP=240;
export function fieldRow(fields){
  if(!Array.isArray(fields)||!fields.length||fields.length>FORM_LIMITS.fields)throw new Error(`A row of fields holds 1 to ${FORM_LIMITS.fields} fields.`);
  const named=fields.map(f=>({label:clip(f.label,FORM_LIMITS.label,'A field label'),width:FIELD_WEIGHT[f.width]?f.width:'medium'}));
  if(named.some(f=>!f.label))throw new Error('Every field needs a label.');
  const widths=split(PAGE-GAP*(named.length-1),named.map(f=>FIELD_WEIGHT[f.width]));
  let at=0;const stops=[];let body='';
  named.forEach((f,i)=>{
    at+=widths[i];
    stops.push(`<w:tab w:val="right" w:leader="underscore" w:pos="${at}"/>`);
    body+=run(f.label)+'<w:r><w:tab/></w:r>';
    if(i<named.length-1){at+=GAP;stops.push(`<w:tab w:val="left" w:pos="${at}"/>`);body+='<w:r><w:tab/></w:r>';}
  });
  return para(STYLE_IDS.formField,body,`<w:tabs>${stops.join('')}</w:tabs>`);
}
// A short credit or permission line under a form.
export function sourceLine(text){
  const words=clip(text,FORM_LIMITS.cell*2,'The credit line');
  if(!words)throw new Error('Type the credit line first.');
  return para(STYLE_IDS.formSource,run(words));
}

// ---- tables
// A checklist: a check box, the item, and up to three more columns to fill in (Who, When, Notes, or words of your own).
export function checkList({items,columns=[],ctx,caption='Checklist'}){
  const list=(items||[]).map(i=>clip(i,FORM_LIMITS.cell,'An item')).filter(Boolean);
  if(!list.length)throw new Error('Add at least one item to the checklist.');
  if(list.length>FORM_LIMITS.items)throw new Error(`A checklist holds up to ${FORM_LIMITS.items} items.`);
  const extra=(columns||[]).map(c=>clip(c,60,'A column name')).filter(Boolean);
  if(extra.length>FORM_LIMITS.extraColumns)throw new Error(`A checklist holds up to ${FORM_LIMITS.extraColumns} extra columns.`);
  const boxW=700,extraW=extra.map(c=>/^notes?$/i.test(c)?2400:1700),itemW=PAGE-boxW-extraW.reduce((a,b)=>a+b,0);
  const widths=[boxW,itemW,...extraW];
  const rows=[row([head('Done',boxW,true),head('Item',itemW),...extra.map((c,i)=>head(c,extraW[i]))],{header:true}),
    ...list.map(item=>row([cell(boxW,para(STYLE_IDS.ratingBox,checkBox(ctx,`${short(item)}: done`))),cell(itemW,para(STYLE_IDS.tableCell,run(item))),...extraW.map(emptyCell)]))];
  return table(widths,rows,{caption:clip(caption,60,'The caption'),description:`A checklist. Each row has a check box to mark when the item is done, the item${extra.length?`, and space for ${extra.join(', ')}`:''}.`});
}

// A table to fill in. columns: [{name,weight,kind}] where kind is text (words that are already there), blank (space to write),
// number (space for a number) or check (a check box). rows: lists of words for the text columns. blankRows: empty rows to write in.
// sample: one row of example words, shown as the first row and marked "Example". totals: {label,columns:[positions]} adds a row to add up.
export function formTable({columns,rows=[],blankRows=0,sample=null,totals=null,ctx,caption='Table'}){
  if(!Array.isArray(columns)||!columns.length||columns.length>FORM_LIMITS.columns)throw new Error(`A table holds 1 to ${FORM_LIMITS.columns} columns.`);
  const cols=columns.map(c=>({name:clip(c.name,60,'A column name'),weight:Number.isFinite(c.weight)&&c.weight>0?c.weight:1,kind:['text','blank','number','check'].includes(c.kind)?c.kind:'text'}));
  if(cols.some(c=>!c.name))throw new Error('Every column needs a name, so a screen reader can say what it holds.');
  whole(blankRows,0,FORM_LIMITS.rows,'Empty rows');
  if(rows.length+blankRows+(sample?1:0)+(totals?1:0)>FORM_LIMITS.rows)throw new Error(`A table holds up to ${FORM_LIMITS.rows} rows.`);
  const widths=split(PAGE,cols.map(c=>c.weight));
  const inner=(c,i,text,r)=>{
    if(c.kind==='check')return cell(widths[i],para(STYLE_IDS.ratingBox,checkBox(ctx,`${c.name}: row ${r}`)));
    if(text!==undefined&&text!==null&&String(text).trim())return cell(widths[i],para(STYLE_IDS.tableCell,run(clip(text,FORM_LIMITS.cell,'A cell'))));
    return emptyCell(widths[i]);
  };
  const out=[row(cols.map((c,i)=>head(c.name,widths[i])),{header:true})];
  let r=0;
  if(sample){
    r++;
    out.push(row(cols.map((c,i)=>cell(widths[i],para(STYLE_IDS.tableCell,run(i===0?`Example: ${clip(sample[i]??'',FORM_LIMITS.cell,'A cell')}`:clip(sample[i]??'',FORM_LIMITS.cell,'A cell')))))));
  }
  for(const data of rows){r++;out.push(row(cols.map((c,i)=>inner(c,i,(data||[])[i],r))));}
  for(let k=0;k<blankRows;k++){r++;out.push(row(cols.map((c,i)=>inner(c,i,'',r))));}
  if(totals){
    const at=new Set(totals.columns||[]);
    out.push(row(cols.map((c,i)=>i===0&&!at.has(0)?cell(widths[i],para(STYLE_IDS.tableHead,run(clip(totals.label||'Total',60,'The totals label')))):emptyCell(widths[i]))));
  }
  return table(widths,out,{caption:clip(caption,60,'The caption'),description:`A table to fill in, with the columns: ${cols.map(c=>c.name).join(', ')}.${totals?` The last row is for totals.`:''}`});
}

// Statements against a scale, each with a check box in every column, and optionally a column to write the evidence in.
export function ratingGrid({statements,labels=['1','2','3','4','5'],evidence=false,evidenceHeading='Evidence',ctx,caption='Rating'}){
  const list=(statements||[]).map(s=>clip(s,FORM_LIMITS.cell,'A statement')).filter(Boolean);
  if(!list.length)throw new Error('Add at least one statement to rate.');
  if(list.length>FORM_LIMITS.statements)throw new Error(`A rating grid holds up to ${FORM_LIMITS.statements} statements.`);
  const marks=(labels||[]).map(l=>clip(l,30,'A scale label')).filter(Boolean);
  if(marks.length<2||marks.length>7)throw new Error('A scale has 2 to 7 points.');
  const evidenceW=evidence?2400:0,markW=marks.length>5?700:900,statementW=PAGE-evidenceW-markW*marks.length;
  const widths=[statementW,...marks.map(()=>markW),...(evidence?[evidenceW]:[])];
  const rows=[row([head('Statement',statementW),...marks.map(m=>head(m,markW,true)),...(evidence?[head(clip(evidenceHeading,60,'The evidence heading'),evidenceW)]:[])],{header:true}),
    ...list.map(s=>row([cell(statementW,para(STYLE_IDS.tableCell,run(s))),...marks.map(m=>cell(markW,para(STYLE_IDS.ratingBox,checkBox(ctx,`${short(s)}: ${m}`)))),...(evidence?[emptyCell(evidenceW)]:[])]))];
  return table(widths,rows,{caption:clip(caption,60,'The caption'),description:`Statements to rate on a scale from ${marks[0]} to ${marks.at(-1)}. Each statement has a check box under each point${evidence?`, and a column to write the evidence in`:''}.`});
}

// ---- cards
// A set of cards to cut apart: each card is one bordered paragraph, with a dashed line between cards that is the line to cut along.
// perPage (2, 4, 6 or 8) starts a new page after that many cards.
export function cardSet({cards,perPage=4,label=''}){
  const list=(cards||[]).map(c=>lines(c)).filter(l=>l.length);
  if(!list.length)throw new Error('Add at least one card.');
  if(list.length>FORM_LIMITS.cards)throw new Error(`A set holds up to ${FORM_LIMITS.cards} cards.`);
  if(![2,4,6,8].includes(perPage))throw new Error('Cards go 2, 4, 6 or 8 to a page.');
  const words=clip(label,FORM_LIMITS.label,'The label');
  const br='<w:r><w:br/></w:r>';
  const body=list.map((ls,i)=>{
    const text=ls.join(' ');
    if(text.length>FORM_LIMITS.cardWords)throw new Error(`A card holds up to ${FORM_LIMITS.cardWords} letters.`);
    return para(STYLE_IDS.formCard,ls.map(l=>run(l)).join(br),i>0&&i%perPage===0?'<w:pageBreakBefore/>':'');
  });
  return (words?para(STYLE_IDS.formText,run(words)):'')+body.join('');
}

// ---- a self-scored instrument
// Statements to rate by writing a number, a grid to add the numbers up by scale, and what each total means.
// statements: the statements, numbered from 1. groups: [{name,items:[statement numbers]}], each statement in exactly one group. divide: add a
// "divide by" and an "average" column to the grid. bands: [{group,from,to,meaning}] (group is optional: leave it out when one meaning applies to every scale).
export function scoreInstrument({statements,groups,bands=[],divide=false,scale='',caption='Self-scored instrument'}){
  const list=(statements||[]).map(s=>clip(s,FORM_LIMITS.cell,'A statement')).filter(Boolean);
  if(!list.length)throw new Error('Add at least one statement to rate.');
  if(list.length>FORM_LIMITS.rows)throw new Error(`An instrument holds up to ${FORM_LIMITS.rows} statements.`);
  if(!Array.isArray(groups)||!groups.length||groups.length>FORM_LIMITS.groups)throw new Error(`An instrument has 1 to ${FORM_LIMITS.groups} scales to add up.`);
  const named=groups.map(g=>({name:clip(g.name,60,'A scale name'),items:[...new Set((g.items||[]).map(Number))]}));
  if(named.some(g=>!g.name))throw new Error('Every scale needs a name.');
  if(new Set(named.map(g=>g.name.toLowerCase())).size!==named.length)throw new Error('Two scales have the same name.');
  const seen=new Map();
  for(const g of named){
    if(!g.items.length)throw new Error(`The scale ${g.name} has no statements to add up.`);
    for(const n of g.items){
      if(!Number.isInteger(n)||n<1||n>list.length)throw new Error(`The scale ${g.name} names statement ${n}, which is not one of the ${list.length}.`);
      if(seen.has(n))throw new Error(`Statement ${n} is in two scales: ${seen.get(n)} and ${g.name}.`);
      seen.set(n,g.name);
    }
  }
  const unmapped=list.map((_,i)=>i+1).filter(n=>!seen.has(n));
  if(unmapped.length)throw new Error(`${unmapped.length===1?`Statement ${unmapped[0]} is`:`Statements ${unmapped.join(', ')} are`} not in any scale, so ${unmapped.length===1?'it':'they'} would not be added up.`);
  const meanings=(bands||[]).map(b=>{
    const from=Number(b.from),to=Number(b.to),meaning=clip(b.meaning,FORM_LIMITS.cell,'A meaning'),group=b.group?clip(b.group,60,'A scale name'):'';
    if(!Number.isFinite(from)||!Number.isFinite(to)||to<from)throw new Error('A range of scores runs from a lower number to a higher one.');
    if(!meaning)throw new Error('Every range of scores needs words that say what it means.');
    if(group&&!named.some(g=>g.name.toLowerCase()===group.toLowerCase()))throw new Error(`A range is for the scale ${group}, which is not one of the scales.`);
    return {group,from,to,meaning};
  });
  if(meanings.length>FORM_LIMITS.bands)throw new Error(`An instrument holds up to ${FORM_LIMITS.bands} ranges.`);
  const intro=text=>para(STYLE_IDS.formText,run(text));
  // 1. rate each statement
  const w1=[900,PAGE-900-1500,1500];
  const t1=table(w1,[row([head('No.',w1[0],true),head('Statement',w1[1]),head(scale?`Score (${clip(scale,40,'The scale')})`:'Score',w1[2],true)],{header:true}),
    ...list.map((s,i)=>row([cell(w1[0],para(STYLE_IDS.tableCell,run(String(i+1))),{}),cell(w1[1],para(STYLE_IDS.tableCell,run(s))),emptyCell(w1[2])]))],
    {caption:clip(caption,60,'The caption'),description:'The statements to rate. Write your score for each one in the last column.'});
  // 2. add up the scores
  const cols=divide?['Scale','Statements to add','Total','Divide by','Average']:['Scale','Statements to add','Total'];
  const w2=divide?[2200,PAGE-2200-1300-1300-1300,1300,1300,1300]:[2200,PAGE-2200-1500,1500];
  const t2=table(w2,[row(cols.map((c,i)=>head(c,w2[i],i>1)),{header:true}),
    ...named.map(g=>row([cell(w2[0],para(STYLE_IDS.tableCell,run(g.name))),cell(w2[1],para(STYLE_IDS.tableCell,run(g.items.slice().sort((a,b)=>a-b).join(', ')))),emptyCell(w2[2]),
      ...(divide?[cell(w2[3],para(STYLE_IDS.tableCell,run(String(g.items.length)),'<w:jc w:val="center"/>')),emptyCell(w2[4])]:[])]))],
    {caption:`${clip(caption,60,'The caption')}: adding up`,description:`For each scale, add the scores of the statements listed${divide?', divide the total by the number shown':''}, and write the result in the last column.`});
  // 3. what the scores mean
  let t3='';
  if(meanings.length){
    const hasGroup=meanings.some(m=>m.group),cols3=hasGroup?['Scale','Score','What it means']:['Score','What it means'];
    const w3=hasGroup?[2200,1500,PAGE-3700]:[1500,PAGE-1500];
    t3=table(w3,[row(cols3.map((c,i)=>head(c,w3[i])),{header:true}),
      ...meanings.map(m=>{
        const range=m.from===m.to?String(m.from):`${m.from} to ${m.to}`;
        return row([...(hasGroup?[cell(w3[0],para(STYLE_IDS.tableCell,run(m.group||'Every scale')))]:[]),cell(hasGroup?w3[1]:w3[0],para(STYLE_IDS.tableCell,run(range))),cell(hasGroup?w3[2]:w3[1],para(STYLE_IDS.tableCell,run(m.meaning)))]);
      })],
      {caption:`${clip(caption,60,'The caption')}: what the scores mean`,description:'Which range of scores means what, for each scale.'});
  }
  return intro('Rate each statement.')+t1+intro('Add up your scores.')+t2+(t3?intro('What your scores mean.')+t3:'');
}
