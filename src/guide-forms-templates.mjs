// The forms the program starts an author with: a title that is numbered by Update guide, the print line, and the elements of each kind of form.
// Pure functions that give Word-format XML, built from the same element builders everything else uses (guide-forms-ooxml.mjs), so a form
// made from a template is the same as one made element by element. Every template is a start to write over, not a finished form: its
// words say what to type where.
import {formTitleBlock,printLine,formText,writingLines,fieldRow,checkList,formTable,ratingGrid,cardSet,sourceLine} from './guide-forms-ooxml.mjs';
import {LABELS,DEFAULT_LABEL} from './guide-forms.mjs';

/** @typedef {{id:string,label:string,kind:string,note:string,body:(ctx:object)=>string}} FormTemplate */

/** @type {readonly FormTemplate[]} */
export const FORM_TEMPLATES=Object.freeze([
  {id:'reading',label:'Reading handout',kind:'Handout',note:'Text to read, with a credit line.',
    body:()=>formText('Type the text of the handout here. Each line you type becomes a paragraph.')+sourceLine('Credit or permission for the text above.')},
  {id:'worksheet',label:'Blank worksheet',kind:'Worksheet',note:'Name and date, a task, and lines to write on.',
    body:()=>fieldRow([{label:'Name',width:'full'},{label:'Date',width:'short'}])+formText('Write the question or the task here.')+writingLines({count:6})},
  {id:'notes',label:'Note-taking grid',kind:'Handout',note:'Topic, notes and questions, with empty rows to write in.',
    body:ctx=>formTable({columns:[{name:'Topic',weight:2},{name:'Notes',weight:4},{name:'Questions',weight:2}],blankRows:8,ctx,caption:'Notes'})},
  {id:'checklist',label:'Checklist',kind:'Checklist',note:'Items with a check box, and who and when.',
    body:ctx=>checkList({items:['First item','Second item','Third item'],columns:['Who','When'],ctx})},
  {id:'rating',label:'Rating form with evidence',kind:'Form',note:'Statements to rate from 1 to 5, with a column for the evidence.',
    body:ctx=>formText('Rate each statement from 1 (not at all) to 5 (completely), and write what you saw or heard.')+ratingGrid({statements:['First statement','Second statement','Third statement'],evidence:true,ctx,caption:'Ratings'})},
  {id:'observer',label:'Observer form',kind:'Form',note:'Behaviours that were met or not met, comments, signature and date.',
    body:ctx=>fieldRow([{label:'Observer',width:'full'},{label:'Date',width:'short'}])
      +formTable({columns:[{name:'Behaviour',weight:4},{name:'Met',weight:1,kind:'check'},{name:'Not met',weight:1,kind:'check'},{name:'Comments',weight:4}],rows:[['First behaviour'],['Second behaviour'],['Third behaviour']],ctx,caption:'Observed behaviours'})
      +fieldRow([{label:'Signature',width:'full'}])},
  {id:'action',label:'Action plan',kind:'Template',note:'What, who, by when, and a check box for done.',
    body:ctx=>formTable({columns:[{name:'Action',weight:4},{name:'Who',weight:2},{name:'By when',weight:2},{name:'Done',weight:1,kind:'check'}],blankRows:6,ctx,caption:'Action plan'})},
  {id:'commitment',label:'Commitment card',kind:'Cards',note:'A card to fill in and keep.',
    body:()=>cardSet({cards:['I will:\n\nBy when:\n\nSigned:'],perPage:4})}
]);

export const templateById=id=>FORM_TEMPLATES.find(t=>t.id===id)||null;

/**
 * A whole form from a template: the numbered title (the number is written by Update guide), the print line, and the template's elements.
 * @param {{template:string, id:string, label?:string, title:string, facilitator?:boolean}} spec
 * @param {object} ctx the insert's context (from the document)
 */
export function buildForm({template,id,label,title,facilitator=false},ctx){
  const t=templateById(template);
  if(!t)throw new Error('Choose a kind of form.');
  const kind=LABELS.includes(label)?label:(t.kind||DEFAULT_LABEL);
  return formTitleBlock({id,label:kind,number:'',title,facilitator})
    +printLine({copies:{n:1,per:'person'},orientation:'portrait'})
    +t.body(ctx);
}
