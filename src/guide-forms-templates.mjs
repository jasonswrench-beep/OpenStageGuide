// The forms the program starts an author with: a title that is numbered by Update guide, the print line, and the elements of each kind of form.
// Pure functions that give Word-format XML, built from the same element builders everything else uses (guide-forms-ooxml.mjs), so a form
// made from a template is the same as one made element by element. Every template is a start to write over, not a finished form: its
// words say what to type where.
import {formTitleBlock,printLine,formText,writingLines,writingBox,fieldRow,checkList,formTable,ratingGrid,cardSet,sourceLine,scoreInstrument} from './guide-forms-ooxml.mjs';
import {LABELS,DEFAULT_LABEL} from './guide-forms.mjs';

/** @typedef {{id:string,label:string,kind:string,note:string,facilitator?:boolean,body:(ctx:object)=>string}} FormTemplate */

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
  {id:'instrument',label:'Self-scored instrument',kind:'Form',note:'Statements to score, a grid to add the scores up by scale, and what each total means.',
    body:()=>scoreInstrument({statements:['First statement','Second statement','Third statement','Fourth statement'],groups:[{name:'First scale',items:[1,2]},{name:'Second scale',items:[3,4]}],
      bands:[{from:2,to:5,meaning:'Say here what a low total means.'},{from:6,to:10,meaning:'Say here what a high total means.'}],scale:'1 (never) to 5 (always)',caption:'Self-scored instrument'})},
  {id:'observer',label:'Observer form',kind:'Form',note:'Behaviors that were met or not met, comments, signature and date.',
    body:ctx=>fieldRow([{label:'Observer',width:'full'},{label:'Date',width:'short'}])
      +formTable({columns:[{name:'Behavior',weight:4},{name:'Met',weight:1,kind:'check'},{name:'Not met',weight:1,kind:'check'},{name:'Comments',weight:4}],rows:[['First behavior'],['Second behavior'],['Third behavior']],ctx,caption:'Observed behaviors'})
      +fieldRow([{label:'Signature',width:'full'}])},
  {id:'evaluation',label:'Evaluation form',kind:'Form',note:'Session and date, statements to rate from 1 to 5, and two questions to answer in words.',
    body:ctx=>fieldRow([{label:'Session',width:'full'},{label:'Date',width:'short'}])
      +formText('Rate each statement from 1 (disagree) to 5 (agree).')
      +ratingGrid({statements:['The session met its goals.','The activities helped me learn.','I can use what I learned.'],ctx,caption:'Ratings'})
      +writingLines({label:'What worked well?',count:3})+writingLines({label:'What would you change?',count:3})},
  {id:'action',label:'Action plan',kind:'Template',note:'What, who, by when, and a check box for done.',
    body:ctx=>formTable({columns:[{name:'Action',weight:4},{name:'Who',weight:2},{name:'By when',weight:2},{name:'Done',weight:1,kind:'check'}],blankRows:6,ctx,caption:'Action plan'})},
  {id:'commitment',label:'Commitment card',kind:'Cards',note:'A card to fill in and keep.',
    body:()=>cardSet({cards:['I will:\n\nBy when:\n\nSigned:'],perPage:4})},
  {id:'rolecards',label:'Role cards',kind:'Cards',note:'One card for each role in a role play, to cut apart.',
    body:()=>cardSet({cards:['First role\nWho you are and what you want.','Second role\nWho you are and what you want.','Third role\nWho you are and what you want.'],perPage:4})},
  {id:'matchup',label:'Match-up cards',kind:'Cards',note:'Pairs of cards to match, such as a term and its meaning, to cut apart.',
    body:()=>cardSet({cards:['First term','Meaning of the first term','Second term','Meaning of the second term','Third term','Meaning of the third term'],perPage:6})},
  {id:'twobytwo',label:'Two-by-two',kind:'Template',note:'Two rows and two columns to sort ideas into, with a name for each side.',
    body:ctx=>formTable({columns:[{name:'Impact',weight:1},{name:'Low effort',weight:2},{name:'High effort',weight:2}],rows:[['High impact'],['Low impact']],ctx,caption:'Two-by-two'})},
  {id:'forcefield',label:'Force-field',kind:'Template',note:'The goal, the forces that help and the forces that hold back.',
    body:ctx=>fieldRow([{label:'The change we want',width:'full'}])+formTable({columns:[{name:'Forces that help',weight:1},{name:'Forces that hold back',weight:1}],blankRows:6,ctx,caption:'Force-field'})},
  {id:'planning',label:'Planning checklist',kind:'Checklist',note:'Tasks before, on the day and after, with who and by when.',
    body:ctx=>checkList({items:['Before: first task','Before: second task','On the day: first task','After: first task'],columns:['Who','By when'],ctx,caption:'Planning checklist'})},
  // The planning kit (kit slice K5). The ones the facilitator or the organiser uses are for the facilitator only whatever the author ticks: they never go in the participant workbook.
  {id:'timeline',label:'Planning timeline',kind:'Checklist',facilitator:true,note:'Stages from weeks before the event to after it, with tasks to tick off, an owner, the lead time and notes.',
    body:ctx=>formTable({columns:[{name:'When',weight:2},{name:'Task',weight:4},{name:'Done',weight:1,kind:'check'},{name:'Owner',weight:2},{name:'Lead time',weight:2},{name:'Notes',weight:3}],
      rows:[['Six weeks out'],['Four weeks out'],['Two weeks out'],['One week out'],['The day before'],['After the event']],blankRows:4,ctx,caption:'Planning timeline'})},
  {id:'techmaterials',label:'Technology and materials',kind:'Checklist',facilitator:true,note:'What is needed, how many, who brings it, and a box to tick when it is ready.',
    body:ctx=>formTable({columns:[{name:'Item',weight:4},{name:'How many',weight:1},{name:'Who is responsible',weight:3},{name:'Ready',weight:1,kind:'check'}],
      rows:[['Projector and screen'],['Flip charts and markers'],['Handouts, copied'],['Name tags']],blankRows:4,ctx,caption:'Technology and materials'})},
  {id:'roomsetup',label:'Room set-up',kind:'Template',facilitator:true,note:'The seating and layout, with a box to draw it and a list of what to do and not to do.',
    body:ctx=>fieldRow([{label:'Room',width:'full'},{label:'Seating',width:'full'}])+formText('Draw the layout here.')+writingBox({lines:8})
      +formTable({columns:[{name:'Do',weight:1},{name:'Do not',weight:1}],blankRows:4,ctx,caption:'Room set-up: what to do and not to do'})},
  {id:'roles',label:'Roles grid',kind:'Template',facilitator:true,note:'Who does what: lead, table lead, timekeeper and the others.',
    body:ctx=>formTable({columns:[{name:'Role',weight:2},{name:'Name',weight:3},{name:'Does what',weight:5}],
      rows:[['Lead facilitator'],['Co-facilitator'],['Table lead'],['Timekeeper'],['Producer']],blankRows:2,ctx,caption:'Roles'})},
  {id:'prework',label:'Pre-work',kind:'Handout',note:'What learners read or do before the session, with a box to tick for each.',
    body:ctx=>formText('Before the session, please do the following.')+checkList({items:['Read the first item','Do the second item'],columns:['Done on'],ctx,caption:'Pre-work'})},
  {id:'followup',label:'Follow-up roadmap',kind:'Template',note:'What happens after the session, week by week, with who and a box to tick.',
    body:ctx=>formTable({columns:[{name:'When',weight:2},{name:'What',weight:5},{name:'Who',weight:2},{name:'Done',weight:1,kind:'check'}],
      rows:[['Week 1'],['Week 4'],['Week 12'],['Week 26']],blankRows:2,ctx,caption:'Follow-up roadmap'})}
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
  // A form of the planning kit is for the facilitator whatever was asked: it is not for learners and never goes in their workbook.
  facilitator=facilitator||t.facilitator===true;
  const kind=LABELS.includes(label)?label:(t.kind||DEFAULT_LABEL);
  return formTitleBlock({id,label:kind,number:'',title,facilitator})
    +printLine({copies:{n:1,per:'person'},orientation:'portrait'})
    +t.body(ctx);
}
