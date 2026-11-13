// Builds the Word content for assessment sections and items. Pure functions, like guide-ooxml.mjs.
// An item is a question, its options and answer line (for everyone), then its answer key (facilitator only, in its
// own style, never hidden text). The key's first paragraph says which kind of item it is, so the item can be read back.
import {para,run,lines,esc} from './guide-ooxml.mjs';
import {STYLE_IDS} from './guide-styles.mjs';
import {itemTypeById,levelById,matchingChoices,scaleLabels,scrambleOrder,SEMANTIC_POINTS,ACTION_PROMPTS,LIMITS,numericKey} from './guide-items.mjs';

const LETTERS='abcdefghijklmnopqrstuvwxyz';
const LINE='5B6B82',HEAD_FILL='E4E8EE',PAGE=9360; // the table's rules and header shading; the text width in twips

// A table is used only where the content really is one. It has a header row that Word knows is a header row, no
// merged cells, a caption and a description for a screen reader, and every cell holds words or a check box.
const cell=(width,inner,{fill}={})=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${fill?`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`:''}<w:vAlign w:val="center"/></w:tcPr>${inner}</w:tc>`;
const row=(cells,{header=false}={})=>`<w:tr><w:trPr><w:cantSplit/>${header?'<w:tblHeader/>':''}</w:trPr>${cells.join('')}</w:tr>`;
const rule=side=>`<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${LINE}"/>`;
const shorten=(text,n)=>{const t=String(text).replace(/\s+/g,' ').trim();return t.length>n?`${t.slice(0,n-1)}…`:t;};
function table(widths,rows,{caption,description}){
  const grid=widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a,b)=>a+b,0)}" w:type="dxa"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(rule).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblCaption w:val="${esc(shorten(caption,120))}"/><w:tblDescription w:val="${esc(shorten(description,250))}"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows.join('')}</w:tbl>`;
}
const head=(text,width,center)=>cell(width,para(STYLE_IDS.tableHead,run(text),center?'<w:jc w:val="center"/>':''),{fill:HEAD_FILL});
// A real Word check box content control: clickable in Word, an empty box on paper, named for a screen reader.
const checkBox=(ctx,alias)=>`<w:sdt><w:sdtPr><w:alias w:val="${esc(alias)}"/><w:tag w:val="osg-rating"/><w:id w:val="${700000+(++ctx.sdt)}"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;
export const EMPTY_BOX='☐';
export {cell,row,table,head,checkBox,PAGE};
const br='<w:r><w:br/></w:r>';
export const questionLabel=n=>`Question ${n}.`;
export const optionLetter=i=>LETTERS[i];

// One heading (in the contents), the note to the learner, and the information lines to fill in.
export function sectionBlock(levelId){
  const level=levelById(levelId);
  if(!level)throw new Error('Choose a Level One, Two or Three assessment.');
  return [
    para(STYLE_IDS.assessment,run(level.title)),
    para(null,run(level.intro)),
    ...level.info.map(line=>para(null,run(line)))
  ].join('');
}

const questionPara=(number,stem,type)=>{
  const text=lines(stem);
  if(!text.length)throw new Error('Type the question first.');
  const body=text.map(l=>run(l)).join(br);
  const instruction=type.instruction?run(` ${type.instruction}`):'';
  return para(STYLE_IDS.question,(type.noNumber?'':run(questionLabel(number),STYLE_IDS.cueLabel)+run(' '))+body+instruction);
};
const answerLines=(label,count)=>[
  para(STYLE_IDS.answerLine,run(label)),
  ...Array.from({length:Math.max(0,count-1)},()=>para(STYLE_IDS.answerLine,''))
];
const keyParas=(type,keyText,{why,points,objective})=>{
  const out=[para(STYLE_IDS.answerKey,run(`Answer key (${type.keyLabel}): `,STYLE_IDS.cueLabel)+run(keyText))];
  if(String(why||'').trim())out.push(para(STYLE_IDS.answerKey,run('Why: ',STYLE_IDS.cueLabel)+run(String(why).trim())));
  if(points!==undefined&&points!==''&&points!==null)out.push(para(STYLE_IDS.answerKey,run('Points: ',STYLE_IDS.cueLabel)+run(String(points))));
  if(String(objective||'').trim())out.push(para(STYLE_IDS.answerKey,run('Objective: ',STYLE_IDS.cueLabel)+run(String(objective).trim())));
  return out;
};

// Matching: the items in a table (number, item, a space for the letter), then the lettered choices, then the key.
function matchingParts(item,ctx){
  const pairs=(item.pairs||[]).map(p=>({item:String(p?.item??'').trim(),match:String(p?.match??'').trim()})).filter(p=>p.item||p.match);
  pairs.forEach((p,i)=>{if(!p.item||!p.match)throw new Error(`Each line needs an item and its match, separated by | or a tab. Line ${i+1} is missing ${p.item?'its match':'the item'}.`);});
  if(pairs.length<2)throw new Error('Type at least two pairs, one to a line, each as an item and its match.');
  if(pairs.length>LIMITS.pairs)throw new Error(`Matching can have up to ${LIMITS.pairs} pairs. Use a second item for more.`);
  const extras=(item.extras||[]).map(e=>String(e??'').trim()).filter(Boolean);
  if(extras.length>LIMITS.extras)throw new Error(`Up to ${LIMITS.extras} extra choices.`);
  const choices=matchingChoices(pairs,extras),letterOf=text=>optionLetter(choices.findIndex(c=>c.toLowerCase()===text.toLowerCase()));
  const widths=[700,5700,PAGE-6400];
  const rows=[
    row([head('No.',widths[0]),head('Item',widths[1]),head('Your match',widths[2])],{header:true}),
    ...pairs.map((p,i)=>row([cell(widths[0],para(STYLE_IDS.tableCell,run(String(i+1)))),cell(widths[1],para(STYLE_IDS.tableCell,run(p.item))),cell(widths[2],para(STYLE_IDS.tableCell,run('Letter:')))]))
  ];
  const list=++ctx.lists;
  return {
    parts:[
      table(widths,rows,{caption:'Matching items',description:'One row for each item. Write the letter of its match in the last column. The choices are listed below the table.'}),
      para(STYLE_IDS.itemNote,run('Choices:')),
      ...choices.map(c=>para(STYLE_IDS.option,run(c),`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list}"/></w:numPr>`))
    ],
    keyText:`${pairs.map((p,i)=>`${i+1} = ${letterOf(p.match)}`).join('; ')}.`
  };
}

// Likert scale (the item type was called Rating statements until 2026-11-07; its id is still rating-statements): one table, a row for each statement, a check box under each rating.
function ratingParts(item,ctx){
  const statements=(item.statements||[]).map(s=>String(s??'').trim()).filter(Boolean);
  if(!statements.length)throw new Error('Type at least one statement, one to a line.');
  if(statements.length>LIMITS.statements)throw new Error(`A table can have up to ${LIMITS.statements} statements. Use a second item for more.`);
  const labels=scaleLabels(item);
  if(labels.length<LIMITS.minScale||labels.length>LIMITS.maxScale)throw new Error(`Give the scale between ${LIMITS.minScale} and ${LIMITS.maxScale} labels, separated by commas.`);
  const boxW=labels.length<=5?1100:Math.floor(5600/labels.length),widths=[PAGE-labels.length*boxW,...labels.map(()=>boxW)];
  const rows=[
    row([head('Statement',widths[0]),...labels.map((l,i)=>head(l,widths[i+1],true))],{header:true}),
    ...statements.map(s=>row([cell(widths[0],para(STYLE_IDS.tableCell,run(s))),...labels.map((l,i)=>cell(widths[i+1],para(STYLE_IDS.ratingBox,checkBox(ctx,`${shorten(s,40)}: ${l}`))))]))
  ];
  return {
    parts:[table(widths,rows,{caption:`Likert scale: ${item.stem}`,description:`One row for each statement. Each row has a check box under each rating: ${labels.join(', ')}.`})],
    keyText:''
  };
}

// Ordering: the items mixed up as a lettered list, one line for the letters in order, and a key that gives the right order.
function orderingParts(item,ctx){
  const items=(item.orderItems||[]).map(s=>String(s??'').trim()).filter(Boolean);
  if(items.length<LIMITS.minOrderItems)throw new Error(`Type at least ${LIMITS.minOrderItems} items, one to a line, in the right order.`);
  if(items.length>LIMITS.orderItems)throw new Error(`An ordering question can have up to ${LIMITS.orderItems} items.`);
  if(new Set(items.map(s=>s.toLowerCase())).size!==items.length)throw new Error('Each item must be different, so the key can name them by letter.');
  const shown=scrambleOrder(items),list=++ctx.lists;
  return {
    parts:shown.map(s=>para(STYLE_IDS.option,run(s),`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list}"/></w:numPr>`)),
    keyText:`${items.map(s=>optionLetter(shown.indexOf(s))).join(', ')}.`
  };
}

// Performance task: the criteria go in the key, one paragraph each, for the facilitator to score by.
function criteriaParas(item){
  const criteria=(item.criteria||[]).map(c=>({text:String(c?.text??'').trim(),points:String(c?.points??'').trim()})).filter(c=>c.text);
  if(!criteria.length)throw new Error('Type at least one criterion for the checklist, one to a line. Add points after a | if you want them.');
  if(criteria.length>LIMITS.criteria)throw new Error(`A checklist can have up to ${LIMITS.criteria} criteria.`);
  return criteria.map((c,i)=>para(STYLE_IDS.answerKey,run(`Criterion ${i+1}: `,STYLE_IDS.cueLabel)+run(`${c.text}${c.points?` (${c.points} ${Number(c.points)===1?'point':'points'})`:''}`)));
}

// Performance task with a rubric: a table in the key, a row for each criterion and a column for each level. It uses the key's own
// table styles, so the learner workbook leaves the whole table out as it leaves out every other part of a key.
function rubricTable(item,learners=false){
  const levels=(item.levels||[]).map(l=>({label:String(l?.text??l?.label??'').trim(),points:String(l?.points??'').trim()})).filter(l=>l.label);
  if(levels.length<LIMITS.minLevels||levels.length>LIMITS.maxLevels)throw new Error(`A rubric has ${LIMITS.minLevels} to ${LIMITS.maxLevels} levels, one to a line, like: Meets | 2. Type the levels.`);
  if(new Set(levels.map(l=>l.label.toLowerCase())).size!==levels.length)throw new Error('Each level must have a different name.');
  levels.forEach(l=>{if(l.points&&!/^\d{1,3}(\.\d{1,2})?$/.test(l.points))throw new Error(`The points for "${l.label}" must be a number, like 2 or 2.5.`);});
  const rows=(item.rubricRows||[]).map(r=>(r||[]).map(c=>String(c??'').trim())).filter(r=>r.some(Boolean));
  if(!rows.length)throw new Error('Type at least one criterion, one to a line: the criterion, then a description of it for each level, separated by |.');
  if(rows.length>LIMITS.rubricRows)throw new Error(`A rubric can have up to ${LIMITS.rubricRows} criteria.`);
  rows.forEach((r,i)=>{
    if(r.length!==levels.length+1)throw new Error(`Line ${i+1} has ${r.length-1} ${r.length===2?'description':'descriptions'} and the rubric has ${levels.length} levels. Write the criterion, then one description for each level, separated by |.`);
    const blank=r.findIndex(c=>!c);
    if(blank>=0)throw new Error(`Line ${i+1} has an empty ${blank===0?'criterion':`description for "${levels[blank-1].label}"`}. Every cell of the table needs words.`);
  });
  const first=2200,rest=Math.floor((PAGE-first)/levels.length),widths=[first,...levels.map(()=>rest)];
  const KEY_FILL='F1EBDD';
  // For the facilitator the table is in the key's styles; for learners it is an ordinary table that goes with the task.
  const keyHead=learners?(text,w)=>head(text,w):(text,w)=>cell(w,para(STYLE_IDS.keyTableHead,run(text)),{fill:HEAD_FILL});
  const keyCell=learners?(text,w)=>cell(w,para(STYLE_IDS.tableCell,run(text))):(text,w)=>cell(w,para(STYLE_IDS.keyTableCell,run(text)),{fill:KEY_FILL});
  const trs=[
    row([keyHead('Criterion',first),...levels.map((l,i)=>keyHead(`${l.label}${l.points?` (${l.points} ${Number(l.points)===1?'point':'points'})`:''}`,rest))],{header:true}),
    ...rows.map(r=>row(r.map((c,i)=>keyCell(c,widths[i]))))
  ];
  return table(widths,trs,{caption:`Rubric: ${item.stem}`,description:`One row for each criterion, described at each level: ${levels.map(l=>l.label).join(', ')}.${learners?'':' This is the facilitator\'s key.'}`});
}

// Semantic differential: pairs of opposite words with a check box for each point of the scale between them.
function semanticParts(item,ctx){
  const pairs=(item.sdPairs||[]).map(p=>({left:String(p?.item??'').trim(),right:String(p?.match??'').trim()})).filter(p=>p.left||p.right);
  pairs.forEach((p,i)=>{if(!p.left||!p.right)throw new Error(`Each line needs two words, one for each end, separated by | or a tab. Line ${i+1} is missing ${p.left?'the second word':'the first word'}.`);});
  if(!pairs.length)throw new Error('Type at least one pair of opposite words, one pair to a line, like Boring | Interesting.');
  if(pairs.length>LIMITS.semanticPairs)throw new Error(`A table can have up to ${LIMITS.semanticPairs} pairs. Use a second item for more.`);
  const n=Number(item.sdPoints)||5;
  if(!SEMANTIC_POINTS.includes(n))throw new Error(`Choose a scale of ${SEMANTIC_POINTS.join(' or ')} points.`);
  const side=2200,box=Math.floor((PAGE-2*side)/n),widths=[side,...Array(n).fill(box),side];
  const rows=[
    row([head('Left word',side),...Array.from({length:n},(_,i)=>head(String(i+1),box,true)),head('Right word',side)],{header:true}),
    ...pairs.map(p=>row([cell(side,para(STYLE_IDS.tableCell,run(p.left))),...Array.from({length:n},(_,i)=>cell(box,para(STYLE_IDS.ratingBox,checkBox(ctx,`${shorten(p.left,25)} to ${shorten(p.right,25)}: ${i+1} of ${n}`)))),cell(side,para(STYLE_IDS.tableCell,run(p.right)))]))
  ];
  return [table(widths,rows,{caption:`Semantic differential: ${item.stem}`,description:`One row for each pair of words. Each row has ${n} check boxes between the two words, numbered 1 to ${n} from the left word to the right word.`})];
}

// Action plan: for each action, four labelled lines to write on.
function actionParts(item){
  const count=Number(item.actions);
  if(!Number.isInteger(count)||count<1||count>LIMITS.actions)throw new Error(`Choose how many actions to ask for: 1 to ${LIMITS.actions}.`);
  return Array.from({length:count},(_,i)=>[para(STYLE_IDS.itemNote,run(`Action ${i+1}`)),...ACTION_PROMPTS.map(p=>para(STYLE_IDS.answerLine,run(p)))]).flat();
}

// item: {type, number, stem, options:[text], correct:[index], truth:true|false, answer, why, points, objective, lines,
//        pairs:[{item,match}], extras:[text], statements:[text], scale:'agree'|'success'|'custom', labels:[text],
//        orderItems:[text], criteria:[{text,points}], sdPairs:[{item,match}], sdPoints:5|7, actions:1..5}
export function itemBlock(item,ctx){
  const type=itemTypeById(item.type);
  if(!type)throw new Error('Choose a type of item first.');
  const out=[questionPara(item.number,item.stem,type)];
  let keyText='',extraKey=[];
  if(type.matching){
    const m=matchingParts(item,ctx);
    out.push(...m.parts);keyText=m.keyText;
  }else if(type.rating){
    out.push(...ratingParts(item,ctx).parts);
  }else if(type.semantic){
    out.push(...semanticParts(item,ctx));
  }else if(type.action){
    out.push(...actionParts(item));
  }else if(type.ordering){
    const o=orderingParts(item,ctx);
    out.push(...o.parts);keyText=o.keyText;
  }else if(type.rubric){
    // The rubric is in the key for the facilitator, or, when the author wants learners to see how they will be scored, under the task.
    if(item.learnerRubric)out.push(rubricTable(item,true));else extraKey=[rubricTable(item,false)];
    keyText='score each criterion by the level the learner reaches.';
  }else if(type.performance){
    extraKey=criteriaParas(item);keyText='score each criterion the learner meets.';
  }else if(type.options){
    const options=(item.options||[]).map(o=>String(o??'').trim()).map((text,i)=>({text,i})).filter(o=>o.text);
    if(options.length<2)throw new Error('Type at least two options.');
    const correct=(item.correct||[]).filter(i=>options.some(o=>o.i===i));
    if(type.pick==='one'&&correct.length!==1)throw new Error('Mark the one correct option.');
    if(type.pick==='many'&&correct.length<1)throw new Error('Mark at least one correct option.');
    const list=++ctx.lists;
    options.forEach(o=>out.push(para(STYLE_IDS.option,run(o.text),`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list}"/></w:numPr>`)));
    // The key names the options by the letters the learner sees, which are their positions among the options typed.
    const letter=i=>optionLetter(options.findIndex(o=>o.i===i));
    keyText=`${correct.map(letter).join(', ')}.`;
  }else if(type.truefalse){
    if(item.truth!==true&&item.truth!==false)throw new Error('Choose whether the statement is true or false.');
    keyText=item.truth?'True.':'False.';
  }else if(type.numeric){
    keyText=numericKey(item);
  }else if(type.answer){
    const a=String(item.answer??'').trim();
    if(!a)throw new Error(`Type the ${type.answer.toLowerCase().replace(/ \(.*\)/,'')} for the key.`);
    keyText=a;
  }
  // Matching, rating statements, semantic differentials and action plans are answered in their tables or on their own lines.
  // A numeric item says in what units the learner answers, when the key gives units.
  const answerLabel=type.numeric&&String(item.units??'').trim()?`Your answer in ${String(item.units).replace(/\s+/g,' ').trim()}:`:type.answerLine;
  if(!type.matching&&!type.rating&&!type.semantic&&!type.action)out.push(...answerLines(answerLabel,type.options||type.truefalse||type.numeric||type.id==='fill-blank'?1:Math.max(1,Number(item.lines)||type.lines||1)));
  if(!type.noKey){
    // A checklist's criteria follow the first key paragraph (which names the type) and come before Why, Points and Objective.
    const key=keyParas(type,keyText,item);
    key.splice(1,0,...extraKey);
    out.push(...key);
  }
  return out.join('');
}

export {esc};
