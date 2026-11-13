// Wholly synthetic long Unicode definition; no responses or external QA dependencies.
export function unicodeAssessmentDefinition(count){
 const order=count===2?[1,0]:[10,0,9,1,8,2,7,3,6,4,5],lo=count===2?-918101:811001;
 const d={schemaVersion:1,id:`synthetic-print-review-${count}`,version:9,title:`Synthetic ${count}-anchor workshop reflection — authored review only`,evidenceStatus:'not-established',items:[],subscales:[]};
 for(let i=0;i<24;i++){
  const semantic=i%2===1;
  const item={id:`KEYONLY_ITEM_${count}_${i}`,type:semantic?'semantic-differential':'likert',prompt:`Public item ${i+1}: Consider the explanation, application and opportunity to reflect. ${'Describe the session context before choosing the anchor that fits your experience. '.repeat(5)}\r\nPreserve this second prompt line & <literal> text\tfrançais, Deutsch, 中文 🙂.`,reverse:i%3===1,options:order.map((n,j)=>({id:`KEYONLY_OPTION_${count}_${i}_${n}`,label:`Public anchor ${i+1}.${j+1} — ${'A deliberately descriptive response label remains complete when it wraps across several printed lines. '.repeat(2)}${j%2?'\r\nSecond label line':'\tIndented clause'} · café, Überblick, 中文 🙂 & < >`,code:lo+n}))};
  if(semantic)Object.assign(item,{target:`Public synthetic learning explanation ${i+1}`,leftEndpoint:`Unclear — ${'needs a more explicit connection to the learning goal '.repeat(3)}`,rightEndpoint:`Clear — ${'the connection to the learning goal is readily explained '.repeat(3)}`,highScoreMeaning:`KEYONLY_HIGH_MEANING_${count}_${i}`});
  d.items.push(item);
 }
 d.subscales=[{id:`KEYONLY_SCALE_MEAN_${count}`,name:`KEYONLY mean teaching review ${count}`,itemIds:d.items.slice(0,12).map(x=>x.id),aggregation:'mean',missing:{mode:'minimum-answered',minimumAnswered:9}},{id:`KEYONLY_SCALE_SUM_${count}`,name:`KEYONLY sum teaching review ${count}`,itemIds:d.items.slice(12).map(x=>x.id),aggregation:'sum',missing:{mode:'complete-case'}}];
 return d;
}

export function unicodeAssessmentOptions(route,baseLanguage){
 return {route,instructions:`Synthetic review instructions — no real respondent data.\r\nSelect one response per item, leave missing separate from Not applicable, and review every complete label.\tDo not infer validation or a diagnostic cutoff. ${'The labels and endpoints describe this example session only; long text must remain readable and complete. '.repeat(7)}\r\nMultilingual text preserved, not per-span language-tagged: français, Deutsch, 中文 🙂.`,baseLanguage,stamp:'2026-10-28T16:55:00.000Z'};
}
