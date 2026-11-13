// Loaded only for explicitly marked disposable packages, never ordinary guides.
export function attachAcceptancePanel({word,run,refresh,commit,getProject,isBusy}){
  let locked=false;const buttons=[];
  async function guarded(action){
    if(locked||isBusy())return;
    locked=true;for(const b of buttons){b.dataset.acceptanceLocked='1';b.disabled=true;}
    try{await action();}finally{locked=false;for(const b of buttons){b.dataset.acceptanceLocked='0';b.disabled=isBusy();}}
  }
  function button(){const b=document.createElement('button');b.dataset.acceptanceAction='1';buttons.push(b);return b;}
  const box=document.createElement('details'),summary=document.createElement('summary'),result=document.createElement('pre');
  result.style.whiteSpace='pre-wrap';result.style.overflowWrap='anywhere';
  summary.textContent='Disposable document — synthetic failure acceptance';box.append(summary);
  const cases=[['Save before preparation',{kind:'save',when:'before',nth:1}],['Save failure after real commit',{kind:'save',when:'after',nth:2}],['Partial native insertion',{kind:'write-sync',when:'after',nth:1}],['Complete native insertion',{kind:'write-sync',when:'after',nth:2}]];
  async function snapshot(){try{result.textContent='Captured diagnostic snapshot; use Inspect again after other commands.\n'+JSON.stringify(await word.acceptanceSnapshot(),null,2);}catch(e){result.textContent=e.message;}}
  for(const [label,point] of cases){
    const b=button();b.textContent=label;
    b.onclick=()=>guarded(async()=>{word.clearAcceptance();try{await run(async()=>{
      await refresh();const {addModule}=await import('./model.mjs');
      const {next,item}=addModule(getProject(),`Synthetic ${label}`,'');
      await word.armAcceptance(point);await commit(next,'module',item);
    });}finally{word.clearAcceptance();}await snapshot();});box.append(b);
  }
  for(const when of ['before','after']){
    const b=button();b.textContent=`Settings failure ${when} duration write`;
    b.onclick=()=>guarded(async()=>{word.clearAcceptance();try{await run(async()=>{
      await refresh();const p=getProject(),a=p.modules[0]?.activities[0];
      if(!a)throw new Error('The disposable settings test needs an existing activity.');
      await word.armAcceptance({kind:'write-sync',when,nth:1});
      await word.updateActivitySettings(p,a.id,{minutes:a.minutes+2.5,lines:8});
      await refresh();
    });}finally{word.clearAcceptance();}await snapshot();});box.append(b);
  }
  const inspect=button();inspect.textContent='Inspect native acceptance state';inspect.onclick=()=>guarded(snapshot);box.append(inspect,result);
  document.querySelector('main').append(box);
}
