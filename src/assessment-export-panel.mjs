import {prepareAssessmentExport,assessmentExportBlob,assessmentExportName} from './assessment-output.mjs';

const INVALID=['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing'];
const raw=n=>({value:n.value,invalid:INVALID.map(k=>!!n.validity?.[k])});
const state=n=>({tag:n.tagName||'',hidden:!!n.hidden,class:n.className||'',style:n.getAttribute?.('style')||'',text:n.textContent,children:Array.from(n.children||[]).map(state)});
// Session-only, definition-only export controls. This module does not read the
// response tester, call the scorer, write Word or persist authoring state.
export function mountAssessmentExports(document,{getSource,isBusy=()=>false,download,getZip=()=>globalThis.JSZip,now=()=>new Date().toISOString(),packageFiles=assessmentExportBlob}={}){
  const n=id=>document.getElementById('measure-export-'+id),view=document.getElementById('view-assessments');
  const fields=['route','instructions','language'],buttons=['preview','download'];
  let reviewed=null,consent=null,privateConsent=null,ackEpoch=0,privateEpoch=0,busy=false,epoch=0;
  const fail=m=>{throw new Error(m);};
  const status=s=>{n('status').textContent=s;};
  const error=e=>{n('error').hidden=false;n('error').textContent=e.message||String(e);n('error').scrollIntoView?.({block:'nearest'});};
  const clearError=()=>{n('error').hidden=true;n('error').textContent='';};
  const visible=()=>{if(view.hidden||n('section').hidden)fail('Open the Assessments export section before reviewing or downloading.');};
  function options(stamp){const values=Object.fromEntries(fields.map(k=>[k,raw(n(k))]));if(Object.values(values).some(v=>v.invalid.some(Boolean)))fail('An export field is invalid. Your draft was kept.');return {route:values.route.value,instructions:values.instructions.value,baseLanguage:values.language.value,stamp};}
  function source(){visible();if(isBusy())fail('Wait for the current Word operation. No export was offered.');return getSource();}
  function sourceKey(){const s=source();return JSON.stringify({fingerprint:s.fingerprint,definition:s.definition,options:fields.map(k=>raw(n(k)))});}
  function display(){const ancestors=[];for(let p=n('output');p;p=p.parentElement||p.parent){ancestors.push({tag:p.tagName||'',hidden:!!p.hidden,class:p.className||'',style:p.getAttribute?.('style')||''});if(p===view)break;}return JSON.stringify({output:state(n('output')),ancestors,section: {hidden:!!n('section').hidden,class:n('section').className||'',style:n('section').getAttribute?.('style')||''},viewHidden:!!view.hidden,reviewLabel:state(n('review-label')),privateLabel:state(n('private-label'))});}
  function invalidate(message='Export choices or definition changed. Refresh the exact export review.'){epoch++;reviewed=null;consent=null;privateConsent=null;n('reviewed').checked=false;n('private').checked=false;n('output').textContent=message;}
  function fresh(){if(!reviewed||reviewed.source!==sourceKey()||reviewed.display!==display())fail('Definition, unfinished draft, export choices or displayed review changed. Refresh and review again.');const regenerated=prepareAssessmentExport(source().definition,reviewed.options);if(JSON.stringify(regenerated.files)!==reviewed.files||JSON.stringify(regenerated.preview.paragraphs)!==reviewed.paragraphs)fail('The reviewed package content changed. Refresh and review again.');return reviewed;}
  function review(){if(busy)fail('An export is already being assembled.');const s=source(),opts=options(now()),prepared=prepareAssessmentExport(s.definition,opts);invalidate();n('private-label').hidden=opts.route!=='author';n('output').textContent=prepared.preview.paragraphs.map(p=>p.text).join('\n\n');reviewed={source:sourceKey(),display:display(),definition:s.definition,options:opts,files:JSON.stringify(prepared.files),paragraphs:JSON.stringify(prepared.preview.paragraphs)};status(opts.route==='author'?'PRIVATE guide review ready. Review its keys and affirm both acknowledgements.':'Blank respondent review ready. No test response or score is included.');}
  const act=fn=>()=>{try{visible();if(busy)fail('Wait for the current export assembly.');clearError();fn();}catch(e){error(e);}finally{renderBusy();}};
  n('preview').onclick=act(review);
  const acknowledge=act(()=>{if(n('reviewed').checked){const r=fresh();consent={review:r,epoch:ackEpoch};}else consent=null;});
  n('reviewed').onchange=()=>{ackEpoch++;acknowledge();};
  const acknowledgePrivate=act(()=>{if(n('private').checked){const r=fresh();if(r.options.route!=='author')fail('Private consent applies only to the author guide.');privateConsent={review:r,epoch:privateEpoch};}else privateConsent=null;});
  n('private').onchange=()=>{privateEpoch++;acknowledgePrivate();};
  n('download').onclick=async()=>{
    if(busy){error(new Error('An export is already being assembled.'));return;}
    try{clearError();const r=fresh();if(!n('reviewed').checked||consent?.review!==r)fail('Acknowledge this exact reviewed export first.');if(r.options.route==='author'&&(!n('private').checked||privateConsent?.review!==r))fail('Affirm the separate PRIVATE guide acknowledgement first.');if(typeof download!=='function')fail('The local file download action is unavailable.');
      const token=epoch,ack=ackEpoch,priv=privateEpoch;
      const guard=()=>{if(token!==epoch||ack!==ackEpoch||priv!==privateEpoch||fresh()!==r||!n('reviewed').checked||consent?.review!==r||(r.options.route==='author'&&(!n('private').checked||privateConsent?.review!==r)))fail('Export review changed during assembly. No file was offered; drafts were kept.');};
      guard();busy=true;renderBusy();const blob=await packageFiles(r.definition,r.options,getZip());guard();
      // The existing strict ZIP helper has already made/validated private byte
      // copies. Take another offered snapshot before the final freshness check.
      const bytes=Uint8Array.from(new Uint8Array(await blob.arrayBuffer()));guard();
      const offered=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});guard();download(offered,assessmentExportName(r.options.route));status('New reviewed copy offered locally. Test responses and scores were excluded; review Word layout and accessibility before use.');
    }catch(e){error(e);}finally{busy=false;renderBusy();}
  };
  for(const id of fields){const change=()=>{invalidate();clearError();n('private-label').hidden=n('route').value!=='author';renderBusy();};n(id).oninput=change;n(id).onchange=change;}
  function renderBusy(){if(isBusy()||view.hidden){if(busy)epoch++;}if(reviewed){let valid=false;try{valid=reviewed.source===sourceKey()&&reviewed.display===display();}catch{}if(!valid)invalidate();}
    const blocked=busy||!!isBusy();for(const id of fields)n(id).disabled=blocked;for(const id of buttons)n(id).disabled=blocked;n('reviewed').disabled=blocked||!reviewed;n('private').disabled=blocked||!reviewed||n('route').value!=='author';n('download').disabled=blocked||!reviewed||!consent||!n('reviewed').checked||(n('route').value==='author'&&(!privateConsent||!n('private').checked));
  }
  n('route').value='respondent';n('language').value='en-US';n('private-label').hidden=true;
  renderBusy();
  return Object.freeze({renderBusy,hasSessionWork:()=>!!n('instructions').value||n('route').value!=='respondent'||n('language').value!=='en-US',invalidate});
}
