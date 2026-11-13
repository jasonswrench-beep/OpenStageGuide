import {ASSESSMENT_FILE_LIMITS,createAssessmentFile,serializeAssessmentFile,parseAssessmentFile,assessmentDraftFromDefinition,assessmentDraftIdentity,assessmentDraftEditorState} from './assessment-file.mjs';

// Manual PRIVATE authoring files only. An opaque workspace fingerprint can
// bind transient tester changes but is never serialized into a file. This
// controller has no scoring, Word, storage, network or consent-restore action.
const byteLength=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(Uint8Array.prototype),'byteLength').get;
const state=node=>({tag:node.tagName||'',hidden:!!node.hidden,class:node.className||'',style:node.getAttribute?.('style')||'',text:node.textContent,children:Array.from(node.children||[]).map(state)});
export function mountAssessmentFiles(document,{isBusy=()=>false,hasPending=()=>false,getFingerprint,getDraft,getDefinition,reserveDefinition,checkCandidate,applyDraft,requestReplacement,download,onStateChange=()=>{}}={}){
  const n=id=>document.getElementById('measure-file-'+id),view=document.getElementById('view-assessments');
  for(const id of ['section','notice','save-draft','save-definition','input','load','output','status','error'])if(!n(id))throw new Error(`Assessment file control missing: ${id}`);
  let busy=false,epoch=0,candidate=null,wasVisible=!view.hidden,wasHostBusy=!!isBusy();
  const fail=message=>{throw new Error(message);};
  const status=message=>{n('status').textContent=message;};
  const error=e=>{n('error').textContent=e.message||String(e);n('error').hidden=false;n('error').scrollIntoView?.({block:'nearest'});};
  const clearError=()=>{n('error').textContent='';n('error').hidden=true;};
  function shown(node){if(node.hidden||/(?:display\s*:\s*none|visibility\s*:\s*(?:hidden|collapse))/iu.test(node.getAttribute?.('style')||''))return false;if(typeof globalThis.getComputedStyle==='function'){try{const css=globalThis.getComputedStyle(node);if(css?.display==='none'||['hidden','collapse'].includes(css?.visibility))return false;}catch{return false;}}return true;}
  function currentVisible(){for(let p=n('section');p;p=p.parentElement||p.parent)if(!shown(p))return false;return shown(view);}
  function visible(){if(!currentVisible())fail('Open the Assessments PRIVATE file section before using its controls.');}
  function allowed(){visible();if(busy||isBusy())fail('Wait for the current operation; current assessment work was kept.');if(hasPending())fail('Keep or finish the pending replacement before using assessment files.');}
  function display(){const ancestors=[];for(let p=n('output');p;p=p.parentElement||p.parent){ancestors.push({tag:p.tagName||'',hidden:!!p.hidden,class:p.className||'',style:p.getAttribute?.('style')||''});if(p===view)break;}return JSON.stringify({output:state(n('output')),ancestors,viewHidden:!!view.hidden,section:state(n('notice'))});}
  const confirmation=()=>JSON.stringify({root:state(document.getElementById('measure-confirm')),text:state(document.getElementById('measure-confirm-text')),keep:state(document.getElementById('measure-keep')),replace:state(document.getElementById('measure-discard'))});
  const filename=kind=>`OpenStage-PRIVATE-Assessment-${kind==='draft'?'Draft':'Definition'}-${globalThis.crypto.randomUUID()}.json`;
  function save(kind){
    try{allowed();clearError();if(typeof download!=='function')fail('The local PRIVATE file download action is unavailable.');
      const before=getFingerprint(),file=createAssessmentFile(kind,kind==='draft'?getDraft():getDefinition()),bytes=serializeAssessmentFile(file);
      visible();if(isBusy()||hasPending()||before!==getFingerprint())fail('The assessment changed before download. No file was offered.');
      if(kind==='definition')reserveDefinition(file.payload);
      try{download(new Blob([new Uint8Array(bytes)],{type:'application/json;charset=utf-8'}),filename(kind));}catch(cause){throw new Error((kind==='definition'?'The local definition download outcome is unconfirmed; its ID/version is reserved in this pane. Repeat the exact definition or author a new version. ':'The local draft download outcome is unconfirmed. Current work was kept. ')+(cause.message||String(cause)));}
      status(`PRIVATE ${kind==='draft'?'editable authoring draft':'definition (checked format)'} offered locally. No tester responses, scores, export instructions or consent were saved. Confirm the download; nothing is saved automatically.`);
    }catch(e){error(e);}finally{renderBusy();}
  }
  n('save-draft').onclick=()=>save('draft');n('save-definition').onclick=()=>save('definition');
  n('input').onchange=()=>{epoch++;candidate=null;if(!n('input').files?.length){status('No replacement file selected. Current assessment work was kept.');renderBusy();return;}clearError();n('output').textContent='Selected PRIVATE file has not been checked. Check it before choosing Keep or Replace.';status('File selected locally. Nothing has been loaded or changed.');renderBusy();};
  n('input').oncancel=()=>{epoch++;candidate=null;status('File selection canceled. Current assessment work was kept.');renderBusy();};
  n('load').onclick=async()=>{
    let reader=null;
    try{allowed();clearError();const chosen=n('input').files;if(!chosen?.length){status('No replacement file selected. Current assessment work was kept.');return;}if(chosen.length!==1)fail('Choose exactly one PRIVATE assessment file.');
      const file=chosen[0],size=file.size;if(!Number.isSafeInteger(size)||size<=0||size>ASSESSMENT_FILE_LIMITS.maxFileBytes||typeof file.stream!=='function')fail('Choose a nonempty PRIVATE assessment JSON file up to 4 MiB with supported local reading.');
      const token=++epoch,source=getFingerprint();candidate=null;
      const guard=()=>{visible();if(token!==epoch||isBusy()||hasPending()||source!==getFingerprint()||n('input').files?.length!==1||n('input').files[0]!==file||file.size!==size)fail('The assessment, workspace or selected file changed while reading. Nothing was loaded; current work was kept.');};
      guard();busy=true;onStateChange();guard();reader=file.stream().getReader();const chunks=[];let total=0,reads=0;
      for(;;){guard();const part=await reader.read();guard();if(++reads>4096)fail('The local file has too many read chunks. Nothing was loaded.');if(part.done)break;if(!(part.value instanceof Uint8Array))fail('The local reader returned unsupported bytes.');const length=byteLength.call(part.value);if(length>size-total||length>ASSESSMENT_FILE_LIMITS.maxFileBytes-total)fail('The local file exceeded its declared or supported size. Nothing was loaded.');const copy=new Uint8Array(part.value);total+=copy.length;chunks.push(copy);}
      if(total!==size)fail('The local file was incomplete. Nothing was loaded.');const bytes=new Uint8Array(total);let at=0;for(const chunk of chunks){bytes.set(chunk,at);at+=chunk.length;}guard();
      const parsed=parseAssessmentFile(bytes),draft=parsed.kind==='draft'?parsed.payload:assessmentDraftFromDefinition(parsed.payload),identity=assessmentDraftIdentity(draft),editors=assessmentDraftEditorState(draft),warning=checkCandidate(draft,identity)||'';guard();
      const summary=[`PRIVATE ${parsed.kind==='draft'?'editable authoring draft':'definition (checked format)'} file`, `Assessment ID: ${draft.definitionId}; authored version text: ${draft.version}`,`Saved items: ${draft.items.length}; saved subscales: ${draft.rules.length}.`, `Active item: ${editors.itemClean?'matches a saved item or pristine new draft':'unfinished raw work retained'}. Active subscale: ${editors.ruleClean?'matches a saved rule or pristine new draft':'unfinished raw work retained'}.`,warning,'Replace restores editable authoring data only. It clears the captured tester and its synthetic responses, definition/export review and acknowledgements. Export instructions, copy choice and language remain current and need fresh review. No file-provided consent or scores are restored.','Exact PRIVATE authoring payload:',JSON.stringify(draft,null,2)].filter(Boolean).join('\n\n');
      n('output').textContent=summary;candidate={token,source,draft,warning,display:display()};busy=false;
      const intent=candidate;
      const fresh=(confirmed=false)=>{visible();if(candidate!==intent||intent.token!==epoch||isBusy()||intent.source!==getFingerprint()||intent.display!==display()||(!confirmed&&intent.confirm&&intent.confirm!==confirmation())||n('input').files?.length!==1||n('input').files[0]!==file||file.size!==size)fail('The checked PRIVATE file, current work or displayed replacement review changed. Nothing was replaced; check the file again.');};
      fresh();requestReplacement('Replace this assessment from the checked PRIVATE authoring file? Current definition/item/subscale drafts and the captured tester with its synthetic responses will be replaced. Keep leaves them unchanged.'+(warning?' '+warning:''),()=>{try{fresh(true);applyDraft(intent.draft,intent.warning,parsed.kind);candidate=null;n('output').textContent='PRIVATE authoring data loaded as an uncaptured editable draft. Review and capture again before testing or Word export.';status('PRIVATE authoring data loaded. No response, score or consent was restored. Current export instructions, copy choice and language were retained.');}catch(e){candidate=null;n('output').textContent='PRIVATE file was not loaded. Current authoring work was retained; check the file again before another replacement attempt.';status('PRIVATE replacement failed without accepting the file. Current assessment drafts were retained.');throw e;}},{guard:fresh,onKeep:()=>{candidate=null;n('output').textContent='Checked file was kept separate. Current assessment drafts and responses were kept.';},replaceText:'Replace assessment from PRIVATE file'});intent.confirm=confirmation();
      status('PRIVATE file checked. Review the exact authoring payload and choose Keep or Replace. Current assessment work has not changed.');
    }catch(e){candidate=null;error(e);}finally{busy=false;if(reader){try{await reader.cancel();}catch{}try{reader.releaseLock?.();}catch{}}onStateChange();renderBusy();}
  };
  function renderBusy(){const visibleNow=currentVisible(),hostBusy=!!isBusy();if(visibleNow!==wasVisible||hostBusy!==wasHostBusy){epoch++;wasVisible=visibleNow;wasHostBusy=hostBusy;}
    const blocked=busy||hostBusy||!!hasPending();for(const id of ['save-draft','save-definition','input','load'])n(id).disabled=blocked;
  }
  renderBusy();return Object.freeze({renderBusy,isBusy:()=>busy});
}
