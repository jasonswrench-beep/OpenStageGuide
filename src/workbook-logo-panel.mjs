import {rawControl} from './activity-draft.mjs';
import {prepareWorkbookLogo,workbookLogoInfo,workbookLogoBase64,copyWorkbookLogoBytes,WORKBOOK_LOGO_LIMITS} from './workbook-logo.mjs';

// One local, session-only verified image. Nothing is saved to Word, presets or a service.
export function mountWorkbookLogoPanel(document,{isUnavailable=()=>false,onChange=()=>{},context=()=>'',prepare=prepareWorkbookLogo}={}){
  const n=id=>document.getElementById('family-logo-'+id);
  let asset=null,lastValid=null,state='empty',generation=0,pending=null,chosenFile=null,chosenValue='',loading=false;
  const tokens=new WeakMap();let token=0;
  const fileToken=f=>{if(!f||typeof f!=='object')return null;if(!tokens.has(f))tokens.set(f,++token);return tokens.get(f);};
  const current=()=>state==='ready'&&!!asset&&(!chosenFile||!n('file').files?.[0]||n('file').files?.[0]===chosenFile);
  const raw=()=>({schemaVersion:1,include:{checked:!!n('include').checked,...rawControl(n('include'))},role:rawControl(n('role')),alt:rawControl(n('alt')),widthIn:rawControl(n('width')),selection:{state,generation,asset,current:current()}});
  const identity=()=>({generation,state,loading,asset:asset?workbookLogoInfo(asset):null,fileValue:n('file').files?.[0]?n('file').value:chosenValue,fileToken:fileToken(n('file').files?.[0]||chosenFile),chosen:fileToken(chosenFile)});
  const imageState=node=>({tag:node?.tagName||'',hidden:!!node?.hidden,class:node?.className||'',style:node?.getAttribute?.('style')||node?.style?.cssText||'',src:node?.getAttribute?.('src')??node?.src??'',alt:node?.getAttribute?.('alt')??node?.alt??'',width:node?.getAttribute?.('width')??node?.width??'',height:node?.getAttribute?.('height')??node?.height??''});
  const shape=node=>({tag:node?.tagName||'',hidden:!!node?.hidden,class:node?.className||'',style:node?.getAttribute?.('style')||node?.style?.cssText||'',text:node?.textContent||'',children:Array.from(node?.children||[],shape)});
  const displayState=()=>({image:imageState(n('image')),controls:shape(n('controls'))});
  const confirmation=()=>JSON.stringify({root:shape(n('confirm')),text:shape(n('confirm-text')),keep:shape(n('keep')),replace:shape(n('replace'))});
  const fingerprint=()=>JSON.stringify({raw:raw(),identity:identity(),context:context(),display:displayState()});
  const visible=target=>{if(!target||target.isConnected===false)return false;for(let node=target;node;node=node.parentElement||node.parent){if(node.hidden||/display\s*:\s*none|visibility\s*:\s*(?:hidden|collapse)/u.test(node.getAttribute?.('style')||node.style?.cssText||''))return false;try{if(typeof globalThis.getComputedStyle==='function'&&node.nodeType===1){const css=globalThis.getComputedStyle(node);if(css.display==='none'||['hidden','collapse'].includes(css.visibility))return false;}}catch{}}return true;};
  const ready=target=>{if(document.getElementById('view-output')?.hidden||!visible(n('controls'))||!visible(target))throw Error('Open the workbook logo controls before changing this draft.');if(isUnavailable()||loading)throw Error('Finish the current operation or replacement choice before changing this logo.');};
  const status=text=>{n('status').textContent=text;};
  const report=e=>{n('error').hidden=false;n('error').textContent=e.message||String(e);};
  const clear=()=>{n('error').hidden=true;n('error').textContent='';};
  const clearFile=()=>{n('file').value='';chosenFile=null;chosenValue='';};
  function thumbnail(){
    const image=n('image');
    if(!current()){image.hidden=true;if(image.removeAttribute)image.removeAttribute('src');else{delete image.attributes?.src;image.src='';}return;}
    const info=workbookLogoInfo(asset),scale=Math.min(160/info.width,100/info.height,1);
    const attrs={src:'data:image/png;base64,'+workbookLogoBase64(asset),alt:n('alt').value.trim()?n('alt').value:'Local logo preview; author an alternative description before including it.',width:String(Math.max(1,Math.round(info.width*scale))),height:String(Math.max(1,Math.round(info.height*scale)))};
    for(const [key,value] of Object.entries(attrs))image.setAttribute(key,value);image.hidden=false;
  }
  const dismiss=()=>{const focus=pending?.returnFocus;pending=null;n('confirm').hidden=true;n('confirm-text').textContent='';if(focus&&!focus.disabled&&visible(focus))focus.focus?.();else n('status').focus?.();};
  const action=(fn,target)=>()=>{try{ready(target);clear();fn();}catch(e){report(e);}finally{renderBusy();}};
  function request(kind,apply,text){
    if(pending)throw Error('Keep or replace the pending logo choice first.');
    pending={kind,apply,returnFocus:document.activeElement||document.focused||null};
    n('confirm-text').textContent=text;n('confirm').hidden=false;n('keep').focus?.();onChange(false);
    pending.fingerprint=fingerprint();pending.display=confirmation();
  }
  n('keep').onclick=action(()=>{if(!pending)throw Error('No logo replacement is pending.');const image=pending.kind==='image';dismiss();status(image?'The new image was not applied. The earlier verified image remains available only through explicit Restore; role, description, width and inclusion drafts were kept.':'The current logo draft was kept.');onChange(false);},n('keep'));
  n('replace').onclick=action(()=>{if(!pending)throw Error('No logo replacement is pending.');if(pending.fingerprint!==fingerprint()||pending.display!==confirmation()){dismiss();onChange(false);throw Error('The logo, source, controls or displayed confirmation changed. Nothing was replaced; start the intended action again.');}const apply=pending.apply;dismiss();apply();generation++;thumbnail();onChange(true);},n('replace'));
  n('reset').onclick=action(()=>request('reset',()=>{asset=lastValid=null;state='empty';clearFile();n('include').checked=false;n('role').value='';n('alt').value='';n('width').value='1.5';status('Only the session logo and its inclusion, role, description and width were reset.');},'Reset only the session logo, inclusion, role, description and width? Course, branding, public text, research and other drafts are kept.'),n('reset'));
  n('remove').onclick=action(()=>request('remove',()=>{asset=lastValid=null;state='empty';clearFile();n('include').checked=false;status('The image was removed and inclusion turned off; role, description and width drafts were kept.');},'Remove this session image and turn logo inclusion off? Its description, role and width drafts are kept.'),n('remove'));
  n('restore').onclick=action(()=>{if(!lastValid)throw Error('There is no earlier verified image in this session.');request('restore',()=>{asset=lastValid;state='ready';clearFile();status('The last verified image was explicitly restored. Review its description, size and inclusion before exporting.');},'Restore the last verified PNG explicitly? The failed selection is cleared; description, role, width and inclusion are kept.');},n('restore'));
  const bindFileInput=target=>{target.onchange=async()=>{
    if(target!==n('file')){report(Error('This logo chooser was replaced. Use the current file control.'));return;}
    // Cancel has no replacement effect. Native File objects are immutable; retain the
    // selected object, reader and byte limit before the first asynchronous boundary.
    const file=n('file').files?.[0];if(!file)return;let run=null,admitted=false;
    try{ready(n('file'));if(pending)throw Error('Keep or replace the pending logo choice first.');clear();admitted=true;
      const size=file.size,stream=file.stream;
      if(typeof size!=='number'||!Number.isSafeInteger(size)||size<=0||size>WORKBOOK_LOGO_LIMITS.bytes||typeof stream!=='function')throw Error('Select one stream-readable PNG up to 2 MiB. The earlier verified image is retained only for explicit Restore.');
      const openStream=stream.bind(file),previous=asset||lastValid;run=++generation;
      lastValid=previous;asset=null;chosenFile=file;chosenValue=n('file').value;state='loading';loading=true;thumbnail();onChange(true);renderBusy();
      const snapshot=JSON.stringify({context:context(),raw:raw(),identity:identity()});
      const still=()=>generation===run&&n('file').files?.[0]===file&&JSON.stringify({context:context(),raw:raw(),identity:identity()})===snapshot&&!isUnavailable()&&visible(n('controls'))&&!document.getElementById('view-output')?.hidden;
      const reader=openStream().getReader(),parts=[];let total=0,done=false;
      try{for(let reads=0;reads<4096;reads++){
        if(!still())throw Error('Source, logo drafts or selection changed while reading. No new image was selected; the earlier verified image can be explicitly restored.');
        const item=await reader.read();if(!still())throw Error('Source, logo drafts or selection changed while reading. No new image was selected; the earlier verified image can be explicitly restored.');
        if(item.done){done=true;break;}const part=copyWorkbookLogoBytes(item.value);total+=part.length;
        if(total>size||total>WORKBOOK_LOGO_LIMITS.bytes)throw Error('The local image stream exceeded its declared or supported byte limit.');parts.push(part);
      }if(!done||total!==size)throw Error('The local image stream was incomplete or exceeded the bounded read limit.');}
      finally{if(!done){try{Promise.resolve(reader.cancel()).catch(()=>{});}catch{}}try{reader.releaseLock?.();}catch{}}
      const bytes=new Uint8Array(total);let at=0;for(const part of parts){bytes.set(part,at);at+=part.length;}
      const candidate=await prepare(bytes);if(!still())throw Error('Source, logo drafts or selection changed while validating. No new image was selected; the earlier verified image can be explicitly restored.');
      loading=false;
      if(previous){state='invalid';asset=null;chosenFile=null;clearFile();request('image',()=>{asset=lastValid=candidate;state='ready';clearFile();status('The newly verified PNG replaced the prior image. Review its authored description and actual fitted size.');},'Replace the prior verified session image with this newly verified PNG? Keep retains the prior image only for explicit Restore; its role, description, width and inclusion drafts stay unchanged.');}
      else{asset=lastValid=candidate;state='ready';thumbnail();status('PNG verified locally. Choose informative, author its description, enable the public cover and explicitly include the logo. Original pixels are not recolored, including in grayscale print mode.');onChange(true);}
    }catch(e){if(admitted&&(run===null||generation===run)){loading=false;lastValid=asset||lastValid;asset=null;chosenFile=file;state='invalid';generation++;thumbnail();onChange(true);report(e);}else if(!admitted)report(e);}finally{renderBusy();}
  };};bindFileInput(n('file'));
  const changed=()=>{clear();thumbnail();onChange(true);renderBusy();};
  for(const id of ['include','role','alt','width'])n(id).oninput=n(id).onchange=changed;
  n('confirm').onkeydown=e=>{if(e.key==='Escape'){e.preventDefault?.();n('keep').onclick();}};
  function renderBusy(){const blocked=!!isUnavailable()||loading;for(const id of ['include','file','role','alt','width'])n(id).disabled=blocked||!!pending;for(const id of ['reset','remove'])n(id).disabled=blocked||!!pending;n('restore').disabled=blocked||!!pending||!lastValid;n('keep').disabled=n('replace').disabled=blocked||!pending;}
  n('include').checked=false;n('role').value='';n('alt').value='';n('width').value='1.5';n('confirm').hidden=true;thumbnail();clear();status('Logo selection and drafts are session-only unless explicitly saved in a design bundle; they are not stored automatically in the course or brand presets. PNG only; decorative export, JPEG/SVG and color-profile conversion are unavailable.');renderBusy();

  // Opaque, reversible design-file transaction. The original native file input
  // node (and its immutable FileList) is retained rather than being cleared.
  const designTokens=new WeakMap();
  function designFileAsset(){
    if(loading||pending)throw Error('Finish or Keep the current logo operation before saving a design bundle.');
    if(state==='invalid'||state==='loading'||state==='ready'&&!current())throw Error('The current logo selection is not ready. Explicitly Restore, replace or remove it before saving; no image was silently omitted.');
    if(state==='empty'&&(asset||lastValid))throw Error('An earlier image is retained only for explicit Restore. Resolve that choice before saving.');
    return state==='ready'?asset:null;
  }
  function snapshotDesignFile(){
    if(loading||pending)throw Error('Finish or Keep the current logo operation before replacing a design bundle.');
    const input=n('file'),parent=input.parentNode||input.parentElement||input.parent;
    if(!parent||typeof parent.replaceChild!=='function')throw Error('This editor cannot safely replace the logo chooser. Current logo and file selection were kept.');
    const token=Object.freeze({kind:'OpenStageLogoDesignTransaction'});
    designTokens.set(token,{asset,lastValid,state,generation,chosenFile,chosenValue,loading,pending,input,parent,children:Array.from(parent.childNodes||parent.children||[]),raw:{include:n('include').checked,role:n('role').value,alt:n('alt').value,widthIn:n('width').value},image:{hidden:n('image').hidden,attributes:['src','alt','width','height','style','class'].map(key=>[key,n('image').getAttribute?.(key)])},status:n('status').textContent,error:n('error').textContent,errorHidden:n('error').hidden});return token;
  }
  function replaceDesignFile(token,raw,preparedAsset){
    const before=designTokens.get(token);if(!before||before.input!==n('file')||loading||pending)throw Error('The logo transaction changed. No bundle was applied.');
    if(preparedAsset!==null)workbookLogoInfo(preparedAsset);
    const replacement=document.createElement('input');replacement.type='file';replacement.setAttribute('type','file');
    for(const key of ['accept','class','aria-label','aria-describedby','tabindex','name','title']){const value=before.input.getAttribute?.(key);if(value!==null&&value!==undefined)replacement.setAttribute(key,value);}
    replacement.disabled=before.input.disabled;bindFileInput(replacement);replacement.id=before.input.id;before.parent.replaceChild(replacement,before.input);
    asset=lastValid=preparedAsset;state=preparedAsset?'ready':'empty';chosenFile=null;chosenValue='';loading=false;pending=null;generation++;
    n('include').checked=raw.include;n('role').value=raw.role;n('alt').value=raw.alt;n('width').value=raw.widthIn;
    thumbnail();clear();status('Design bundle logo restored under current PNG policy. Review the current thumbnail, description, inclusion and size before exporting.');renderBusy();
  }
  function rollbackDesignFile(token){
    const before=designTokens.get(token);if(!before)throw Error('The earlier logo transaction is unavailable.');
    before.parent.replaceChildren(...before.children);before.input.id=before.input.id;
    ({asset,lastValid,state,generation,chosenFile,chosenValue,loading,pending}=before);
    n('include').checked=before.raw.include;n('role').value=before.raw.role;n('alt').value=before.raw.alt;n('width').value=before.raw.widthIn;
    for(const[key,value]of before.image.attributes){if(value===null||value===undefined)n('image').removeAttribute?.(key);else n('image').setAttribute(key,value);}n('image').hidden=before.image.hidden;
    n('status').textContent=before.status;n('error').textContent=before.error;n('error').hidden=before.errorHidden;designTokens.delete(token);renderBusy();
  }
  function finishDesignFile(token){if(!designTokens.has(token))throw Error('The logo transaction is unavailable.');designTokens.delete(token);}

  const hasSessionWork=()=>!!asset||!!lastValid||loading||!!pending||n('include').checked||n('role').value!==''||n('alt').value!==''||n('width').value!=='1.5';
  const beforeUnload=e=>{if(hasSessionWork()){e.preventDefault?.();e.returnValue='';return '';}};globalThis.addEventListener?.('beforeunload',beforeUnload);
  return Object.freeze({raw,identity,displayState,designFileAsset,snapshotDesignFile,replaceDesignFile,rollbackDesignFile,finishDesignFile,isPending:()=>!!pending,isLoading:()=>loading,renderBusy,hasSessionWork,beforeUnload,reportPending:()=>report(Error('Keep or replace the pending logo choice before reviewing or applying.'))});
}
