// Verified, session-only PNG cover assets. No Word, course, storage or network.
export const WORKBOOK_LOGO_VERSION=1;
export const WORKBOOK_LOGO_MEDIA_PART='word/media/organization-logo.png';
export const WORKBOOK_LOGO_LIMITS=Object.freeze({bytes:2*1024*1024,dimension:2048,pixels:4194304,chunks:256,altTextUnits:1000,minWidthIn:0.5,maxWidthIn:3,maxHeightIn:2});
export const WORKBOOK_LOGO_VALIDITY=Object.freeze(['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing']);
export const EMPTY_WORKBOOK_LOGO=Object.freeze({schemaVersion:1,include:false,asset:null,role:'',alt:'',widthIn:null});
const assets=new WeakMap(),fail=message=>{throw new Error(`Cover logo: ${message}`);};
const typed=Object.getPrototypeOf(Uint8Array.prototype),byteLength=Object.getOwnPropertyDescriptor(typed,'byteLength').get,byteOffset=Object.getOwnPropertyDescriptor(typed,'byteOffset').get,bufferOf=Object.getOwnPropertyDescriptor(typed,'buffer').get;
const freeze=value=>{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value);}return value;};
const own=(value,keys,label)=>{if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be an own-data record.`);const found=Reflect.ownKeys(value);if(found.length!==keys.length||found.some(k=>typeof k!=='string'||!keys.includes(k)))fail(`${label} fields are unsupported.`);for(const k of keys){const d=Object.getOwnPropertyDescriptor(value,k);if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${k} must be enumerable data.`);}};
const xmlText=(value,max,label)=>{if(typeof value!=='string'||value.length>max)fail(`${label} must be bounded text.`);for(const s of value){const cp=s.codePointAt(0);if(!(cp===9||cp===10||cp===13||cp>=0x20&&cp<=0xd7ff||cp>=0xe000&&cp<=0xfffd||cp>=0x10000&&cp<=0x10ffff))fail(`${label} contains unsupported XML text.`);}return value.replace(/\r\n?/gu,'\n');};
export function copyWorkbookLogoBytes(value){
  if(arguments.length!==1)fail('provide one PNG byte array.');
  if(!value||typeof value!=='object')fail('provide ordinary PNG bytes, not a custom byte object.');
  let size,source;try{size=byteLength.call(value);source=bufferOf.call(value);byteOffset.call(value);}catch{fail('PNG bytes are unavailable.');}
  if(Object.getPrototypeOf(value)!==Uint8Array.prototype||Object.getOwnPropertySymbols(value).length||['constructor','buffer','byteOffset','byteLength','length'].some(k=>Object.hasOwn(value,k)))fail('provide ordinary PNG bytes, not a custom byte object.');
  if(Object.getPrototypeOf(source)!==ArrayBuffer.prototype||size>WORKBOOK_LOGO_LIMITS.bytes)fail('PNG bytes exceed the limit or use unsupported shared storage.');
  const out=new Uint8Array(size);Uint8Array.prototype.set.call(out,value);return out;
}
const crcTable=Array.from({length:256},(_,n)=>{let c=n;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0;});
const crc=(bytes,start=0,end=bytes.length)=>{let c=0xffffffff;for(let i=start;i<end;i++)c=crcTable[(c^bytes[i])&255]^(c>>>8);return (c^0xffffffff)>>>0;};
const sig=new Uint8Array([137,80,78,71,13,10,26,10]);
function join(parts){const out=new Uint8Array(parts.reduce((n,p)=>n+p.length,0));let at=0;for(const p of parts){out.set(p,at);at+=p.length;}return out;}
function chunk(type,data){const out=new Uint8Array(12+data.length),view=new DataView(out.buffer);view.setUint32(0,data.length);for(let i=0;i<4;i++)out[4+i]=type.charCodeAt(i);out.set(data,8);view.setUint32(out.length-4,crc(out,4,out.length-4));return out;}
function parsePng(bytes){
  if(bytes.length<45||sig.some((v,i)=>bytes[i]!==v))fail('choose an actual supported PNG file.');
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),idats=[];let at=8,count=0,ihdr=null,width=0,height=0,channels=0,srgb=null,seenPhys=false,seenIdat=false,endedIdat=false,ended=false;
  const strip=new Set(['tEXt','zTXt','iTXt','eXIf','tIME','pHYs']);
  while(at<bytes.length){
    if(++count>WORKBOOK_LOGO_LIMITS.chunks||at+12>bytes.length)fail('PNG chunk inventory is unsupported.');
    const len=view.getUint32(at),end=at+12+len;if(end>bytes.length||len>WORKBOOK_LOGO_LIMITS.bytes)fail('PNG chunk length is invalid.');
    const type=String.fromCharCode(...bytes.subarray(at+4,at+8));if(!/^[A-Za-z]{4}$/u.test(type)||type[2]!==type[2].toUpperCase()||view.getUint32(end-4)!==crc(bytes,at+4,end-4))fail('PNG chunk type or CRC is invalid.');
    const data=bytes.slice(at+8,end-4);
    if(!ihdr&&type!=='IHDR')fail('PNG must start with IHDR.');
    if(type==='IHDR'){
      if(ihdr||at!==8||len!==13)fail('PNG IHDR is invalid.');ihdr=data;const h=new DataView(data.buffer);width=h.getUint32(0);height=h.getUint32(4);channels=data[9]===2?3:data[9]===6?4:0;
      if(!width||!height||width>WORKBOOK_LOGO_LIMITS.dimension||height>WORKBOOK_LOGO_LIMITS.dimension||width*height>WORKBOOK_LOGO_LIMITS.pixels||data[8]!==8||!channels||data[10]!==0||data[11]!==0||data[12]!==0)fail('use a noninterlaced8-bit RGB/RGBA PNG within2048 pixels per dimension.');
    }else if(type==='IDAT'){
      if(endedIdat||!len)fail('PNG IDAT data must be nonempty and contiguous.');seenIdat=true;idats.push(data);
    }else if(type==='IEND'){
      if(!seenIdat||len!==0||end!==bytes.length)fail('PNG IEND or trailing data is invalid.');ended=true;at=end;break;
    }else{
      if(seenIdat)endedIdat=true;
      if(type==='sRGB'){if(srgb||seenIdat||len!==1||data[0]>3)fail('PNG sRGB metadata is invalid.');srgb=data;}
      else if(['iCCP','gAMA','cHRM'].includes(type))fail('this color profile is unsupported; convert the logo to an untagged/default-sRGB or standard sRGB PNG.');
      else if(['acTL','fcTL','fdAT'].includes(type))fail('animated PNG logos are unsupported.');
      else if(type==='pHYs'){if(seenPhys||seenIdat||len!==9||data[8]>1)fail('PNG physical geometry is unsupported.');seenPhys=true;const pv=new DataView(data.buffer);if(pv.getUint32(0)!==pv.getUint32(4)||data[8]===1&&!pv.getUint32(0))fail('PNG nonsquare-pixel geometry is unsupported; use square pixels.');}
      else if(!strip.has(type))fail(`PNG chunk ${type} is unsupported; use a simple RGB/RGBA PNG.`);
    }
    at=end;
  }
  if(!ended||at!==bytes.length)fail('PNG ending is incomplete.');
  return {width,height,channels,ihdr,srgb,compressed:join(idats),decodedBytes:height*(1+width*channels)};
}
async function transform(bytes,Constructor,limit,label){
  if(typeof Constructor!=='function')fail('this runtime cannot safely process PNG compression. Keep your draft and use a supported runtime.');
  let stream;try{stream=new Constructor('deflate');}catch{fail('this runtime cannot safely process PNG compression.');}
  const reader=stream.readable.getReader(),writer=stream.writable.getWriter();let writeError=null;
  const writing=(async()=>{try{await writer.write(bytes);await writer.close();}catch(e){writeError=e;try{await writer.abort(e);}catch{}}})();
  const parts=[];let size=0;
  try{for(;;){const item=await reader.read();if(item.done)break;if(!(item.value instanceof Uint8Array))fail(`${label} returned unsupported bytes.`);size+=item.value.length;if(size>limit)fail(`${label} exceeds its bounded size.`);parts.push(new Uint8Array(item.value));}await writing;if(writeError)fail(`${label} failed.`);return join(parts);}
  catch(e){try{await reader.cancel(e);}catch{}try{await writer.abort(e);}catch{}await writing;if(String(e.message||e).startsWith('Cover logo:'))throw e;fail(`${label} is invalid.`);}
  finally{try{reader.releaseLock();writer.releaseLock();}catch{}}
}
export async function prepareWorkbookLogo(input){
  if(arguments.length!==1)fail('provide one local PNG byte array.');
  const bytes=copyWorkbookLogoBytes(input),p=parsePng(bytes);
  const scanlines=await transform(p.compressed,globalThis.DecompressionStream,p.decodedBytes,'PNG decoded data');
  if(scanlines.length!==p.decodedBytes)fail('PNG decoded scanline length is invalid.');
  const stride=1+p.width*p.channels;for(let at=0;at<scanlines.length;at+=stride)if(scanlines[at]>4)fail('PNG row filter is invalid.');
  const compressed=await transform(scanlines,globalThis.CompressionStream,WORKBOOK_LOGO_LIMITS.bytes,'PNG normalized data');
  const roundTrip=await transform(new Uint8Array(compressed),globalThis.DecompressionStream,p.decodedBytes,'PNG normalized readback');
  if(roundTrip.length!==scanlines.length||roundTrip.some((v,i)=>v!==scanlines[i]))fail('normalized PNG pixels changed during compression.');
  const normalized=join([sig,chunk('IHDR',p.ihdr),...(p.srgb?[chunk('sRGB',p.srgb)]:[]),chunk('IDAT',compressed),chunk('IEND',new Uint8Array())]);
  if(normalized.length>WORKBOOK_LOGO_LIMITS.bytes)fail('normalized PNG exceeds2MiB.');
  if(!globalThis.crypto?.subtle||typeof globalThis.crypto.randomUUID!=='function')fail('verified PNG identity is unavailable in this runtime.');
  const digest=new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256',new Uint8Array(normalized))),sha256=Array.from(digest,b=>b.toString(16).padStart(2,'0')).join('');
  const handle=Object.freeze({kind:'OpenStageWorkbookLogo',schemaVersion:1,id:globalThis.crypto.randomUUID()});
  assets.set(handle,{bytes:new Uint8Array(normalized),width:p.width,height:p.height,sha256});return handle;
}
function asset(handle){const value=handle&&typeof handle==='object'?assets.get(handle):null;if(!value)fail('select or explicitly restore a verified PNG from this session.');return value;}
export function workbookLogoInfo(handle){if(arguments.length!==1)fail('provide one verified logo.');const a=asset(handle);return Object.freeze({id:handle.id,sha256:a.sha256,width:a.width,height:a.height,bytes:a.bytes.length});}
export function workbookLogoBytes(handle){if(arguments.length!==1)fail('provide one verified logo.');return new Uint8Array(asset(handle).bytes);}
export function workbookLogoBase64(handle){if(arguments.length!==1)fail('provide one verified logo.');const b=workbookLogoBytes(handle);let value='';for(let at=0;at<b.length;at+=8192)value+=String.fromCharCode(...b.subarray(at,at+8192));return btoa(value);}
export function validateWorkbookLogo(value){
  if(arguments.length!==1)fail('provide one logo choice.');own(value,['schemaVersion','include','asset','role','alt','widthIn'],'logo choice');
  if(value.schemaVersion!==1||typeof value.include!=='boolean'||!['','informative','decorative'].includes(value.role))fail('logo version, inclusion or role is unsupported.');
  if(value.asset!==null)asset(value.asset);const alt=xmlText(value.alt,WORKBOOK_LOGO_LIMITS.altTextUnits,'alternative description');
  if(value.widthIn!==null&&(typeof value.widthIn!=='number'||!Number.isFinite(value.widthIn)||Object.is(value.widthIn,-0)||value.widthIn<0.5||value.widthIn>3))fail('logo width must be0.5-3 inches.');
  if(value.include){if(!value.asset)fail('select a verified PNG before including the logo.');if(value.role==='decorative')fail('decorative image export is not supported yet; choose informative with an authored description or leave the logo off.');if(value.role!=='informative'||!alt.trim())fail('choose informative and write a meaningful alternative description.');if(value.widthIn===null)fail('enter a valid logo width.');}
  return Object.freeze({schemaVersion:1,include:value.include,asset:value.asset,role:value.role,alt,widthIn:value.widthIn});
}
export function workbookLogoPresentation(value,frame){
  if(arguments.length!==2)fail('provide a logo choice and page geometry.');const v=validateWorkbookLogo(value);own(frame,['coverIncluded','pageWidthTwips','marginTwips'],'logo frame');
  if(typeof frame.coverIncluded!=='boolean'||!Number.isInteger(frame.pageWidthTwips)||!Number.isInteger(frame.marginTwips)||frame.marginTwips<0||frame.pageWidthTwips<=2*frame.marginTwips)fail('page geometry is unsupported.');
  if(!v.include)return null;if(!frame.coverIncluded)fail('enable the public cover before including its logo.');const info=workbookLogoInfo(v.asset);
  let cx=Math.min(Math.round(v.widthIn*914400),(frame.pageWidthTwips-2*frame.marginTwips)*635),cy=Math.round(cx*info.height/info.width);const maxCy=Math.round(WORKBOOK_LOGO_LIMITS.maxHeightIn*914400);
  if(cy>maxCy){cy=maxCy;cx=Math.round(cy*info.width/info.height);}if(cx<1||cy<1)fail('logo proportions cannot fit this cover.');
  return freeze({schemaVersion:1,id:info.id,sha256:info.sha256,pixelWidth:info.width,pixelHeight:info.height,bytes:info.bytes,role:v.role,alt:v.alt,requestedWidthIn:v.widthIn,widthIn:cx/914400,heightIn:cy/914400,cx,cy});
}
export function workbookLogoRequest(raw){
  if(arguments.length!==1)fail('provide one raw logo draft.');own(raw,['schemaVersion','include','role','alt','widthIn','selection'],'raw logo');
  if(raw.schemaVersion!==1)fail('raw logo version is unsupported.');own(raw.include,['checked','value','invalid'],'logo inclusion');own(raw.include.invalid,WORKBOOK_LOGO_VALIDITY,'logo inclusion validity');
  if(typeof raw.include.checked!=='boolean'||typeof raw.include.value!=='string'||raw.include.value.length>100||WORKBOOK_LOGO_VALIDITY.some(k=>typeof raw.include.invalid[k]!=='boolean'||raw.include.invalid[k]))fail('logo inclusion is invalid; its draft was kept.');
  const read=(field,key)=>{own(field,['value','invalid'],key);own(field.invalid,WORKBOOK_LOGO_VALIDITY,`${key} validity`);if(typeof field.value!=='string'||WORKBOOK_LOGO_VALIDITY.some(k=>typeof field.invalid[k]!=='boolean'||raw.include.checked&&field.invalid[k]))fail(`${key} has unfinished or invalid text; its draft was kept.`);return field.value;};
  const role=read(raw.role,'logo role'),alt=read(raw.alt,'logo description'),width=read(raw.widthIn,'logo width');
  if(width.length>32)fail('logo width text exceeds the supported draft limit.');
  own(raw.selection,['state','generation','asset','current'],'logo selection');const s=raw.selection;
  if(!['empty','ready','loading','invalid'].includes(s.state)||!Number.isSafeInteger(s.generation)||s.generation<0||typeof s.current!=='boolean')fail('logo selection state is unsupported.');
  if(s.asset!==null)asset(s.asset);if(raw.include.checked&&(s.state!=='ready'||!s.current||!s.asset))fail('the current PNG selection is not ready; restore or select a valid logo before exporting.');
  let widthIn=null;if(raw.include.checked){if(!/^(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$/u.test(width))fail('logo width must be a complete plain decimal.');widthIn=Number(width);}
  return validateWorkbookLogo({schemaVersion:1,include:raw.include.checked,asset:s.state==='ready'&&s.current?s.asset:null,role,alt,widthIn});
}
