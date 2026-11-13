import {captureEnvelope} from './checkpoint.mjs';
import {captureWorkbookLogoOutput,FAMILY_LOGO_PARTS,FAMILY_LOGO_CONTENT_TYPES_XML,FAMILY_LOGO_DOCUMENT_RELS_XML,FAMILY_WORKBOOK_SETTINGS_XML} from './workbook-family-output.mjs';
import {WORKBOOK_LOGO_MEDIA_PART} from './workbook-logo.mjs';

const fail=message=>{throw new Error(`Logo workbook package: ${message}`);};
const typed=Object.getPrototypeOf(Uint8Array.prototype),byteLength=Object.getOwnPropertyDescriptor(typed,'byteLength').get,bufferOf=Object.getOwnPropertyDescriptor(typed,'buffer').get;
const copy=value=>{let size,buf;try{size=byteLength.call(value);buf=bufferOf.call(value);}catch{fail('exporter returned unsupported bytes.');}if(Object.getPrototypeOf(value)!==Uint8Array.prototype||Object.getPrototypeOf(buf)!==ArrayBuffer.prototype||Object.getOwnPropertySymbols(value).length||['constructor','buffer','byteLength','byteOffset','length'].some(k=>Object.hasOwn(value,k))||size>32*1024*1024)fail('exporter returned unsupported bytes.');const out=new Uint8Array(size);Uint8Array.prototype.set.call(out,value);return out;};
const decode=value=>{if(typeof value!=='string'||value.length>3*1024*1024||!value.length||value.length%4!==0||!/^[A-Za-z0-9+/]*={0,2}$/u.test(value))fail('media encoding changed.');let raw;try{raw=atob(value);}catch{fail('media encoding changed.');}if(btoa(raw)!==value)fail('media encoding is not canonical.');const out=new Uint8Array(raw.length);for(let i=0;i<raw.length;i++)out[i]=raw.charCodeAt(i);return out;};
const same=(a,b)=>a.length===b.length&&a.every((v,i)=>v===b[i]);
function roster(bytes){
  if(bytes.length<22)fail('ZIP is incomplete.');const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),end=bytes.length-22;
  if(view.getUint32(end,true)!==0x06054b50||view.getUint16(end+20,true)!==0)fail('ZIP comments or trailing data are unsupported.');
  let names;try{names=captureEnvelope(bytes);}catch(e){fail(`ZIP envelope changed: ${e.message}`);}
  if(names.length!==FAMILY_LOGO_PARTS.length||names.some((name,i)=>name!==[...FAMILY_LOGO_PARTS].sort()[i])||view.getUint16(end+8,true)!==FAMILY_LOGO_PARTS.length||view.getUint16(end+10,true)!==FAMILY_LOGO_PARTS.length)fail('ZIP parts changed or duplicated.');
  let at=view.getUint32(end+16,true);
  for(let i=0;i<FAMILY_LOGO_PARTS.length;i++){
    const n=view.getUint16(at+28,true),extra=view.getUint16(at+30,true),comment=view.getUint16(at+32,true),local=view.getUint32(at+42,true);
    if(extra||comment||view.getUint16(local+28,true)||view.getUint16(at+36,true)||view.getUint32(at+38,true)||view.getUint16(at+4,true)!==0x14)fail('ZIP entry metadata changed.');
    at+=46+n+extra+comment;
  }
}
export async function logoWorkbookBlob(output,Zip){
  if(arguments.length!==2)fail('provide the owned generated output and ZIP implementation only.');
  // The renderer owns the frozen output and private expected-part snapshot. A
  // clone or caller-authored hash/preview/file map never substitutes for it.
  const snapshot=captureWorkbookLogoOutput(output),files=snapshot.files;
  if(typeof Zip!=='function'||typeof Zip.loadAsync!=='function')fail('the DOCX exporter is unavailable.');
  if(files['[Content_Types].xml']!==FAMILY_LOGO_CONTENT_TYPES_XML||files['word/_rels/document.xml.rels']!==FAMILY_LOGO_DOCUMENT_RELS_XML||files['word/settings.xml']!==FAMILY_WORKBOOK_SETTINGS_XML||Object.keys(files).length!==8||FAMILY_LOGO_PARTS.some(k=>typeof files[k]!=='string'))fail('image profile changed.');
  const encoder=new TextEncoder(),expected=Object.fromEntries(FAMILY_LOGO_PARTS.map(k=>[k,k===WORKBOOK_LOGO_MEDIA_PART?decode(files[k]):encoder.encode(files[k])]));
  if(!same(expected[WORKBOOK_LOGO_MEDIA_PART],snapshot.media))fail('verified PNG binding changed.');
  const zip=new Zip();for(const name of FAMILY_LOGO_PARTS)zip.file(name,copy(expected[name]),{createFolders:false,binary:true});
  const bytes=copy(await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}));roster(bytes);
  const loaded=await Zip.loadAsync(copy(bytes),{checkCRC32:true,createFolders:false}),names=Object.keys(loaded.files);
  if(names.length!==8||FAMILY_LOGO_PARTS.some(k=>!names.includes(k)))fail('readback inventory changed.');
  const decoder=new TextDecoder('utf-8',{fatal:true});
  for(const name of FAMILY_LOGO_PARTS){const actual=copy(await loaded.file(name)?.async('uint8array'));if(!same(actual,expected[name]))fail(`readback ${name} changed.`);if(name!==WORKBOOK_LOGO_MEDIA_PART){let text;try{text=decoder.decode(actual);}catch{fail('XML UTF-8 changed.');}if(!same(encoder.encode(text),expected[name]))fail('XML canonical bytes changed.');}}
  roster(bytes);return new Blob([copy(bytes)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
