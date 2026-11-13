// This profile deliberately retains the legacy six payload parts and the four
// directory entries JSZip has always generated for them. It does not widen the
// newer family/assessment or logo package profiles.
export const LEGACY_DOCX_PARTS=Object.freeze(['[Content_Types].xml','_rels/.rels','word/_rels/document.xml.rels','word/document.xml','word/styles.xml','docProps/core.xml']);
const directories=Object.freeze(['_rels/','word/','word/_rels/','docProps/']);
const roster=Object.freeze([...LEGACY_DOCX_PARTS,...directories].sort());
const mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const fail=message=>{throw new Error(`Generated DOCX: ${message}`);};
function ownParts(files){
  if(!files||typeof files!=='object'||Array.isArray(files)||![Object.prototype,null].includes(Object.getPrototypeOf(files)))fail('parts must be a plain data record.');
  const keys=Reflect.ownKeys(files);
  if(keys.length!==LEGACY_DOCX_PARTS.length||keys.some(key=>typeof key!=='string'||!LEGACY_DOCX_PARTS.includes(key)))fail('the six-part profile changed.');
  for(const name of LEGACY_DOCX_PARTS){const d=Object.getOwnPropertyDescriptor(files,name);
    if(!d?.enumerable||!Object.hasOwn(d,'value')||typeof d.value!=='string')fail('parts must be visible own text data.');
  }
}
function envelope(data,expected){
  if(data.length<22)fail('ZIP output is incomplete.');
  const v=new DataView(data.buffer,data.byteOffset,data.byteLength),end=data.length-22;
  if(v.getUint32(end,true)!==0x06054b50||v.getUint16(end+20,true)||v.getUint16(end+10,true)!==roster.length)fail('ZIP roster, comment or trailing data changed.');
  const offset=v.getUint32(end+16,true),size=v.getUint32(end+12,true);
  if(v.getUint16(end+4,true)||v.getUint16(end+6,true)||v.getUint16(end+8,true)!==roster.length||offset+size!==end)fail('ZIP directory bounds changed.');
  let at=offset;const found=[],ranges=[];
  for(let i=0;i<roster.length;i++){
    if(at+46>end||v.getUint32(at,true)!==0x02014b50)fail('ZIP directory entry changed.');
    const n=v.getUint16(at+28,true),extra=v.getUint16(at+30,true),comment=v.getUint16(at+32,true),local=v.getUint32(at+42,true),flags=v.getUint16(at+8,true),method=v.getUint16(at+10,true),compressed=v.getUint32(at+20,true),expanded=v.getUint32(at+24,true);
    if(at+46+n+extra+comment>end)fail('ZIP directory entry bounds changed.');
    const name=new TextDecoder('utf-8',{fatal:true}).decode(data.subarray(at+46,at+46+n));found.push(name);
    const directory=directories.includes(name);
    // Match expanded lengths to already captured canonical inputs instead of
    // imposing an unrelated master-backup size limit on legacy output.
    if(!roster.includes(name)||expanded!==(directory?0:expected[name].length)||method!==(directory?0:8)||(flags&~0x0800)!==0||v.getUint16(at+34,true)||directory&&(compressed||v.getUint32(at+16,true)))fail('ZIP entry profile or size changed.');
    if(local+30>offset||v.getUint32(local,true)!==0x04034b50||v.getUint16(local+6,true)!==flags||v.getUint16(local+8,true)!==method)fail('ZIP local header changed.');
    const ln=v.getUint16(local+26,true),le=v.getUint16(local+28,true),finish=local+30+ln+le+compressed;
    if(finish>offset||ln!==n||!data.subarray(local+30,local+30+ln).every((byte,j)=>byte===data[at+46+j])||v.getUint32(local+14,true)!==v.getUint32(at+16,true)||v.getUint32(local+18,true)!==compressed||v.getUint32(local+22,true)!==expanded)fail('ZIP local bounds, name or CRC changed.');
    if(extra||comment||le||v.getUint16(at+36,true)||v.getUint32(at+38,true)!==(directory?16:0)||v.getUint16(at+4,true)!==0x14||v.getUint16(at+6,true)!==10||v.getUint16(local+4,true)!==10)fail('ZIP entry metadata changed.');
    ranges.push([local,finish]);
    at+=46+n+extra+comment;
  }
  if(at!==end||found.sort().some((name,i)=>name!==roster[i]))fail('ZIP directory roster changed.');
  ranges.sort((a,b)=>a[0]-b[0]);let boundary=0;for(const [start,finish]of ranges){if(start!==boundary)fail('ZIP payloads overlap or contain unlisted data.');boundary=finish;}if(boundary!==offset)fail('ZIP payload boundary changed.');
}

export async function legacyDocxBlob(files,Zip){
  if(typeof Zip!=='function'||typeof Zip.loadAsync!=='function')fail('the ZIP implementation is unavailable.');
  ownParts(files);
  const source={...files},encoder=new TextEncoder(),decoder=new TextDecoder('utf-8',{fatal:true});
  const expected=Object.fromEntries(LEGACY_DOCX_PARTS.map(name=>[name,encoder.encode(source[name])]));
  if(LEGACY_DOCX_PARTS.some(name=>decoder.decode(expected[name])!==source[name]))fail('text contains an unsupported Unicode scalar.');
  // Encoding each complete string avoids splitting a surrogate pair at JSZip's
  // streaming UTF-16 chunk boundary. Never pass the private expectations to it.
  const zip=new Zip();for(const name of LEGACY_DOCX_PARTS)zip.file(name,new Uint8Array(expected[name]),{binary:true});
  const generated=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  if(!(generated instanceof Uint8Array))fail('the ZIP implementation returned unsupported bytes.');
  const offered=Uint8Array.from(generated);envelope(offered,expected);
  const loaded=await Zip.loadAsync(Uint8Array.from(offered),{checkCRC32:true,createFolders:false}),names=Object.keys(loaded.files).sort();
  if(names.length!==roster.length||names.some((name,i)=>name!==roster[i])||directories.some(name=>loaded.files[name]?.dir!==true))fail('decoded ZIP roster changed.');
  for(const name of LEGACY_DOCX_PARTS){
    if(loaded.files[name]?.dir!==false)fail(`payload ${name} changed type.`);
    const actual=await loaded.file(name)?.async('uint8array'),canonical=expected[name];
    if(!(actual instanceof Uint8Array)||actual.length!==canonical.length||actual.some((byte,i)=>byte!==canonical[i]))fail(`payload ${name} changed.`);
  }
  envelope(offered,expected);ownParts(files);
  if(LEGACY_DOCX_PARTS.some(name=>files[name]!==source[name]))fail('source strings changed during packaging.');
  return new Blob([Uint8Array.from(offered)],{type:mime});
}
