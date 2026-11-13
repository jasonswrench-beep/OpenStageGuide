// Getting a picture ready to go into the guide: what kind of picture it is, how big, and (in a browser) shrinking one that is
// larger than a page can use, so the document stays small. The reading of a file's kind and size is plain code and is tested
// without a browser; the shrinking needs a canvas and is only done where one exists.
export const MAX_INPUT_BYTES=12*1024*1024;
export const KEEP_BYTES=1500000;     // a picture this small and no larger than MAX_SIDE is used as it is
export const MAX_SIDE=1600;          // pixels on the longer side after shrinking
export const MAX_UNSHRUNK_BYTES=5*1024*1024; // where nothing can shrink a picture, anything over this is refused
export const PICTURE_TYPES=Object.freeze({png:'image/png',jpeg:'image/jpeg',gif:'image/gif'});
export const PICTURE_HELP='Choose a PNG, JPG or GIF picture.';

const u16=(b,i)=>(b[i]<<8)|b[i+1],u32=(b,i)=>((b[i]<<24)|(b[i+1]<<16)|(b[i+2]<<8)|b[i+3])>>>0;
// {type,mime,width,height} from the first bytes of a PNG, JPEG or GIF file; anything else is refused in plain words.
export function imageInfo(bytes){
  const b=bytes instanceof Uint8Array?bytes:new Uint8Array(bytes);
  if(b.length>=24&&b[0]===0x89&&b[1]===0x50&&b[2]===0x4e&&b[3]===0x47&&u32(b,12)===0x49484452){
    const width=u32(b,16),height=u32(b,20);
    if(width&&height)return {type:'png',mime:PICTURE_TYPES.png,width,height};
  }
  if(b.length>=10&&b[0]===0x47&&b[1]===0x49&&b[2]===0x46&&b[3]===0x38){
    const width=b[6]|(b[7]<<8),height=b[8]|(b[9]<<8);
    if(width&&height)return {type:'gif',mime:PICTURE_TYPES.gif,width,height};
  }
  if(b.length>=4&&b[0]===0xff&&b[1]===0xd8){
    let i=2;
    while(i+9<b.length){
      if(b[i]!==0xff){i++;continue;}
      const marker=b[i+1];
      if(marker===0xff){i++;continue;}
      if(marker===0xd8||marker===0x01||(marker>=0xd0&&marker<=0xd7)){i+=2;continue;}
      const len=u16(b,i+2);
      // The start-of-frame markers hold the size; C4, C8 and CC are other things in the same range.
      if(marker>=0xc0&&marker<=0xcf&&![0xc4,0xc8,0xcc].includes(marker)){
        const height=u16(b,i+5),width=u16(b,i+7);
        if(width&&height)return {type:'jpeg',mime:PICTURE_TYPES.jpeg,width,height};
        break;
      }
      i+=2+len;
    }
  }
  throw new Error(`That file is not a picture the guide can use. ${PICTURE_HELP}`);
}

export function toBase64(bytes){
  if(typeof Buffer!=='undefined')return Buffer.from(bytes).toString('base64');
  let s='';
  for(let i=0;i<bytes.length;i+=0x8000)s+=String.fromCharCode.apply(null,bytes.subarray(i,i+0x8000));
  return btoa(s);
}

// The picture's size on the page, in EMU (914400 to the inch), for the width the author chose, never taller than 8 inches.
export const PICTURE_WIDTHS=Object.freeze([
  {id:'large',label:'Fits the block (4.8 inches)',inches:4.8},
  {id:'medium',label:'Medium (3.5 inches)',inches:3.5},
  {id:'small',label:'Small (2.2 inches)',inches:2.2}
]);
export function pictureSize(info,widthId='large'){
  const w=(PICTURE_WIDTHS.find(x=>x.id===widthId)||PICTURE_WIDTHS[0]).inches;
  let cx=Math.round(w*914400),cy=Math.round(cx*info.height/info.width);
  const max=8*914400;
  if(cy>max){cx=Math.round(cx*max/cy);cy=max;}
  return {cx,cy};
}

// file: a browser File (name, size, arrayBuffer()). Returns {bytes,type,mime,width,height,name,shrunk}.
// env: where a canvas can be found; the default is the browser. Without one a picture is used as it is, or refused if it is huge.
export async function prepareImage(file,{env=globalThis}={}){
  if(!file)throw new Error(`Choose a picture file first. ${PICTURE_HELP}`);
  if(file.size>MAX_INPUT_BYTES)throw new Error(`That picture is ${(file.size/1048576).toFixed(1)} MB. Choose one under ${MAX_INPUT_BYTES/1048576} MB.`);
  const bytes=new Uint8Array(await file.arrayBuffer());
  const info=imageInfo(bytes),name=String(file.name||'picture');
  const big=Math.max(info.width,info.height)>MAX_SIDE||bytes.length>KEEP_BYTES;
  if(!big||info.type==='gif')return {bytes,...info,name,shrunk:false};
  const canvasOk=typeof env.createImageBitmap==='function'&&env.document&&typeof env.document.createElement==='function';
  if(!canvasOk){
    if(bytes.length>MAX_UNSHRUNK_BYTES)throw new Error(`That picture is too large to use as it is (${(bytes.length/1048576).toFixed(1)} MB). Shrink it first, then choose it again.`);
    return {bytes,...info,name,shrunk:false};
  }
  const bitmap=await env.createImageBitmap(new env.Blob([bytes],{type:info.mime}));
  const scale=Math.min(1,MAX_SIDE/Math.max(bitmap.width,bitmap.height)),width=Math.max(1,Math.round(bitmap.width*scale)),height=Math.max(1,Math.round(bitmap.height*scale));
  const canvas=env.document.createElement('canvas');canvas.width=width;canvas.height=height;
  canvas.getContext('2d').drawImage(bitmap,0,0,width,height);
  bitmap.close?.();
  // A PNG stays a PNG (it may have see-through parts); a JPEG stays a JPEG.
  const blob=await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('The picture could not be shrunk.')),info.mime,0.85));
  const out=new Uint8Array(await blob.arrayBuffer());
  // If shrinking made it no smaller, the original is the better file.
  if(out.length>=bytes.length&&scale===1)return {bytes,...info,name,shrunk:false};
  return {bytes:out,type:info.type,mime:info.mime,width,height,name,shrunk:true};
}
