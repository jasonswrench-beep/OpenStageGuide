export const MAX_MASTER_BYTES=64*1024*1024;
const failure=message=>Object.assign(new Error(message),{code:'OSG_CHECKPOINT'});
function call(action){return new Promise((resolve,reject)=>{try{action(r=>r.status==='succeeded'?resolve(r.value):reject(failure(r.error?.message||'Word could not read the complete document.')));}catch(e){reject(failure(e.message));}});}
export async function readCompressedDocument(document){
  let file;
  try{
    file=await call(cb=>document.getFileAsync('compressed',{sliceSize:65536},cb));
    if(!Number.isSafeInteger(file.size)||file.size<=0||file.size>MAX_MASTER_BYTES||!Number.isSafeInteger(file.sliceCount)||file.sliceCount<1||file.sliceCount>1024)throw failure('This backup supports Word masters up to 64 MB.');
    const bytes=new Uint8Array(file.size);let offset=0;
    for(let i=0;i<file.sliceCount;i++){
      const slice=await call(cb=>file.getSliceAsync(i,cb));
      if(slice.index!==i||!slice.data||slice.data.length===0||slice.data.length>65536||offset+slice.data.length>bytes.length||Array.from(slice.data).some(n=>!Number.isInteger(n)||n<0||n>255))throw failure('Word returned an incomplete document slice. No backup was prepared.');
      bytes.set(slice.data,offset);offset+=slice.data.length;
    }
    if(offset!==bytes.length)throw failure('The Word master was not read completely. No backup was prepared.');
    return bytes;
  }finally{if(file)await call(cb=>file.closeAsync(cb));}
}
