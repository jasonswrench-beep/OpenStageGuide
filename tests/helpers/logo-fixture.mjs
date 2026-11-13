import {deflateSync} from 'node:zlib';
const table=Array.from({length:256},(_,n)=>{let c=n;for(let j=0;j<8;j++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c>>>0;});
export const logoCrc=b=>{let c=0xffffffff;for(const v of b)c=table[(c^v)&255]^(c>>>8);return (c^0xffffffff)>>>0;};
export function logoChunk(type,data=new Uint8Array()) {const b=Buffer.alloc(data.length+12);b.writeUInt32BE(data.length);b.write(type,4,'ascii');b.set(data,8);b.writeUInt32BE(logoCrc(b.subarray(4,b.length-4)),b.length-4);return b;}
export function syntheticLogo({width=4,height=2,colorType=6,depth=8,interlace=0,before=[],after=[],scanlines=null,idat=null}={}){
 const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(width);ihdr.writeUInt32BE(height,4);ihdr[8]=depth;ihdr[9]=colorType;ihdr[12]=interlace;
 const channels=colorType===2?3:4,raw=scanlines||Buffer.alloc(height*(1+width*channels));
 if(!scanlines)for(let y=0;y<height;y++)for(let x=0;x<width;x++){const at=y*(1+width*channels)+1+x*channels;raw[at]=x*20;raw[at+1]=y*40;raw[at+2]=120;if(channels===4)raw[at+3]=255;}
 return new Uint8Array(Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),logoChunk('IHDR',ihdr),...before,logoChunk('IDAT',idat||deflateSync(raw)),...after,logoChunk('IEND')]));
}
