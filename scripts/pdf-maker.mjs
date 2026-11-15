// Makes a PDF of the guide, on this computer, for File > Save as PDF. The guide's web page (src/guide-html.mjs) is printed to a PDF by Microsoft Edge, which
// comes with Windows and is the same engine the program's window shows its pages with. Edge writes a tagged PDF when it prints from the command line: the
// headings, paragraphs and tables are marked, and it says its language and title. That is looked for in the file that comes back, not assumed.
//
// The page is written to a temporary folder with a name nobody can guess, Edge is run on it with a fixed list of options (no shell, no address but that file's)
// and a time limit, and the folder is removed again, whatever happens. The page has no script (and says so to the browser: Content-Security-Policy), so
// nothing in a guide can run anything. Nothing is sent anywhere: Edge is not given an address on the network.
import {spawn} from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import {existsSync} from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {pathToFileURL} from 'node:url';

export const PDF_MAX_HTML=40*1024*1024;
export const PDF_TIMEOUT_MS=90000;

/** Where Edge is on this computer (Windows keeps it with the system's program files), or null. */
export function findEdge({env=process.env,exists=existsSync}={}){
  for(const base of [env['ProgramFiles(x86)'],env.ProgramFiles,env.ProgramW6432]){
    if(!base)continue;
    const file=path.join(base,'Microsoft','Edge','Application','msedge.exe');
    if(exists(file))return file;
  }
  return null;
}

/** Whether a PDF is marked for screen readers: it holds a structure tree and says it is marked. (The PDF Edge writes keeps these names in the clear.) @param {Buffer} bytes */
export function pdfIsTagged(bytes){
  const text=Buffer.from(bytes).toString('latin1');
  return text.includes('/StructTreeRoot')&&/\/Marked\s+true/.test(text);
}

export class PdfError extends Error{
  constructor(message,code='failed'){super(message);this.name='PdfError';this.code=code;}
}

/**
 * The PDF of a page, as bytes, and whether it is tagged. `run` and `edge` are replaced in tests.
 * @param {string} html @param {{ edge?: string | null, run?: typeof spawn, tmp?: string, timeoutMs?: number }} [options]
 * @returns {Promise<{ bytes: Buffer, tagged: boolean }>}
 */
export async function makePdf(html,{edge=findEdge(),run=spawn,tmp=os.tmpdir(),timeoutMs=PDF_TIMEOUT_MS}={}){
  if(typeof html!=='string'||!html.trim())throw new PdfError('There is nothing to make a PDF of.','empty');
  if(html.length>PDF_MAX_HTML)throw new PdfError('This guide is too large to make into a PDF here.','too-large');
  if(/<script[\s>]/i.test(html))throw new PdfError('That page holds a script, so it was not printed.','refused');
  if(!edge)throw new PdfError('The program makes a PDF with Microsoft Edge, which is not on this computer. Save the Word file and make the PDF in Word instead.','no-edge');
  const dir=await fs.mkdtemp(path.join(tmp,`osg-pdf-${crypto.randomBytes(4).toString('hex')}-`));
  try{
    const page=path.join(dir,'page.html'),out=path.join(dir,'out.pdf');
    await fs.writeFile(page,html,{encoding:'utf8',flag:'wx'});
    const args=['--headless=new','--disable-gpu','--no-pdf-header-footer','--no-first-run','--no-default-browser-check','--disable-extensions','--disable-sync','--disable-background-networking',
      `--user-data-dir=${path.join(dir,'profile')}`,`--print-to-pdf=${out}`,pathToFileURL(page).href];
    await new Promise((resolve,reject)=>{
      let child,done=false;
      const finish=err=>{if(done)return;done=true;clearTimeout(timer);err?reject(err):resolve();};
      const timer=setTimeout(()=>{try{child?.kill();}catch{/* already gone */}finish(new PdfError('Edge did not finish making the PDF in time.','timeout'));},timeoutMs);
      try{child=run(edge,args,{windowsHide:true,stdio:'ignore',shell:false});}catch(e){finish(new PdfError(`Edge could not be started: ${e.message}`,'start'));return;}
      child.once('error',e=>finish(new PdfError(`Edge could not be started: ${e.message}`,'start')));
      child.once('exit',()=>finish());
    });
    let bytes;
    try{bytes=await fs.readFile(out);}catch{throw new PdfError('Edge did not make a PDF.','no-output');}
    if(bytes.length<64||bytes.subarray(0,5).toString('latin1')!=='%PDF-')throw new PdfError('Edge made something that is not a PDF.','bad-output');
    return {bytes,tagged:pdfIsTagged(bytes)};
  }finally{
    await fs.rm(dir,{recursive:true,force:true}).catch(()=>{});
  }
}
