// The fonts that come with the program: a small pack of free fonts made to be easy to read, kept in program/fonts with their licenses. The page
// draws words in them whatever the computer has installed, and a Word file made from a guide that uses one carries it inside (src/guide-docx.mjs),
// so the file looks right on a computer that does not have the font.
//
// The fonts are the makers' own files, unmodified, each under the SIL Open Font License 1.1, which allows them to be used, shipped and
// embedded in documents. Their credits are in NOTICE. Every one is a static TrueType font whose own settings allow embedding, which is what
// Word can carry inside a file.
export const FONT_DIR='program/fonts';
export const FONT_URL='/program/fonts/';
export const STYLES=Object.freeze(['regular','bold','italic','boldItalic']);

/** @typedef {{family:string, kind:'sans'|'serif', note:string, files:Partial<Record<'regular'|'bold'|'italic'|'boldItalic',string>>, license:string, credit:string, source:string}} PackFont */
/** @type {readonly PackFont[]} */
export const FONT_PACK=Object.freeze([
  {family:'Atkinson Hyperlegible',kind:'sans',note:'Made by the Braille Institute so that every letter is easy to tell apart.',
    files:{regular:'AtkinsonHyperlegible-Regular.ttf',bold:'AtkinsonHyperlegible-Bold.ttf',italic:'AtkinsonHyperlegible-Italic.ttf',boldItalic:'AtkinsonHyperlegible-BoldItalic.ttf'},
    license:'OFL-AtkinsonHyperlegible.txt',credit:'Copyright 2020 Braille Institute of America, Inc.',source:'https://github.com/googlefonts/atkinson-hyperlegible'},
  {family:'Lexend',kind:'sans',note:'Made to make reading easier and faster, with wide, even letter spacing.',
    files:{regular:'Lexend-Regular.ttf',bold:'Lexend-Bold.ttf'},
    license:'OFL-Lexend.txt',credit:'Copyright 2019 The Lexend Project Authors',source:'https://github.com/googlefonts/lexend'},
  {family:'Source Sans 3',kind:'sans',note:'A clean, open sans serif for long reading on a screen or in print.',
    files:{regular:'SourceSans3-Regular.ttf',bold:'SourceSans3-Bold.ttf',italic:'SourceSans3-It.ttf',boldItalic:'SourceSans3-BoldIt.ttf'},
    license:'LICENSE-SourceSans3.md',credit:'Copyright 2010-2024 Adobe, with Reserved Font Name Source',source:'https://github.com/adobe-fonts/source-sans'},
  {family:'Source Serif 4',kind:'serif',note:'A readable serif for printed handouts and longer text.',
    files:{regular:'SourceSerif4-Regular.ttf',bold:'SourceSerif4-Bold.ttf',italic:'SourceSerif4-It.ttf',boldItalic:'SourceSerif4-BoldIt.ttf'},
    license:'LICENSE-SourceSerif4.md',credit:'Copyright 2014-2024 Adobe, with Reserved Font Name Source',source:'https://github.com/adobe-fonts/source-serif'},
]);

/** The pack's entry for a font name, or null. */
export const packFont=name=>FONT_PACK.find(f=>f.family===name)||null;

/** The font names a piece of Word XML asks for (w:ascii, w:hAnsi, w:cs, w:eastAsia), once each. @param {...string} texts */
export function fontNamesIn(...texts){
  const out=new Set();
  for(const t of texts)for(const m of String(t||'').matchAll(/w:(?:ascii|hAnsi|cs|eastAsia)="([^"]{1,64})"/g))out.add(m[1]);
  return [...out];
}

/** The pack's fonts that some Word XML uses, in the pack's order. @param {...string} texts @returns {PackFont[]} */
export function packFontsUsed(...texts){
  const names=new Set(fontNamesIn(...texts));
  return FONT_PACK.filter(f=>names.has(f.family));
}

/** Fonts for a font box to offer: the pack first (with what each is for), then the fonts Word documents usually have. */
export const COMMON_FONTS=Object.freeze(['Aptos','Calibri','Cambria','Arial','Verdana','Tahoma','Segoe UI','Georgia','Times New Roman','Trebuchet MS']);
export function fontChoices(){
  return [...FONT_PACK.map(f=>({value:f.family,label:`${f.family}: ${f.note} (comes with the program)`})),...COMMON_FONTS.map(f=>({value:f,label:f}))];
}

/**
 * Obfuscate (or, the same way round, read back) a font file as Word keeps embedded fonts: the first 32 bytes are mixed with the bytes of a GUID
 * written in reverse. The same call undoes it.
 * @param {Uint8Array} bytes @param {string} guid like {D9AC0A92-6A2B-4C0E-8F5B-1E6A3C9D7B21}
 */
export function obfuscateFont(bytes,guid){
  const hex=String(guid).replace(/[{}-]/g,'');
  if(!/^[0-9A-Fa-f]{32}$/.test(hex))throw new Error('A font key is a GUID.');
  const key=Array.from({length:16},(_,i)=>parseInt(hex.slice(i*2,i*2+2),16)).reverse();
  const out=new Uint8Array(bytes);
  for(let i=0;i<Math.min(32,out.length);i++)out[i]^=key[i%16];
  return out;
}

/** A new GUID in Word's form, in capitals. */
export function newGuid(){
  const b=new Uint8Array(16);
  globalThis.crypto.getRandomValues(b);
  b[6]=(b[6]&0x0f)|0x40;b[8]=(b[8]&0x3f)|0x80;
  const h=[...b].map(x=>x.toString(16).padStart(2,'0')).join('').toUpperCase();
  return `{${h.slice(0,8)}-${h.slice(8,12)}-${h.slice(12,16)}-${h.slice(16,20)}-${h.slice(20)}}`;
}
