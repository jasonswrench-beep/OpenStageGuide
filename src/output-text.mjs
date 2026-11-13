// XML1.0 public-output projection only. Does not validate/rewrite master data.
// Retains existing safe C0 stripping; invalid scalar/noncharacter text is refused.
export function publicOutputText(value){
  if(typeof value!=='string')throw new Error('Unsupported public output text.');
  let result='';
  for(const symbol of value){
    const cp=symbol.codePointAt(0);
    if(cp<0x20&&cp!==9&&cp!==10&&cp!==13)continue;
    if(!(cp===9||cp===10||cp===13||cp>=0x20&&cp<=0xd7ff||cp>=0xe000&&cp<=0xfffd||cp>=0x10000&&cp<=0x10ffff))throw new Error('Some saved output text contains an unsupported character. Preserve your work and correct the public title, instructions, supply name or unit before exporting.');
    result+=symbol;
  }
  return result.replace(/\r\n?/g,'\n');
}
