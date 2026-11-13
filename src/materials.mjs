const clean=v=>v.normalize('NFKC').trim().replace(/\s+/gu,' ');
export const materialKey=m=>JSON.stringify([clean(m.item).toLowerCase(),clean(m.unit),m.use]);
export function normalizeMaterials(list){
  if(!Array.isArray(list)||list.length>100)throw new Error('Use a materials list with up to 100 items per activity.');
  const seen=new Set();return list.map(m=>{
    const names=['item','quantity','unit','use','notes'];
    if(!m||typeof m!=='object'||Array.isArray(m)||Object.keys(m).length!==names.length||names.some(k=>!Object.hasOwn(m,k)))throw new Error('Unsupported material entry. Preserve the course and use the matching tool version.');
    const out={};for(const [key,max] of [['item',200],['unit',80],['notes',2000]]){if(typeof m[key]!=='string'||m[key].length>max)throw new Error('Enter supported material text.');out[key]=m[key].replace(/\r\n?/g,'\n').trim();}
    if(!out.item||!out.unit)throw new Error('Enter a material name and unit.');
    if((typeof m.quantity!=='number'&&typeof m.quantity!=='string')||String(m.quantity).trim()==='')throw new Error('Enter a positive material quantity.');
    const n=Number(m.quantity),scaled=Math.round(n*1000);
    if(!Number.isFinite(n)||n<=0||n>100000||Math.abs(n*1000-scaled)>1e-7||scaled===0)throw new Error('Use a quantity above zero, up to 100000, with at most three decimal places.');
    if(!['consumable','reusable'].includes(m.use))throw new Error('Choose consumable or reusable.');out.quantity=scaled/1000;out.use=m.use;
    const key=materialKey(out);if(seen.has(key))throw new Error('Combine duplicate material names, units and uses into one activity entry.');seen.add(key);return out;
  });
}
