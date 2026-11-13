// Synthetic captured-package fixture only; never opens Word.
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import {CUE_ASSETS} from '../../src/cue-assets.mjs';
import {createCueOwnership,cueOwnerTag} from '../../src/cue-ownership.mjs';
import {masterFixture} from './checkpoint-fixture.mjs';
import {xml} from '../../src/worksheet.mjs';
export async function ownedFixture(cell=false){
  const {project}=await masterFixture(),item=project.modules[0].activities[0];project.cueOwnership={version:1};item.cueImage=createCueOwnership(item.id,item.icon);item.privateUnknown={preserve:'UNKNOWN ACTIVITY DATA'};
  const {master}=await masterFixture(project),zip=await JSZip.loadAsync(master),png=new Uint8Array(await fs.readFile(new URL('../../public/assets/icons/'+CUE_ASSETS[item.icon].file,import.meta.url)));
  const original=await zip.file('word/document.xml').async('string'),anchor=`<w:sdt><w:sdtPr><w:tag w:val="osg:activity:${item.id}"`,at=original.indexOf(anchor),end=original.indexOf('</w:sdt>',at)+8,title=original.slice(at,end);
  const picture=`<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="Instructional cue" title="${xml('OpenStage Guide '+item.icon+' cue')}" descr="${xml(item.icon+' icon')}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="ownedImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  const owned=`<w:p><w:sdt><w:sdtPr><w:tag w:val="${cueOwnerTag(item.cueImage)}"/></w:sdtPr><w:sdtContent><w:r>${picture}</w:r></w:sdtContent></w:sdt></w:p>`,cue=`<w:p><w:r><w:t>${xml(item.icon+' · '+item.minutes+' min')}</w:t></w:r></w:p>`,block=owned+cue+title;
  zip.file('word/document.xml',original.replace(title,cell?`<w:tbl><w:tr><w:tc>${block}</w:tc></w:tr></w:tbl>`:block));
  zip.file('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="ownedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/owned.png"/></Relationships>');zip.file('word/media/owned.png',png);
  zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  return {project,item,zip,picture,owned,cue,title};
}