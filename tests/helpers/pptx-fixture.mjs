import JSZip from 'jszip';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const rel=(id,type,target,mode='')=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}" ${mode?`TargetMode="${mode}"`:''}/>`;
const relationships=body=>`<Relationships xmlns="${PKG}">${body}</Relationships>`;
const shape=(text,type='body')=>`<p:sp><p:nvSpPr><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
export async function pptxFixture(change=()=>{},compression='DEFLATE'){
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>');
  zip.file('_rels/.rels',relationships(rel('rId1','officeDocument','ppt/presentation.xml')));
  zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="258" r:id="s2"/><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels',relationships(rel('s1','slide','slides/slide1.xml')+rel('s2','slide','slides/slide2.xml')));
  for(let i=1;i<=2;i++)zip.file(`ppt/slides/slide${i}.xml`,`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${shape('Title '+i,'title')}${shape('Text &amp; task '+i)}</p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels',relationships(rel('n1','notesSlide','../notesSlides/notesSlide1.xml')+rel('h1','hyperlink','https://example.invalid/private','External')));
  zip.file('ppt/notesSlides/notesSlide1.xml',`<p:notes xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>${shape('PRIVATE SPEAKER NOTE')}${shape('SLIDE NUMBER OMIT','sldNum')}</p:spTree></p:cSld></p:notes>`);
  await change(zip);
  return zip.generateAsync({type:'uint8array',compression});
}
