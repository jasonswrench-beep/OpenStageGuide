import {FAMILY_BRAND_FIELDS} from '../../src/workbook-family-flow.mjs';
import {WORKBOOK_COVER_FIELDS,WORKBOOK_COPYRIGHT_FIELDS} from '../../src/workbook-public-frontmatter.mjs';
import {WORKBOOK_RESEARCH_FIELDS} from '../../src/workbook-research.mjs';
export function workbookDesignDraft(familyId='clearline'){
 const brand=Object.fromEntries(FAMILY_BRAND_FIELDS.map(k=>[k,'']));brand.organizationName='Synthetic organization \u4e2d\u6587 \u{1f331}';
 const cover=Object.fromEntries(WORKBOOK_COVER_FIELDS.map(k=>[k,'']));Object.assign(cover,{title:'Synthetic public cover \u{1d11e}',subtitle:'Authored subtitle\nSecond line',author:'Synthetic author',edition:'First',year:'2026',audience:'Workshop participants'});
 const copyright=Object.fromEntries(WORKBOOK_COPYRIGHT_FIELDS.map(k=>[k,'']));Object.assign(copyright,{rightsHolder:'Synthetic rights holder',year:'2026',noticeChoice:'custom-text',noticeText:'Authored notice & <literal text>.\nNo rights are inferred.',permissionsContact:'Synthetic contact',credits:'Synthetic credits'});
 const research={include:true,...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,`Authored ${k}\nEvidence not verified.`]))};
 return {familyId,baseLanguage:'en-US',brand,frontMatter:{includeCover:true,includeCopyright:true,cover,copyright},contents:{include:true},research,logo:{include:false,role:'informative',alt:'Synthetic geometric organization mark',widthIn:'1.5'}};
}
