# Public workbook front matter extension v1

This generated-copy extension adds optional text-only cover and copyright metadata to the five workbook families, in full-course or selected-plan scope. It is separate from the existing course/frontMatter schema and private facilitator sections. It never reads those private sections or infers organization, author, rights holder, audience, year or license from them.

## Persistence and privacy boundary

All new controls are session-only raw drafts. They are not written to the course, Word master, localStorage or the version1 sixteen-field organization-brand presets. Navigation and changing design, scope or brand preserve the current raw text. Reload/document switching may lose it; a best-effort unload warning is not guaranteed by Word. Organization uses the existing explicit public Organization branding control, not a second cover organization source. Only included public rows enter the output/preview. Inactive cover/copyright and unused license fields stay in the raw draft and never appear in files or public preview.

## Canonical own-data metadata

`validateWorkbookFrontMatter(metadata)` takes exactly one own-data plain record and returns a new deeply frozen canonical copy without changing/freezing caller data:

```js
{
  schemaVersion: 1,
  includeCover: false,
  includeCopyright: false,
  cover: {title:'', subtitle:'', author:'', edition:'', year:'', audience:''},
  copyright: {
    rightsHolder:'', year:'', noticeChoice:'', noticeText:'',
    permissionsContact:'', credits:'',
    licenseName:'', licenseVersion:'', licenseUrl:''
  }
}
```

Every field is explicit bounded XML-safe text. Unknown/hidden/symbol/accessor/inherited data, unsupported version/null/type and malformed scalar text refuse without reading getters. Actual CRLF/CR normalize to LF only in the public projection; literal backslash text is preserved. Raw controls retain spelling/validity. Individual text limits and a 64,000-unit ceiling bound output. Cover title, when included, must be nonblank. Copyright requires a separate rights holder, year/range, notice choice and an authored or explicitly requested nonblank notice. A year is four digits or an ordered range separated by hyphen or en dash; no current year is inferred. Optional fields containing only whitespace are not printed (an active nonempty unsupported year refuses).

Notice choices are initially unselected: `all-rights-reserved`, `custom-permissions`, `named-license`, or `custom-text`. Choice alone never inserts wording or selects ownership. Named-license additionally requires the author's explicit name, version and HTTP/HTTPS terms URL with no embedded credentials; URLs are printed text, not fetched or added as relationships. The app does not establish legal sufficiency or permissions. Any authored public text can itself contain private information and needs review.

`workbookFrontMatterRequest(raw)` validates the raw control record and returns canonical metadata. It has the same root/cover/copyright keys, but each cover/copyright text is `{value,invalid}` using the existing ten native-validity booleans. Each inclusion choice is `{checked,value,invalid}` so checked state and native validity are separately bound. All raw records remain strictly own-data. Active fields must be native-valid. Inactive cover/copyright patterns and unused named-license flags do not block another chosen notice, while their types/XML/text limits are still checked.

`workbookCopyrightNoticeStarter(metadata)` is explicit, definition-only text generation. It does not choose a notice. All-rights-reserved produces only its editable sentence; custom-permissions produces a neutral instruction to specify authorized uses/conditions; named-license uses only authored name/version/URL; custom-text is written directly. Applying a starter and resetting the draft require guarded Keep/Replace controls. Pending replacements bind exact raw values/validity, current course/design/scope/generation and confirmation text/structure/visibility. A stale or tampered confirmation cannot overwrite raw work. Keep preserves any otherwise-current review. Pending choices block preview/Apply; hidden/detached/busy cached draft actions refuse locally.

## Rendering and compatibility

`prepareFamilyWorkbook` accepts an optional canonical `frontMatter` config field. `renderWorkbookFamily` accepts optional `frontMatter` beside `baseLanguage`. Omission and both inclusion choices off preserve the exact existing profile3 seven-part files and preview. This is separately versioned extension1, not a new package roster, asset profile, course migration or legacy generator change.

Enabled cover replaces only the old generated title/subtitle/organization opening. It uses the explicit cover title for the visible title and offered filename; core title contains that title followed by the authored subtitle when nonblank. Explicit author becomes core creator. Original semantic course identity/content is unchanged. With no cover, the old opening and old core identity remain. Enabled copyright uses a real Heading1 and normal-flow paragraphs; it does not create an implicit cover. Order is optional cover, optional copyright, then the unchanged scope/identity/learner body. Cover-to-copyright and front-matter-to-scope use explicit pageBreakBefore. The last cover paragraph overrides keepNext to avoid pulling a sparse cover into later content. Long notice/credit paragraphs remain breakable. No floating shapes, tables, image assets, headers or footers are added.

Existing family styles, applied brand colors/font/spacing and response rules are reused unchanged. All file parts except document/core remain byte-identical; core changes only for explicit enabled cover title/author. Copyright choices never silently add core rights claims. Public review includes every included row and the old opening in copyright-only mode. Raw front-matter controls and nested visible structure join existing review/source/options guards before and after packaging; acknowledgement event epochs detect an uncheck/recheck during async assembly.

## Qualification boundary and next work

Automated source, mock-panel and ZIP/XML tests are not actual Word page or accessibility qualification. These outputs still require native pagination, Navigation Pane, keyboard/screen-reader, Accessibility Checker, print and PDF review. Logo/cover-image assets, semantic TOC/update/page-number behavior and the educational research/rationale/citation section remain separate pending increments. No network, server activation, GitHub push, learner data collection or native Word mutation is part of this source slice.
