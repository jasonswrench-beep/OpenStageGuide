# Opening educational research extension v1

This optional generated-copy section contains only authored educational rationale, evidence, application, limitations and references. It is broader than psychometrics. No findings, citations, norms, validation status or applicability are invented or verified. The pane and source preview state that authored claims/references need review; generated pages contain only the authored wording and fixed semantic headings.

## Persistence boundary

New research controls are session-only raw drafts, off by default. They do not enter the course, private legacy front matter, publication-v1, persisted brand preset-v1 or scoring schemas. Family/scope/brand/publication/contents changes preserve raw text. Reload or another Word-document pane may lose it; the best-effort unload warning may not be shown by Word. Existing live preview/Apply still uses its inherited refresh/persist infrastructure, but this extension adds no research field to that payload.

## Data and rendering contract

`validateWorkbookResearch` accepts exactly one own-data record `{schemaVersion:1, include:false, rationale:'', evidence:'', application:'', limitations:'', references:''}` and returns a new deeply frozen canonical copy. Each text field is at most16,000 UTF-16 units; the aggregate limit is64,000. All types, own enumerable descriptors, scalar/XML text and budgets are checked even when off. Unknown/accessor/hidden/symbol/inherited/version/null data refuses without reading getters. Actual CRLF/CR normalize to LF in the output; literal backslash text remains literal. Included research requires at least one nonblank authored field. Optional whitespace-only fields are retained but not printed.

`workbookResearchRequest` accepts the corresponding raw record: inclusion is `{checked,value,invalid}` and each field `{value,invalid}`, using the ten native-validity booleans. Active fields must be native-valid; inactive pattern/validity states remain in the raw draft without being published. Shape/type/XML/budgets always apply. No private course/Background fields populate this record automatically.

`workbookResearchRows` constructs a fixed Heading1 `Educational research and rationale`, followed by Heading2 plus ordinary breakable text for each nonblank authored field: Educational rationale, Research evidence, Application to this course, Limitations, References. It creates no fetched link, citation format or external relationship. The section sits after the contents block and before the existing scope/identity/learner body. Its actual emitted headings join the same TOC traversal with stable role/project/section keys. The heading ceiling rises from4001 to4007 for the six possible added headings; existing label/identity budgets remain bounded.

Optional `research` configuration extends family preparation/rendering. Omission or include=false preserves accepted662's exact seven-part DOCX files and public output preview across all five families, selected/full scope and publication/contents combinations. Enabled research adds only ordinary document content and its enabled preview field; no styles, settings, core identity, package parts, assets or relationships are introduced. Existing family styles/contrast/language and learner response geometry remain unchanged.

## Draft and review protection

A separate reset asks Keep/Replace before clearing only research text/inclusion. Confirmation binds raw values/validity, current source/family/scope/options and displayed confirmation; stale/tampered/hidden/busy actions refuse locally. Publication and research reset dialogs cannot be opened over one another. Keep preserves an otherwise-current review; replacement invalidates it. Research raw controls and structure join exact source/review/display/consent guards before and after asynchronous packaging, including the accepted hidden-workspace guard. Other drafts and publication/branding/contents choices are not reset.

## Qualification limits

Source/mock-panel/XML/package tests are not Word page rendering, keyboard/screen-reader, Navigation Pane, Accessibility Checker, print/PDF, research validity or citation verification. Native desktop/server input is not part of this slice. Logo/image branding assets remain a separate next increment; this text section does not satisfy that requirement.
