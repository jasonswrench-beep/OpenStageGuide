# The OpenStage Guide window

A Tauri shell around the program. Everything the program does is done by its local server, which is Node running the program's own code
(`../scripts/project-server.mjs`). This folder is only the window: it starts that server, waits for it to say where it is listening, and
points the window at it. It has no commands for the page to call, and the page it shows is given no access to the machine beyond what the
server allows it. When the window closes the server stops; if the window is killed, the server notices its input has closed and stops.

It is **not part of the program's own dependencies**: installing the repository never fetches the Tauri tools, and the program keeps its two
runtime packages (acorn and jszip). The window is the one place that has more, and it is not needed to use the program in a browser
(`Start OpenStage Guide.cmd` at the top of the repository).

It follows the OpenAssess window (same owner, Apache-2.0), without its Save dialog: a file the program makes (the participant workbook, a
backup) goes to the Downloads folder, and the page says so. Guides are kept in `OpenStage Guides` inside the person's own Documents folder,
wherever Windows has put it.

**Closing the window.** The page saves a moment after each change, and a page that is going away cannot be sure its last request is sent, so the
first request to close is held: the window asks the page (`window.osgClosing`, run by the window itself; there is still no command for the page to
call), the page saves the open guide and tells its server (`/api/closing`), and the server says one line on its output that the window reads
(`OPENSTAGE_CLOSE_READY`, or `OPENSTAGE_CLOSE_HELD` when the guide could not be saved, which keeps the window open with the page's message; closing
it again then leaves anyway). A page that does not answer within five seconds is not waited for. Then the server's input is closed, which it takes as
the end, and it is stopped by force only if it has not gone in two seconds. `tests/program-closing.test.mjs` and `npm run check:closing` try the
server and the page; the window itself was tried on 2026-11-14 with a test build (below).

**Trying the window without touching the guides.** `OPENSTAGE_GUIDES_FOLDER` (a full path) names another guides folder, for trying a build; nothing
else can change where the guides are. Build with `TAURI_CONFIG='{"identifier":"org.openstage.guide.closetest"}'` and `--features tauri/custom-protocol`
to give the test build its own WebView2 profile (the installed program's is `org.openstage.guide`), and set `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` to
`--remote-debugging-port=<port>` to drive its page from a script.

## What it needs to build

- Node.js 20 or later (a copy of the one that builds it is carried inside the window, so the computer it is installed on needs nothing)
- Rust with the `x86_64-pc-windows-msvc` toolchain and the Visual Studio C++ build tools
- The Microsoft Edge WebView2 runtime (Windows 11 has it)

## Building

```
cd shell
npm install
npm run build:exe     # the window on its own: src-tauri/target/release/openstage-guide-shell.exe
npm run build         # the installer too: src-tauri/target/release/bundle/nsis/ (not signed)
python -I tools/build_installer.py --sign   # the signed installer: records the release, builds, signs every program of ours and the setup, checks them, writes the SHA-256
```

`prepare.mjs` runs first and gathers the program's source, the acorn package and a copy of Node into `src-tauri/resources`. The first build
compiles a few hundred Rust crates and takes several minutes. The first installer build downloads the NSIS tools.

The setup this makes is **not signed**. Signing is a separate step with the author's own signing account; until it is done Windows will say the
publisher is unknown. A log of what the window did can be had by setting `OPENSTAGE_GUIDE_SHELL_LOG` to a file name before starting it.

The icons in `src-tauri/icons` are made from the product icon (`public/assets/ribbon-*.png`).

## State (2026-11-15)

**Published: v0.1.0-alpha.11** (2026-11-24, a plain release, the latest, with `OpenStage-Guide-0.1.0-alpha.11-Setup.exe` and its `.sha256.txt`; `main` was fast-forwarded to `375da41` and pushed with the tag, and the release was
downloaded again without signing in: same size, same SHA-256 and a valid, timestamped signature by Jason Wrench; the update check, run on the real list of releases, tells alpha 8 and alpha 10 that alpha 11 is out and alpha 11 that it
is the newest. The seven browser checks were run again on the clean checkout of the tag before it was published: 192, 29, 289, 42, 7, 6 and 30.)
(2026-11-23; release label on `eed6d87`, which the lightweight tag `v0.1.0-alpha.11` points at and which the build came from, on the branch `alpha-11`, a clean checkout of that tag in
`C:\Dev\OpenStageGuide-rel11` (`git describe --tags --exact-match HEAD` printed the tag and `git status --porcelain` nothing before it was built); `OpenStage Guide_0.1.0_x64-setup.exe`, 29,277,224 bytes, SHA-256
`2d3f062a317efc643aac829145c83949fc8d674197dcbc155111e509e3c4dfd3`, About reads build `c979a240`), built with `python -I tools/build_installer.py --sign`. It is alpha 10 and the session work of the product specification of
2026-11-22, all four milestones (`docs/SESSION-MATERIALS-2026-11-23.md`): session templates, the 46 timed activities, phases, the run sheet and the preparation checklist, five layouts of the participant workbook, card decks
and a PowerPoint file, outputs that say when the guide has changed since, bringing changes back from a presentation, and the seam for optional AI help (no model is built in). Checked: 2,226 automatic tests on the clean checkout
(2,221 pass; five skipped, all opt-in validators) and the browser checks (192, 29, 289, 7, 6, 42 and the new sessions check's 30; see above for where they were run). Signed and timestamped (8 of our files by Jason Wrench; Node keeps
its makers' signature), re-checked with `sign_windows.py --check`; unpacked with 7-Zip: the 296 program files of `src`, `program` and `public` are identical to the commit (the four the add-in alone uses, `picker.html` and the PDF
library, are left out on purpose by `prepare.mjs`) and `release.txt` says `v0.1.0-alpha.11`. **Not done for this build:** the browser checks were not run again on the unpacked files with the setup's own Node, and the signed
window was not started from the unpacked setup. The release page says so.

**Published: v0.1.0-alpha.10** (2026-11-16, a plain release; pushed to `main` as 684f99a and tagged at the built commit `ab7fd6b`; downloaded again without signing in, same SHA-256 and a valid, timestamped signature, and the update check, run on the real list of releases, tells alpha 8 and alpha 9 that alpha 10 is out and alpha 10 that it is the newest) (2026-11-16, release label on `ee57aa8`, built from `ab7fd6b` on the branch `release-alpha10`, a clean checkout in `C:\Dev\OpenStageGuide-rel`
made from `main` with the alpha 9 fixes merged in (`3102bda`); `OpenStage Guide_0.1.0_x64-setup.exe`, 29,210,984 bytes, SHA-256 `6c1b60853199fc9ab8aafe7f194847b24904622705442deb350cadd5f408d86f`, About reads
build `4c380b62`), built with `python -I tools/build_installer.py --sign`. It is alpha 9 (below, never published) and what was built on `main` after it: the **design desk** (the Design menu: needs and objectives with Bloom's
verbs, models and activities with their citations and an optional connection to Prism Reference Manager, the course map, the storyboard that makes a new guide or a skeleton for Author or Scroll or item slots for
OpenAssess, and the evaluation plan by Kirkpatrick level as a PrismStatistics file), **Insert > Block...**, **File > Save for PrismStatistics...**, bookmarks in the PDF, and the Author export in Author's published format
(0.49.0). Checked: 2,053 automatic tests (2,050 pass; three skipped: Author's validator twice, while Author's unpublished working copy is ahead of its published format, and Scroll's schema, which is opt-in) and the six
browser checks on the clean checkout (192, 29, 289, 7, 6 and the design desk's 42). Signed and timestamped (8 of our files by Jason Wrench; Node keeps its makers' signature), unpacked: 293 files of the program
identical to the commit, `release.txt` says `v0.1.0-alpha.10`, the six browser checks pass on the unpacked files with the setup's own Node (192, 29, 289, 7, 6, 42), and its signed window, started from the unpacked
folder and driven through its debug port, answered with version 0.1.0 and build `4c380b62`, had fifteen menus and opened the design desk, made a tagged PDF with bookmarks with the setup's own Node and Edge, and closed in
2.0 seconds with a change not yet saved, saving it first. Not opened in a real Word, not run against the real Prism Reference Manager, PrismStatistics, OpenAssess, Author or Scroll. Not installed by the author. The release notes are in `D:\Claude Audio Test\installers-signed-2026-11-16-alpha10`.

**Alpha 9 (built and signed, never published; superseded by alpha 10, which holds everything in it):** (2026-11-16, release label on 7b08e59, built from 0a98a77 on the branch `alpha9-fixes`, a clean checkout in `C:\Dev\OpenStageGuide-a9` so that no other session's
uncommitted work got in; `OpenStage Guide_0.1.0_x64-setup.exe`, 29,172,336 bytes, SHA-256
`d3f4f6e99668cb5abe212b634873346984d524bf5b301ef6b2f67825e9ccba03`, About reads build `e8023e92`), built with `python -I tools/build_installer.py --sign`. It is alpha 8 and the seven pieces built after it:
**File > Save as PDF...** (a tagged PDF made by Edge on this computer), **Bullets**, **moving a whole module or lesson** (drag, Alt+Shift+Up or Down, Edit menu), **pictures inside table cells**,
**copying a picture between guides**, and the Layout tool's **icon column** and **slide picture beside the script** (tables in the Word file only); and the fixes for what the Guide book found in alpha 7:
a block, form or table added from a lesson's Time line or a fact goes after the last fact; the Review notes an Expected answer with no Ask; a planning form ticks and dims the facilitator-only box;
the Glossary and References headings and the licence page reach the participant workbook; the Materials list files a facilitator-only form under The facilitator; Ctrl+P prints; the tools and About speak
of the guide, not of a document or the Word add-in. Before the first build the browser checks found a lost update in the Layout tool (fixed in 0750e48, with a test that fails without the fix) and two timing
races in the checks themselves. The first build of this alpha (sha `7ce745cd`, build `df04cfed`) was set aside and rebuilt after the Guide book fixes. Checked: 1,969 automatic tests (one opt-in skipped) and
the five browser checks on the clean checkout (186, 29, 278, 7, 6). Signed and timestamped (8 of our files by Jason Wrench; Node keeps its makers' signature), unpacked: 274 files of the program identical to
the commit, `release.txt` says `v0.1.0-alpha.9`, the five browser checks pass on the unpacked files with the setup's own Node (186, 29, 278, 7, 6), and the Word files the program makes pass the Word schema
validator (icon column included). Its signed window, started from the unpacked folder and driven through its debug port, answered with version 0.1.0 and build `e8023e92`, had the new menu items, made a tagged
PDF with the setup's own Node and Edge, moved a module with Edit > Move module or lesson down and took it back with one Ctrl+Z, moved one by dragging its heading (the mouse movement was sent by a test
program, not a hand), put a form in from a lesson's Time line after the lesson's fact, and closed in 1.3 seconds with a change not yet saved, saving it first. Not opened in a real Word: the layout, bullet and
cell-picture files. Not pushed, not published, not installed by the author.

**Latest published: v0.1.0-alpha.8**, published as a release (not marked pre-release; tag on b89c28f; on GitHub `OpenStage-Guide-0.1.0-alpha.8-Setup.exe`, the same bytes) (2026-11-15, release label on b89c28f; `OpenStage Guide_0.1.0_x64-setup.exe`, 29,170,672 bytes, SHA-256
`189f62accb4e9ba8a831fd1eb6b33a2daf7732931d072eefb2bf5be26431c71f`, About reads build `5105a1e7`), built with `python -I tools/build_installer.py --sign`. It is alpha 7
and one change: **Clear previous guides** on the home screen leaves the guides off the list (a time is kept as a setting; no guide file is deleted, moved or changed), a
button brings them back, and Open guide still lists every guide. Signed and timestamped as below (8 of our files by Jason Wrench; Node keeps its makers' signature), unpacked:
272 files of the program identical to the commit, `release.txt` says `v0.1.0-alpha.8`, the five browser checks pass on the unpacked files with the setup's own Node (173, 29,
234, 7, 6), and its signed window, started from the unpacked folder, answered with version 0.1.0 and build `5105a1e7` and its home screen had the new button (not pressed:
it would have written a setting in the author's folder). Downloaded again from GitHub without signing in, the file has the same hash and a valid, timestamped signature, and a copy that says it is alpha 7 asked Check for updates and was told alpha 8 is out. Not yet installed by the author.

**v0.1.0-alpha.7**, published the same day as a pre-release (tag on 3baba18; on GitHub `OpenStage-Guide-0.1.0-alpha.7-Setup.exe`, the same bytes) (release label on 3baba18; `OpenStage Guide_0.1.0_x64-setup.exe`, 29,168,344 bytes, SHA-256
`3fc537c56221da757d90c284e4c8d747b99d57293812b0c23e733951ffe4a8a0`, About reads build `0bd2e09a`), built with `python -I tools/build_installer.py --sign`. It holds alpha 6's
fixes (menus stay inside the window, one name per tool) and the whole kit written on 2026-11-15 (lesson details, a wider time model, script mechanics, reference pages, the
planning kit and Materials list, the Layout tool and one-page summary, import of a Word file and pasted text, Save as web page, Print, table columns and joined cells,
table titles, picture details; see `docs/KIT-SPEC-2026-11-15.md`). Signed and timestamped as below (8 of our files by Jason Wrench; Node keeps its makers' signature),
unpacked: 272 files of the program identical to the commit, `release.txt` says `v0.1.0-alpha.7`, the five browser checks pass on the unpacked files with the setup's own
Node (162, 29, 234, 7, 6), and its signed window, started from the unpacked folder, answered with version 0.1.0 and build `0bd2e09a` and its menus held the new File
and Make items. Downloaded again from GitHub without signing in, the file has the same hash and a valid, timestamped signature, and a copy that says it is alpha 5 asked Check for updates and was told alpha 7 is out (a copy that says alpha 7 was told it is the newest). Not yet installed or seen by the author. Alpha 6 (built and signed, `installers-signed-2026-11-14-alpha6`) is skipped: it will not be published.

**v0.1.0-alpha.5**, published the same day as a pre-release (tag on 8bd69c8; on GitHub `OpenStage-Guide-0.1.0-alpha.5-Setup.exe`, 29,099,424 bytes,
SHA-256 `68fa136c63bec5883f13bfda14b10bf9811631d094b1aaae7bb9925b7bcbd028`, About reads build `b1b05fdd`), built with `python -I tools/build_installer.py --sign`. It adds
only the new wording of the Accessibility screen (the Word check is advised for a Word file saved from the guide; Go there says "in the guide"). Signed and timestamped as below
(8 of our files by Jason Wrench; Node keeps its makers' signature), unpacked: 259 files of the program identical to the commit, `release.txt` says `v0.1.0-alpha.5`, the five
browser checks pass on the unpacked files with the setup's own Node (112, 29, 217, 7, 6), and its signed window, started from the unpacked folder, showed the fourteen menus
on one row at 1360 by 900. Downloaded again from GitHub without signing in, the file has the same hash and a valid signature, and a copy that says it is alpha 4 asked
Check for updates and was told alpha 5 is out. Not yet seen installed by the author.

**v0.1.0-alpha.4**, published the same day as a pre-release (tag on 8af0500; on GitHub `OpenStage-Guide-0.1.0-alpha.4-Setup.exe`, 29,092,600 bytes,
SHA-256 `b174758911e0395313f37e1678c476a113fd0b61c1571f0798cbaeaf1f36f236`, About reads build `e27dca16`), built with `python -I tools/build_installer.py --sign`. It adds
the fourteen menus (Start, Sections, Activities, Slides, Assessment, eLearning, Check & Plan, Make between View and Help; `PROGRAM_MENUS` in `src/guide-tools.mjs`), the four
tools only the program has (Add Lesson, Learning Outcomes, E-Learning Cues, Accessibility) and the blocks' icons drawn on the guide's page. Signed and timestamped as below
(8 of our files by Jason Wrench; Node keeps its makers' signature), unpacked: 259 files of the program identical to the commit, `release.txt` says `v0.1.0-alpha.4`, the five
browser checks pass on the unpacked files with the setup's own Node (112, 29, 217, 7, 6), and its signed window, started from the unpacked folder, showed the fourteen menus
on one row at 1360 by 900 with the icons on the page. Downloaded again from GitHub without signing in, the file has the same hash and a valid signature, and a copy
that says it is alpha 3 asked Check for updates and was told alpha 4 is out. Not yet seen installed by the author.

**v0.1.0-alpha.3**, published the same day as a pre-release (tag on 7120c09; on GitHub `OpenStage-Guide-0.1.0-alpha.3-Setup.exe`, 29,088,176 bytes,
SHA-256 `e59eda6e05d07c7625c37087ef31ca87de2bd27734778a97cc483f72e366ff2f`), built with `python -I tools/build_installer.py --sign`. It adds zoom for the guide's page (View menu,
the formatting bar, Ctrl+Alt+= - 0, Ctrl and the wheel over the guide) and the fix for the empty space below a long guide. Signed and timestamped as below (8 of our files by
Jason Wrench; Node keeps its makers' signature), unpacked: 259 files of the program identical to the commit, `release.txt` says `v0.1.0-alpha.3`, and the server in its own
files, with its own Node, answered, kept a setting and told the window the page had saved. Downloaded again from GitHub without signing in, the file has the same hash and a
valid signature, and copies of alpha 1 and 2 asked Check for updates and were told alpha 3 is out. Not yet seen installed by the author.

**v0.1.0-alpha.2**, published the same day as a pre-release (tag on 8463662; on GitHub `OpenStage-Guide-0.1.0-alpha.2-Setup.exe`, 29,093,272 bytes,
SHA-256 `0ce8a60f71db498246ca5ee782fc46babe96f684ad35944552b19d7441dc2e3b`), built with `python -I tools/build_installer.py --sign`. It adds what the
program remembers in a file beside the guides, one guides folder for the window and the browser, saving before the window closes, and no "prototype"
in what the program shows. Signed and timestamped as below (8 of our files by Jason Wrench; Node keeps its makers' signature), unpacked: 258 files of
the program identical to the commit, `release.txt` says `v0.1.0-alpha.2`, and the server in its own files, with its own Node, answered, kept a setting and told the
window the page had saved. Downloaded again from GitHub without signing in, the file has the same hash and a valid signature, and a copy of alpha 1
asked Check for updates and was told alpha 2 is out. Not yet installed or opened by anyone.

The first signed setup, v0.1.0-alpha.1:

The signed setup is built: `python -I tools/build_installer.py --sign`, from commit 80a6eda, release `v0.1.0-alpha.1` (recorded in `release.txt` for Check
for updates). `OpenStage Guide_0.1.0_x64-setup.exe`, 29,100,176 bytes, SHA-256
`021df4a19d5171bf334ad9b8eca11b30919f993ead07ea2f3f200cb57ab9009b`. The setup and every program of ours in it (the window, the uninstaller and the five
installer plugins) are signed by Jason Wrench and timestamped; the Node that runs the server keeps the OpenJS Foundation's valid signature; the program
holds no folder name of the computer that built it. Its uninstaller deletes only the files the setup put in the install folder, and never the guides in
`Documents\OpenStage Guides`. This setup was published on 2026-11-14 as the pre-release v0.1.0-alpha.1 (on GitHub the file is named
`OpenStage-Guide-0.1.0-alpha.1-Setup.exe`, the same bytes) and was installed and opened by the author.

The setup built from a08348d (SHA-256 `263e8c60...7bc0`) is superseded and must not be used: its file list left out `scripts/updates.mjs`, which the server
loads, so an installed copy would not have started. 80a6eda adds it, and `tests/shell-prepare.test.mjs` fails if the server ever loads a script, source folder
or package that `prepare.mjs` does not carry. The new setup was unpacked and the server in its own files was started with its own Node: `/health`, the
program page, the tools pane and `/api/updates` answered, and the add-in's file window and PDF library were not there.

Earlier (2026-11-13), built and tried by script on the author's computer: the window starts the carried Node, the server says where it is, the page
answers, and the server stops when the window is killed.
