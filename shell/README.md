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

## State (2026-11-14)

**Latest: v0.1.0-alpha.5**, published the same day as a pre-release (tag on 8bd69c8; on GitHub `OpenStage-Guide-0.1.0-alpha.5-Setup.exe`, 29,099,424 bytes,
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
