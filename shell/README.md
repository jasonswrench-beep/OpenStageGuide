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

## State (2026-11-13)

Built and tried by script on the author's computer: the window starts the carried Node, the server says where it is, the page answers, and the
server stops when the window is killed. The unsigned setup builds (27 MB, `src-tauri/target/release/bundle/nsis/`). The setup itself has not been
run by anyone, nothing has been signed, and nothing has been published.
