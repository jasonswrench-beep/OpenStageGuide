# Check for updates

OpenStage Guide has **Help > Check for updates...**. It follows the pattern every tool of the OpenStage family uses for this button
(OpenAssess, OpenStage Author, OpenStage Recorder, OpenExplainer and PDF Remediator have the same one).

## What it does

* **Only when chosen.** Nothing is checked at start-up or on a timer, and nothing runs in the background.
* **Notify only.** It says whether a newer release is published and, if so, links to that release's page on GitHub (what changed, the
  checksum, the signed setup). It downloads and installs nothing. Running the new setup updates the copy in place and keeps the guides.
* **Where it looks.** One HTTPS request, made by the program's own server, to
  `https://api.github.com/repos/jasonswrench-beep/OpenStageGuide/releases?per_page=30`. Node checks the certificate. No redirect is followed,
  at most 2 MB is read, and it gives up after 15 seconds. Pre-releases count, because every release so far is one; drafts never do. The
  request carries only a plain user agent (`OpenStage-Guide-update-check`). Nothing about the person, their guides or their computer is sent.
  A repository that is not public yet answers 404, which is said as "no release found".
* **What it compares.** The installed copy's exact release with the highest published tag (`vMAJOR.MINOR.PATCH`, or that with `-alpha.N`,
  `-beta.N` or `-rc.N`; a release sorts after its own pre-releases). The release is recorded when the setup is built: `RELEASE` in
  `shell/tools/build_installer.py` becomes `OPENSTAGE_GUIDE_RELEASE`, which `shell/prepare.mjs` writes to `release.txt` beside the program's
  files, which the window passes to the server as `--release`. Raise it for every release. A copy not made by the setup (the browser version from
  `Start OpenStage Guide.cmd`, or a build made another way) is told which release is the newest.
* **Links.** Only a page of OpenStage Guide's releases (the list, or `.../releases/tag/v...`) is ever offered. The server builds the address,
  the page checks it again before showing it as a plain link, and the window opens it in the person's own browser and never inside itself. A
  browser opens it in a new tab. (The window also opens the icon license page, which the About text and the tools name; it opens no other
  address outside the program.)
* **Turning it off.** Set `OPENSTAGE_UPDATE_CHECK=off` in the environment, or put a file `update-check.txt` that says `off` beside
  `openstage-guide-shell.exe` (the way an IT department turns it off for the copies it installs). `--no-update-check` does the same for the
  server started from the command line. The answer then says that checking is turned off on this computer, and nothing is asked of the network.
* **Signing.** No update-signing key is needed, because nothing is installed automatically. The setup itself is signed with the publisher's
  code-signing certificate, and the release page carries its SHA-256.

## Code and tests

* `scripts/updates.mjs`: the comparison, the request and the answers. The route is `POST /api/updates` in `scripts/project-server.mjs`. It is a
  POST in the form the program sends (same site, JSON), like every action. `--release` and `--no-update-check` are handled at the foot of that file.
* `program/ui/updates.js` and the Help menu in `program/ui/commands.js`: the dialog, which shows the answer as text.
* `shell/src-tauri/src/main.rs`: the release passed to the server, the switch, and the opening of a release page in the browser
  (`is_release_page`, `navigate_to`).
* `tests/updates.test.mjs` covers the comparison, the answers, the links, the exact request (with a stand-in for the network) and the switch;
  `tests/program-server.test.mjs` covers the route and the command line; `tests/program-commands.test.mjs` the menu; the program check
  (`scripts/check-program-browser.mjs`) presses the button in a real browser with checking turned off. Nothing reaches the network.
  `cargo test` in `shell/src-tauri` covers the window's address rules.
