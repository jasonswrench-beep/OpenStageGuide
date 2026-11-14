#!/usr/bin/env python3
# SPDX-License-Identifier: Apache-2.0
"""Builds the OpenStage Guide setup, and with --sign signs it.

  python -I shell/tools/build_installer.py            build the setup (not signed)
  python -I shell/tools/build_installer.py --sign     build it and sign every program of ours in it, and the setup itself
  python -I shell/tools/build_installer.py --check SETUP   only check a setup made before: signatures, and no folder names of this computer

What it does:
  1. Records the release this build is (RELEASE, below) for the window to give the server, so that Check for updates can compare it with the
     published ones (shell/prepare.mjs writes it beside the program's files).
  2. Builds with the repository's and Cargo's folders remapped, so that the program holds no folder name of the computer that built it.
  3. With --sign: Tauri signs, through its sign command (sign_windows.py, beside this file), the window, the installer's plugins, the uninstaller and
     the setup. Then the setup is unpacked with 7-Zip and every program in it, and the setup, must be signed and timestamped by the same
     publisher (OPENSTAGE_SIGN_PUBLISHER, when it is set). The Node that runs the server keeps its makers' signature, which must be valid.
  4. Writes the setup's SHA-256 beside it (after signing: it is always the hash of the file that is published).

The signing account comes from the environment (the top of sign_windows.py says which variables); nothing about it is written into the repository.
Without --sign anyone can build from source, and the setup is not signed.
"""

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SHELL = HERE.parent
REPO = SHELL.parent
TAURI = SHELL / "src-tauri"
SETUPS = TAURI / "target" / "release" / "bundle" / "nsis"
EXE = TAURI / "target" / "release" / "openstage-guide-shell.exe"
SIGN_SCRIPT = HERE / "sign_windows.py"
SEVEN_ZIP = Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) / "7-Zip" / "7z.exe"

# The release this build is. Raise it for every release: the installed copy compares it with the published ones (docs/UPDATE-CHECK.md).
RELEASE = "v0.1.0-alpha.2"

# Programs in the setup that their makers sign, which are never signed here: the Node that runs the server (the OpenJS Foundation's).
MAKERS = {"node.exe": "OpenJS Foundation"}
PROGRAM_FILES = (".exe", ".dll", ".pyd")


def say(text):
    print(text, flush=True)


def fail(text):
    print("STOPPED: " + text, flush=True)
    sys.exit(1)


def sign_module():
    sys.path.insert(0, str(HERE))
    import sign_windows  # noqa: E402  (beside this file)

    return sign_windows


def sign_config():
    """Tauri's sign command for this build, in a file Tauri reads with --config: this Python runs sign_windows.py on each file Tauri signs."""
    config = {"bundle": {"windows": {"signCommand": {"cmd": sys.executable, "args": [str(SIGN_SCRIPT), "--description", "OpenStage Guide", "%1"]}}}}
    path = TAURI / "target" / "signing" / "tauri-sign.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(config, indent=2) + "\n", encoding="utf-8")
    return path


def remap_flags():
    """Rust's own flags that write the build folders as short names, so the program holds no folder of this computer."""
    cargo = os.environ.get("CARGO_HOME") or str(Path.home() / ".cargo")
    rustup = os.environ.get("RUSTUP_HOME") or str(Path.home() / ".rustup")
    flags = [f"--remap-path-prefix={REPO}=guide", f"--remap-path-prefix={cargo}=cargo", f"--remap-path-prefix={rustup}=rustup"]
    return " ".join(flags)


def find_setup():
    found = sorted(SETUPS.glob("*-setup.exe"), key=lambda p: p.stat().st_mtime)
    return found[-1] if found else None


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def holds_folder_names(path):
    """The names of this computer's folders that the file holds (in either of the two ways a Windows program writes text), or []."""
    data = Path(path).read_bytes()
    found = []
    for name in {Path.home().name, os.environ.get("USERNAME", "")} - {""}:
        for text in (f"\\Users\\{name}", f"/Users/{name}"):
            if text.encode("utf-8").lower() in data.lower() or text.encode("utf-16-le").lower() in data.lower():
                found.append(text)
    return sorted(set(found))


def unpack(setup, into):
    if not SEVEN_ZIP.is_file():
        fail(f"7-Zip is not at {SEVEN_ZIP}; it is needed to look inside the setup.")
    done = subprocess.run([str(SEVEN_ZIP), "x", "-y", f"-o{into}", str(setup)], capture_output=True, text=True, encoding="utf-8", errors="replace")
    if done.returncode not in (0, 1):  # 1 is a warning (7-Zip warns about the NSIS script it does not unpack)
        fail("7-Zip could not unpack the setup:\n" + done.stdout[-800:] + done.stderr[-400:])


def check_installer(setup, signed):
    """Check a setup: no folder name of this computer in the window's program, and (with signed) the setup and every program in it signed and timestamped by
    one publisher, except the makers' files, which must be validly signed by their makers. Returns a list of problems."""
    problems = []
    sw = sign_module()
    for path in (EXE,):
        if path.is_file():
            names = holds_folder_names(path)
            if names:
                problems.append(f"{path.name} holds this computer's folder names: {', '.join(names)}")
    if not signed:
        return problems
    with tempfile.TemporaryDirectory(prefix="openstage-guide-setup-") as temp:
        unpack(setup, temp)
        inside = sorted(p for p in Path(temp).rglob("*") if p.is_file() and p.suffix.lower() in PROGRAM_FILES)
        if not inside:
            return problems + ["no program was found inside the setup, so none could be checked"]
        found = sw.signatures([setup, *inside])
        publisher = os.environ.get("OPENSTAGE_SIGN_PUBLISHER") or found[0]["signer"]
        ours, theirs = [], []
        for info in found:
            name = Path(info["path"]).name
            where = "the setup" if Path(info["path"]) == Path(setup).resolve() else name
            maker = MAKERS.get(name.lower())
            if maker and where != "the setup":
                if info["status"] != "Valid":
                    problems.append(f"{where}: its maker's signature is {info['status']}")
                elif maker not in (info["signer"] or ""):
                    problems.append(f"{where}: signed by {info['signer']}, not by {maker}")
                theirs.append(f"{name} (by {info['signer']})")
                continue
            why = sw.problem(info, publisher)
            if why:
                problems.append(f"{where}: {why}")
            ours.append(name)
        say(f"  {Path(setup).name}: {len(ours)} files signed by {publisher} ({', '.join(sorted(set(ours)))}); kept their makers' signatures: {', '.join(theirs) or 'none'}")
    return problems


def build(sign):
    npm = shutil.which("npm.cmd") or shutil.which("npm")
    if not npm:
        fail("npm was not found; it runs the build.")
    env = dict(os.environ)
    env["OPENSTAGE_GUIDE_RELEASE"] = RELEASE
    env["RUSTFLAGS"] = (env.get("RUSTFLAGS", "") + " " + remap_flags()).strip()
    command = [npm, "run", "build", "--"]
    if sign:
        command += ["--config", str(sign_config())]
    say("Building (the first build compiles a few hundred Rust crates and takes several minutes)...")
    done = subprocess.run(command, cwd=SHELL, env=env)
    if done.returncode != 0:
        fail("the build failed (see above).")


def main(argv):
    args = list(argv)
    if args[:1] == ["--check"]:
        if len(args) != 2 or not Path(args[1]).is_file():
            fail("--check needs the setup file after it.")
        problems = check_installer(Path(args[1]), signed=True)
        for p in problems:
            say("  PROBLEM: " + p)
        return 1 if problems else 0
    sign = "--sign" in args
    say(f"OpenStage Guide {RELEASE}: building the installer{' and signing it' if sign else ''}.")
    if sign:
        sw = sign_module()
        try:
            sw.settings()
            sw.find_signtool()
            sw.find_dlib()
        except sw.SignError as e:
            fail("signing is not ready: " + str(e))
    build(sign)
    setup = find_setup()
    if setup is None:
        fail(f"no setup was found in {SETUPS}.")
    problems = check_installer(setup, signed=sign)
    if problems:
        fail("the setup did not pass: " + "; ".join(problems))
    digest = sha256(setup)
    Path(str(setup) + ".sha256.txt").write_text(f"{digest} *{setup.name}\n", encoding="utf-8")
    say(f"  {setup.name}: {setup.stat().st_size:,} bytes, SHA-256 {digest}{', signed and timestamped' if sign else ', NOT signed'}")
    say(f"  {setup}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
