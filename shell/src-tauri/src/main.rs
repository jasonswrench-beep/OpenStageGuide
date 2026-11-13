//! The OpenStage Guide window.
//!
//! Everything the program does is done by its local server, which is Node running the program's own code (`scripts/project-server.mjs`).
//! This starts that server, waits for it to say where it is listening, and points the window at it. There is nothing else in here on
//! purpose: the window has no commands for the page to call, and the page it shows is given no access to the machine beyond what the
//! server allows it. When the window closes the server is stopped, and if the window is killed the server notices its input has
//! closed and stops itself.
//!
//! The shape of this follows the OpenAssess window (the same owner, Apache-2.0), without its Save dialog: a file the program makes (the
//! participant workbook, a backup) is saved to the Downloads folder by the window, and the page says so.

#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::fs::OpenOptions;
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::Mutex;

use tauri::{Manager, RunEvent, WebviewWindowBuilder};

#[cfg(windows)]
use std::os::windows::process::CommandExt;

/// The server, kept so that it can be stopped when the window goes.
struct Server(Mutex<Option<Child>>);

/// A path without Windows' extended-length prefix (backslash, backslash, question mark, backslash), which Tauri puts on the folders it reports and which Node does not take for the same path as
/// the one its own module loader sees (the server then does not start).
fn plain(path: PathBuf) -> PathBuf {
    let text = path.to_string_lossy();
    match text.strip_prefix(r"\\?\") {
        Some(rest) if !rest.starts_with("UNC\\") => PathBuf::from(rest),
        _ => path,
    }
}

/// The folder that holds `scripts/project-server.mjs`: the one the installer put beside the window, one named by the person running it,
/// or, when it is run from a build in the repository, the repository.
fn app_root(resources: &Path) -> Option<PathBuf> {
    let mut candidates: Vec<PathBuf> = Vec::new();
    if let Ok(root) = std::env::var("OPENSTAGE_GUIDE_APP_ROOT") {
        candidates.push(PathBuf::from(root));
    }
    candidates.push(resources.join("resources").join("app"));
    candidates.push(resources.join("app"));
    if let Ok(exe) = std::env::current_exe() {
        for up in 1..=8 {
            if let Some(dir) = exe.ancestors().nth(up) {
                candidates.push(dir.to_path_buf());
            }
        }
    }
    candidates
        .into_iter()
        .find(|c| c.join("scripts").join("project-server.mjs").exists())
}

/// The Node that runs the server: the one carried with the window, or the one on the path.
fn node_path(resources: &Path) -> PathBuf {
    let name = if cfg!(windows) { "node.exe" } else { "node" };
    for dir in [resources.join("resources").join("node"), resources.join("node")] {
        let candidate = dir.join(name);
        if candidate.exists() {
            return candidate;
        }
    }
    PathBuf::from(name)
}

/// Says in the window what went wrong, in words, for a failure to start.
fn say(window: &tauri::WebviewWindow, message: &str) {
    let safe: String = message
        .chars()
        .filter(|c| !matches!(c, '\'' | '"' | '\\' | '<' | '>' | '`' | '\n' | '\r'))
        .collect();
    let _ = window.eval(&format!("document.getElementById('note').textContent = '{}';", safe));
}

/// A line in the log file, when one was asked for with `OPENSTAGE_GUIDE_SHELL_LOG`, for finding out what the window did.
fn note(message: &str) {
    if let Some(mut file) = std::env::var("OPENSTAGE_GUIDE_SHELL_LOG")
        .ok()
        .and_then(|p| OpenOptions::new().create(true).append(true).open(p).ok())
    {
        let _ = writeln!(file, "{message}");
    }
}

fn main() {
    let app = tauri::Builder::default()
        .setup(|app| {
            let config = app
                .config()
                .app
                .windows
                .first()
                .cloned()
                .expect("the configuration names the window");
            let window = WebviewWindowBuilder::from_config(app.handle(), &config)?.build()?;
            let resources = plain(app.path().resource_dir().unwrap_or_else(|_| PathBuf::from(".")));

            note(&format!("resources: {}", resources.display()));
            let Some(root) = app_root(&resources) else {
                note("no folder with scripts/project-server.mjs was found");
                say(&window, "OpenStage Guide could not find its own files. Reinstall it.");
                return Ok(());
            };
            note(&format!("program folder: {}", root.display()));
            let root = plain(root);
            let script = root.join("scripts").join("project-server.mjs");

            // Where guides are kept: a folder in the person's own Documents (wherever Windows has put it, OneDrive included).
            let guides = plain(app.path().document_dir().unwrap_or_else(|_| PathBuf::from("."))).join("OpenStage Guides");

            let mut command = Command::new(node_path(&resources));
            command
                .arg(&script)
                .arg("--port")
                .arg("0")
                .arg("--folder")
                .arg(&guides)
                .arg("--exit-when-stdin-closes")
                .current_dir(&root);
            command.stdin(Stdio::piped()).stdout(Stdio::piped()).stderr(Stdio::null());
            // No console window of its own beside the program's window.
            #[cfg(windows)]
            command.creation_flags(0x0800_0000);

            note(&format!("starting {} with guides in {}", node_path(&resources).display(), guides.display()));
            let mut child = match command.spawn() {
                Ok(child) => child,
                Err(error) => {
                    note(&format!("the server could not be started: {error}"));
                    say(&window, &format!("OpenStage Guide could not start its server: {error}"));
                    return Ok(());
                }
            };
            let Some(stdout) = child.stdout.take() else {
                say(&window, "OpenStage Guide could not read from its server.");
                return Ok(());
            };
            app.manage(Server(Mutex::new(Some(child))));

            std::thread::spawn(move || {
                let mut shown = false;
                for line in BufReader::new(stdout).lines().map_while(Result::ok) {
                    note(&line);
                    if shown {
                        continue;
                    }
                    if let Some(url) = line.strip_prefix("OPENSTAGE_READY ") {
                        // Only the address the server gave, and only if it is this machine's own: nothing else is ever opened.
                        let plain = url.chars().all(|c| c.is_ascii_alphanumeric() || ":/.?=-_".contains(c));
                        if url.starts_with("http://127.0.0.1:") && plain {
                            let _ = window.eval(&format!("window.location.replace('{url}');"));
                            shown = true;
                        } else {
                            say(&window, "OpenStage Guide's server gave an address it will not open.");
                        }
                    }
                }
            });
            Ok(())
        })
        .build(tauri::generate_context!())
        .expect("the window could not be built");

    app.run(|handle, event| {
        if let RunEvent::Exit = event {
            if let Some(server) = handle.try_state::<Server>() {
                if let Ok(mut guard) = server.0.lock() {
                    if let Some(mut child) = guard.take() {
                        let _ = child.kill();
                    }
                }
            }
        }
    });
}
