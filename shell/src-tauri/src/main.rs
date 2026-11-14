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
use std::sync::atomic::{AtomicU16, Ordering};
use std::sync::Mutex;

use tauri::webview::NewWindowResponse;
use tauri::{Manager, RunEvent, Url, WebviewWindowBuilder};

#[cfg(windows)]
use std::os::windows::process::CommandExt;

/// The server, kept so that it can be stopped when the window goes.
struct Server(Mutex<Option<Child>>);

/// The port the server listens on, set when it says where it is (0 until then). The window is only ever sent to that address on this machine.
static SERVER_PORT: AtomicU16 = AtomicU16::new(0);

/// Whether the window may show this address: its own first page, and the server's address on this machine (the port it said). A web page of
/// any other site is not the program, and is never shown inside the program's window, where it could be mistaken for it. (Pictures and files the
/// program makes arrive as blob: and data: addresses, which stay.)
fn may_open(url: &Url) -> bool {
    match url.scheme() {
        "tauri" | "about" | "blob" | "data" => true,
        "http" | "https" => {
            let host = url.host_str().unwrap_or("");
            let port = SERVER_PORT.load(Ordering::SeqCst);
            host == "tauri.localhost" || (port != 0 && host == "127.0.0.1" && url.port() == Some(port))
        }
        _ => false,
    }
}

/// An address with none of the extras a link can carry (a port, a name and password, a query, a fragment), on https.
fn plain_https(url: &Url, host: &str) -> bool {
    url.scheme() == "https" && url.host_str() == Some(host) && url.port().is_none() && url.username().is_empty() && url.password().is_none() && url.query().is_none() && url.fragment().is_none()
}

/// Whether an address is a page of OpenStage Guide's releases on GitHub (the list, or one release by its tag): what Check for updates offers.
fn is_release_page(url: &Url) -> bool {
    const RELEASES: &str = "/jasonswrench-beep/OpenStageGuide/releases";
    if !plain_https(url, "github.com") {
        return false;
    }
    let path = url.path();
    path == RELEASES
        || path.strip_prefix(RELEASES).and_then(|rest| rest.strip_prefix("/tag/v")).is_some_and(|tag| {
            (1..=40).contains(&tag.len()) && tag.starts_with(|c: char| c.is_ascii_digit()) && tag.chars().all(|c| c.is_ascii_alphanumeric() || c == '.' || c == '-')
        })
}

/// The license page of the training icons, which the About text and the tools name.
fn is_license_page(url: &Url) -> bool {
    plain_https(url, "creativecommons.org") && url.path() == "/licenses/by-nc-sa/4.0/"
}

/// The pages outside the program that the window opens, and then in the person's own browser, never inside itself. Nothing else is, whatever a
/// page asks.
fn opens_in_browser(url: &Url) -> bool {
    is_release_page(url) || is_license_page(url)
}

/// Opens such a page in the person's own browser. The address is one argument, never a command line.
#[cfg(windows)]
fn open_in_browser(url: &Url) {
    if opens_in_browser(url) {
        let _ = Command::new("explorer.exe").arg(url.as_str()).spawn();
    }
}

#[cfg(not(windows))]
fn open_in_browser(_url: &Url) {}

/// What the window does when it is asked to go somewhere: its own pages are let through, a page that belongs in the browser is opened there
/// instead, and anything else is refused. Only the scheme and the host are written to the log.
fn navigate_to(url: &Url) -> bool {
    if may_open(url) {
        return true;
    }
    if opens_in_browser(url) {
        open_in_browser(url);
    } else {
        note(&format!("the window was kept from going to {}://{}", url.scheme(), url.host_str().unwrap_or("")));
    }
    false
}

/// The release this copy is, as the installer's build recorded it (release.txt beside the program's files, from OPENSTAGE_GUIDE_RELEASE), for
/// Check for updates to compare with the published ones. A window built any other way has none, and Check for updates then says which release is
/// the newest. The server checks that it is a version.
fn recorded_release(root: &Path) -> Option<String> {
    std::fs::read_to_string(root.join("release.txt")).ok().map(|text| text.trim().to_string()).filter(|text| !text.is_empty() && text.len() <= 40)
}

/// Whether Check for updates is turned off on this computer: OPENSTAGE_UPDATE_CHECK=off, or a file update-check.txt beside the program that says
/// off (the way an IT department turns it off for the copies it installs). The server is then told so.
fn update_check_off() -> bool {
    let off = |text: &str| text.trim().eq_ignore_ascii_case("off");
    if std::env::var("OPENSTAGE_UPDATE_CHECK").is_ok_and(|v| off(&v)) {
        return true;
    }
    std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(|dir| dir.join("update-check.txt")))
        .and_then(|file| std::fs::read_to_string(file).ok())
        .is_some_and(|text| off(&text))
}

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
            let window = WebviewWindowBuilder::from_config(app.handle(), &config)?
                .on_navigation(navigate_to)
                // A link that asks for a new window (Check for updates' link to the release page) never gets one: a page that belongs in the
                // browser opens there, and anything else is refused.
                .on_new_window(|url, _features| {
                    navigate_to(&url);
                    NewWindowResponse::Deny
                })
                .build()?;
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
            if let Some(release) = recorded_release(&root) {
                command.arg("--release").arg(release);
            }
            if update_check_off() {
                command.arg("--no-update-check");
            }
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
                        let port = url
                            .strip_prefix("http://127.0.0.1:")
                            .and_then(|rest| rest.split('/').next())
                            .and_then(|digits| digits.parse::<u16>().ok());
                        if let (Some(port), true) = (port, plain) {
                            // The window may go to this address (and only this one on this machine) from now on.
                            SERVER_PORT.store(port, Ordering::SeqCst);
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

#[cfg(test)]
mod tests {
    use super::{is_license_page, is_release_page, may_open, opens_in_browser};
    use tauri::Url;

    fn url(text: &str) -> Url {
        Url::parse(text).expect("an address")
    }

    #[test]
    fn the_window_shows_only_its_own_page_and_its_servers_port() {
        assert!(may_open(&url("http://tauri.localhost/index.html")));
        assert!(may_open(&url("tauri://localhost/index.html")));
        assert!(may_open(&url("about:blank")));
        assert!(!may_open(&url("http://127.0.0.1:47321/")), "no port of the server is let through before it says where it is");
        for other in ["https://example.invalid/", "http://example.invalid/", "file:///C:/x.html", "ftp://example.invalid/x", "https://github.com/jasonswrench-beep/OpenStageGuide/releases"] {
            assert!(!may_open(&url(other)), "{other}");
        }
    }

    #[test]
    fn the_pages_that_open_in_the_browser_are_this_programs_releases_and_the_icon_license() {
        for yes in [
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases",
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/v0.1.0-alpha.1",
            "https://GitHub.com/jasonswrench-beep/OpenStageGuide/releases/tag/v1.0.0",
            "https://creativecommons.org/licenses/by-nc-sa/4.0/",
        ] {
            assert!(opens_in_browser(&url(yes)), "{yes}");
        }
        for no in [
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/latest",
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/v",
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/v0.1.0?x=1",
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/v0.1.0#a",
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/download/v0.1.0/setup.exe",
            "https://github.com/jasonswrench-beep/OpenStageGuide/releases/tag/v0.1.0/../../../x",
            "https://github.com/someone/OpenStageGuide/releases",
            "http://github.com/jasonswrench-beep/OpenStageGuide/releases",
            "https://github.com:8443/jasonswrench-beep/OpenStageGuide/releases",
            "https://user@github.com/jasonswrench-beep/OpenStageGuide/releases",
            "https://github.com.example/jasonswrench-beep/OpenStageGuide/releases",
            "https://creativecommons.org/licenses/by/4.0/",
            "https://creativecommons.org/licenses/by-nc-sa/4.0/?x=1",
            "https://example.invalid/jasonswrench-beep/OpenStageGuide/releases",
        ] {
            assert!(!opens_in_browser(&url(no)), "{no}");
        }
        assert!(is_release_page(&url("https://github.com/jasonswrench-beep/OpenStageGuide/releases")));
        assert!(!is_license_page(&url("https://github.com/jasonswrench-beep/OpenStageGuide/releases")));
    }
}
