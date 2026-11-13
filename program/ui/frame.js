// @ts-check
/**
 * The OpenStage suite's frame round every screen: the navy bar along the top (the icon and the name, the guide that is open and whether it is
 * saved, the command search and the guide's main buttons), the menus under it, and the two windows the suite shares, the command search
 * (Ctrl+K) and the list of keyboard keys (Ctrl+/). Every command is in the one list of commands.js, so a command has the same words, icon
 * and keys in a menu, in the command search and on a button. After OpenAssess's frame (see NOTICE).
 */
import { buildMenus } from "./commands.js";
import { openCommandSearch, searchEntries } from "./commandsearch.js";
import { openKeys } from "./keys.js";
import { MenuBar, el, uiIcon } from "./menubar.js";

const SAVE_WORDS = /** @type {Record<string, string>} */ ({ new: "No changes yet", saving: "Saving", saved: "Saved", unsaved: "Not saved yet", error: "Not saved" });

/**
 * Build the frame into the page, once.
 * @param {Document} doc
 * @param {{ edit?: import("./commands.js").EditApi, insert?: (name: string) => unknown, guide: () => { name: string } | null, openTool: (id: string) => void, updateGuide: () => void, newGuide: () => void, openGuide: () => void, closeGuide: () => void, about: () => void }} shell
 */
export function startFrame(doc, shell) {
  const contextHost = /** @type {HTMLElement} */ (doc.getElementById("bar-context"));
  const buttonsHost = /** @type {HTMLElement} */ (doc.getElementById("bar-buttons"));
  const searchHost = /** @type {HTMLElement} */ (doc.getElementById("bar-search"));
  const menuHost = /** @type {HTMLElement} */ (doc.getElementById("menu-row"));
  const main = /** @type {HTMLElement} */ (doc.getElementById("main"));
  let saveState = "";

  const search = () => openCommandSearch(doc, searchEntries(menus), { fallback: () => main });
  const keys = () => openKeys(doc);
  const { menus, quick } = buildMenus({ ...shell, hasGuide: () => shell.guide() !== null, search, keys });
  const bar = new MenuBar(doc, menus);
  menuHost.append(bar.element);

  // The guide stays in sight under the navy bar and the menus, which are taller when the window is narrow and they wrap.
  const top = doc.querySelector(".top");
  if (top && typeof ResizeObserver === "function") {
    new ResizeObserver(() => doc.documentElement.style.setProperty("--top-height", `${Math.round(top.getBoundingClientRect().height)}px`)).observe(top);
  }

  const searchButton = el(doc, "button", { type: "button", class: "bar-button", id: "bar-search-button", "aria-keyshortcuts": "Control+K", "aria-haspopup": "dialog", title: "Search for a command (Ctrl+K)" }, uiIcon(doc, "search"), el(doc, "span", { text: "Search" }));
  searchButton.setAttribute("aria-label", "Search for a command");
  searchButton.addEventListener("click", search);
  searchHost.append(searchButton);

  // Ctrl+K and Ctrl+/ work anywhere, except over a dialog that is asking something (the search itself takes Ctrl+K back to its box).
  doc.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
    const asking = [...doc.querySelectorAll("dialog[open]")].some((d) => !d.classList.contains("command-search"));
    if (e.key.toLowerCase() === "k" && !asking) {
      e.preventDefault();
      search();
    } else if (e.key === "/" && !asking) {
      e.preventDefault();
      keys();
    }
  });

  /** Draw the navy bar for what is open now: the guide's name and whether it is saved, and its main buttons. */
  const drawBar = () => {
    const guide = shell.guide();
    contextHost.replaceChildren();
    buttonsHost.replaceChildren();
    if (!guide) return;
    contextHost.append(
      el(doc, "span", { class: "bar-context-label", text: "Guide" }),
      el(doc, "span", { class: "bar-context-name", text: guide.name, title: guide.name }),
      el(doc, "span", { class: "bar-context-state", "data-state": saveState, text: SAVE_WORDS[saveState] ?? "" }),
    );
    if (!SAVE_WORDS[saveState]) contextHost.lastElementChild?.setAttribute("hidden", "");
    quick.forEach((cmd, i) => {
      const button = el(doc, "button", { type: "button", class: i === quick.length - 1 ? "bar-button primary" : "bar-button", title: cmd.title ?? cmd.label }, cmd.icon ? uiIcon(doc, cmd.icon) : "", el(doc, "span", { text: cmd.label }));
      button.addEventListener("click", () => cmd.run());
      buttonsHost.append(button);
    });
  };

  return {
    menus,
    bar,
    /** The guide that is open changed (or was closed): draw the bar again. */
    refresh: drawBar,
    /** The guide's save state changed ("saved", "saving", "unsaved", "error", "new"): the bar says so in words. */
    setSaveState(/** @type {string} */ state) {
      saveState = state;
      drawBar();
    },
    search,
    keys,
  };
}
