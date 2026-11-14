// @ts-check
/**
 * Help > Check for updates: asks this program's own server, once, whether a newer OpenStage Guide is published, and says the answer in words.
 * Notify only: nothing is downloaded or installed. The server asks GitHub (scripts/updates.mjs); the page itself reaches for no other address, and
 * offers a link only to a page of OpenStage Guide's releases, which the window opens in the person's own browser. The same button and rules as the
 * rest of the OpenStage family (docs/UPDATE-CHECK.md).
 */
import { post } from "./api.js";
import { openDialog } from "./dialogs.js";
import { el } from "./menubar.js";

/** The path of OpenStage Guide's releases on GitHub. The server gives the whole address; this is only what it must be. */
export const RELEASES_PATH = "/jasonswrench-beep/OpenStageGuide/releases";

/**
 * The address to offer, when it is a page of OpenStage Guide's releases (the list, or one release by its tag), or null.
 * @param {unknown} url
 * @returns {string | null}
 */
export function releaseLink(url) {
  if (typeof url !== "string") return null;
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return null;
  }
  if (parsed.protocol !== "https:" || parsed.hostname !== "github.com" || parsed.port !== "" || parsed.username !== "" || parsed.password !== "" || parsed.search !== "" || parsed.hash !== "") return null;
  const path = parsed.pathname;
  const tag = path.startsWith(`${RELEASES_PATH}/tag/`) ? path.slice(`${RELEASES_PATH}/tag/`.length) : null;
  if (path === RELEASES_PATH || (tag !== null && /^v\d{1,6}\.\d{1,6}\.\d{1,6}(?:-(?:alpha|beta|rc)\.\d{1,6})?$/.test(tag))) return parsed.href;
  return null;
}

/**
 * What the dialog says, from the server's answer.
 * @param {any} answer
 * @returns {{ words: string, link: string | null, linkLabel: string }}
 */
export function describeAnswer(answer) {
  const words = typeof answer?.message === "string" && answer.message !== "" ? answer.message : "The answer could not be read. Try again later.";
  const link = answer?.state === "off" ? null : releaseLink(answer?.url);
  const linkLabel = answer?.state === "newer" && typeof answer.latest === "string" ? `Open the page of OpenStage Guide ${answer.latest}` : "Open the releases page";
  return { words, link, linkLabel };
}

/**
 * Asks, and shows the answer.
 * @param {Document} doc
 */
export async function checkForUpdates(doc) {
  /** @type {any} */
  let answer;
  try {
    answer = await post("updates", {});
  } catch (error) {
    answer = { state: "error", message: /** @type {Error} */ (error).message };
  }
  const { words, link, linkLabel } = describeAnswer(answer);
  await openDialog(doc, {
    title: "Check for updates",
    body: () => [
      el(doc, "p", { text: words, role: "status" }),
      // A new window is never opened for the link: the window opens the page in the person's own browser, and a browser opens a tab.
      ...(link ? [el(doc, "p", {}, el(doc, "a", { href: link, target: "_blank", rel: "noopener noreferrer", text: linkLabel }))] : []),
      el(doc, "p", { class: "hint", text: "OpenStage Guide asks GitHub only when you choose this, sends nothing about you or your guides, and installs nothing." }),
    ],
    confirm: "Close",
    cancel: "",
    finish: () => true,
  });
}
