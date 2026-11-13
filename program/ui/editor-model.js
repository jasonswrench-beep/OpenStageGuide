// @ts-check
/**
 * The editor's decisions, apart from the page: which rows of the drawn guide stay and which are drawn again, what each key does, and the
 * words that name a paragraph's text box. Pure, so they are tested without a page (tests/program-editor-model.test.mjs).
 */

/** A row's signature: it is drawn again only when this changes. @param {{ style?: string, xml?: string }} entry */
export function rowSignature(entry) {
  return `${entry.style ?? ""}\u0001${entry.xml ?? ""}`;
}

/** @param {string[]} a @param {string[]} b */
export function sameList(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/**
 * Which rows stay as they are and which are replaced: the same at the start and at the end, different in the middle. Rows [0, prefix) and the
 * rows after the replaced ones are kept; old rows [prefix, oldEnd) are replaced by new rows [prefix, newEnd).
 * @param {string[]} before @param {string[]} after
 */
export function reconcile(before, after) {
  const most = Math.min(before.length, after.length);
  let prefix = 0;
  while (prefix < most && before[prefix] === after[prefix]) prefix++;
  let suffix = 0;
  while (suffix < most - prefix && before[before.length - 1 - suffix] === after[after.length - 1 - suffix]) suffix++;
  return { prefix, oldEnd: before.length - suffix, newEnd: after.length - suffix };
}

/**
 * @typedef {{ key: string, shift?: boolean, ctrl?: boolean, alt?: boolean, meta?: boolean }} KeyInfo
 * @typedef {KeyInfo & { start: number, end: number, length: number, composing?: boolean }} WordsKeyInfo
 */

/**
 * What a key does in the words of a paragraph being edited. `caret` is where the caret goes in the paragraph the key leads to.
 * @param {WordsKeyInfo} k
 * @returns {{ do: "split" | "join" | "previous" | "next" | "stop" | "undo" | "redo", caret?: "start" | "end" } | null}
 */
export function wordsKey(k) {
  if (k.composing) return null;
  const plain = !k.shift && !k.ctrl && !k.alt && !k.meta;
  const command = (k.ctrl || k.meta) && !k.alt;
  const letter = k.key.length === 1 ? k.key.toLowerCase() : "";
  if (command && letter === "z") return { do: k.shift ? "redo" : "undo" };
  if (command && letter === "y" && !k.shift) return { do: "redo" };
  if (!plain) return null;
  const atStart = k.start === 0 && k.end === 0;
  const atEnd = k.start === k.length && k.end === k.length;
  switch (k.key) {
    case "Enter":
      return { do: "split" };
    case "Escape":
      return { do: "stop" };
    case "Backspace":
      return atStart ? { do: "join" } : null;
    case "ArrowUp":
    case "ArrowLeft":
      return atStart ? { do: "previous", caret: "end" } : null;
    case "ArrowDown":
    case "ArrowRight":
      return atEnd ? { do: "next", caret: "start" } : null;
    default:
      return null;
  }
}

/**
 * What a key does when a row of the guide has the keyboard and its words are not being edited.
 * @param {KeyInfo} k
 * @returns {"down" | "up" | "first" | "last" | "edit" | "undo" | "redo" | "extendDown" | "extendUp" | "moveDown" | "moveUp" | "delete" | "copy" | "cut" | "add" | "selectAll" | "clear" | null}
 */
export function rowKey(k) {
  const command = (k.ctrl || k.meta) && !k.alt;
  const letter = k.key.length === 1 ? k.key.toLowerCase() : "";
  if (command && letter === "z") return k.shift ? "redo" : "undo";
  if (command && letter === "y" && !k.shift) return "redo";
  if (command && !k.shift && letter === "c") return "copy";
  if (command && !k.shift && letter === "x") return "cut";
  if (command && !k.shift && letter === "a") return "selectAll";
  if (command && !k.shift && k.key === "Enter") return "add";
  if (k.alt && !k.ctrl && !k.meta && !k.shift) {
    if (k.key === "ArrowUp") return "moveUp";
    if (k.key === "ArrowDown") return "moveDown";
    return null;
  }
  if (k.shift && !k.ctrl && !k.alt && !k.meta) {
    if (k.key === "ArrowDown") return "extendDown";
    if (k.key === "ArrowUp") return "extendUp";
    return null;
  }
  if (k.shift || k.ctrl || k.alt || k.meta) return null;
  switch (k.key) {
    case "Delete":
      return "delete";
    case "Escape":
      return "clear";
    case "ArrowDown":
      return "down";
    case "ArrowUp":
      return "up";
    case "Home":
      return "first";
    case "End":
      return "last";
    case "Enter":
    case "F2":
      return "edit";
    default:
      return null;
  }
}

/** The name of the text box for a paragraph's words, for a screen reader. @param {string} style */
export function wordsLabel(style) {
  /** @type {Record<string, string>} */
  const names = {
    "OSG Guide Title": "Guide title",
    Title: "Guide title",
    "OSG Module": "Module title",
    "OSG Lesson": "Lesson title",
    "OSG Assessment": "Assessment title",
    "OSG Cue": "Block words",
    "OSG Cue Facilitator": "Block words, for the facilitator",
    "OSG Cue Body": "Block text",
    "OSG Facilitator Notes": "Facilitator notes",
    "heading 1": "Heading",
    "heading 2": "Heading",
    "heading 3": "Heading",
  };
  return names[style] ?? "Paragraph words";
}
