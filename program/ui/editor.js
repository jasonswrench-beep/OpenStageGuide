// @ts-check
/**
 * The guide as the author writes it. The guide is drawn as rows, one for each paragraph (a table is one row). The arrow keys move between
 * rows; Enter, F2 or a click starts changing a row's words, which become a text box that grows with them; Escape stops. Enter in the words
 * starts the next paragraph; Backspace at the start joins a paragraph to the one before; Up and Down (and Left and Right) at the ends of the
 * words go to the next or the last paragraph. Every change is made on the guide's document (guide-word.mjs, with its saving) through
 * src/guide-edit.mjs, which changes the words and nothing else in a paragraph, so what the editor cannot show as words (a picture, a check
 * box, a citation, the time line, a table) is never rewritten: those rows are read, and changed with the tools beside the guide.
 *
 * The design is in docs/EDITING-SURFACE-DESIGN-2026-11-13.md. The page writes text and attributes only, never markup.
 */
import { KINDS, isHeading, joinWords, kindChange, nextParagraphSpec, paragraphWith, readParagraph, setWords, splitWords, withStyle } from "../../src/guide-edit.mjs";
import { reconcile, rowKey, rowSignature, sameList, wordsKey, wordsLabel } from "./editor-model.js";
import { describeGuide, drawNodes } from "./preview.js";

/** @param {string} xml */
const parseXml = (xml) => new DOMParser().parseFromString(xml, "application/xml");
const io = { parse: parseXml, serialize: (/** @type {Element} */ el) => new XMLSerializer().serializeToString(el) };
/** How long after the author stops moving or typing the tools are told where the cursor is. */
const SELECTION_DELAY = 250;

/**
 * @typedef {{ xml?: string, text?: string, style?: string }} Entry
 * @typedef {{ sig: string, el: HTMLElement }} Row
 * @typedef {{ row: Row, ta: HTMLTextAreaElement, group: string }} Editing
 */

/**
 * @param {Document} doc
 * @param {HTMLElement} host where the guide is drawn
 * @param {any} guide the guide's document, with its saving
 * @param {{ say: (message: string, bad?: boolean) => void, choose: (options: { title: string, intro?: string, confirm: string, options: { id: string, label: string, note?: string }[], current?: string }) => Promise<string | null> }} ui
 */
export function mountEditor(doc, host, guide, ui) {
  const page = doc.createElement("article");
  page.className = "page";
  page.setAttribute("aria-label", "The guide");
  const hint = doc.createElement("p");
  hint.id = "edit-hint";
  hint.className = "sr-only";
  hint.textContent = "Enter starts a new paragraph. Escape stops changing the words. Up at the start or Down at the end goes to the next paragraph. Control Z undoes.";
  const empty = doc.createElement("p");
  empty.className = "g-empty";
  empty.textContent = "This guide is empty. Choose Guide overview in the tools beside it to start it.";
  host.replaceChildren(page, hint);

  /** @type {Entry[]} */
  let entries = [];
  /** @type {string[]} */
  let entrySigs = [];
  /** @type {Row[]} */
  let shown = [];
  /** @type {Editing | null} */
  let editing = null;
  let groupCounter = 0;
  let queue = Promise.resolve();
  /** @type {number | null} */
  let selectionTimer = null;
  let gone = false;
  /** The rows chosen as a range (by the index of their paragraph), as with Shift+Down. @type {{ anchor: number, head: number } | null} */
  let selection = null;
  /** The row the author was last on, so a command from a menu knows which one it is for. @type {Row | null} */
  let lastRow = null;
  /** What was last cut or copied inside the guide: whole paragraphs, to be pasted back as they were. @type {{ paragraphs: { xml: string, style: string }[], text: string } | null} */
  let clip = null;

  /** @param {Row} row */
  const idx = (row) => Number(row.el.dataset.index);
  /** @param {number} index */
  const rowAt = (index) => shown.find((r) => idx(r) === index) ?? null;
  /** The row for a paragraph, or the nearest one above it that is drawn (some paragraphs, such as the contents lines, have no row). @param {number} index */
  const rowNear = (index) => rowAt(index) ?? [...shown].reverse().find((r) => idx(r) <= index) ?? shown[0] ?? null;
  /** @param {Row} row */
  const entryOf = (row) => entries[idx(row)] ?? { xml: "", text: "", style: "Normal" };
  /** @param {unknown} e */
  const messageOf = (e) => (e instanceof Error ? e.message : String(e));

  function tellTools() {
    if (selectionTimer !== null) window.clearTimeout(selectionTimer);
    selectionTimer = window.setTimeout(() => {
      selectionTimer = null;
      try {
        guide.fireSelection();
      } catch {
        /* the tools may not be there yet */
      }
    }, SELECTION_DELAY);
  }

  /** @param {Row} row */
  function placeCursor(row) {
    try {
      guide.place(idx(row));
    } catch {
      /* the document may be closing */
    }
    tellTools();
  }

  /** @param {Row | null} row the row that gets the Tab stop */
  function setTabStop(row) {
    for (const r of shown) r.el.tabIndex = r === row ? 0 : -1;
  }

  /** @param {Row} row @param {Entry} entry */
  function decorate(row, entry) {
    const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
    row.el.dataset.editable = info.editable ? "true" : "false";
    row.el.setAttribute("aria-description", info.editable ? "Press Enter to change the words." : info.reason ?? "This part is changed with the tools beside the guide.");
  }

  /** Read the document and bring the rows up to date: only what changed is drawn again. */
  function render() {
    if (gone) return;
    entries = guide.exportState().paragraphs;
    const sigs = entries.map(rowSignature);
    if (sameList(sigs, entrySigs)) return;
    entrySigs = sigs;
    const nodes = describeGuide(entries, parseXml, { blank: true });
    const rowSigs = nodes.map((n) => rowSignature(entries[n.index ?? 0]));
    const { prefix, oldEnd, newEnd } = reconcile(shown.map((r) => r.sig), rowSigs);
    // A row that is being replaced cannot go on being edited: the document changed under it (Undo, or the tools).
    if (editing && shown.indexOf(editing.row) >= prefix && shown.indexOf(editing.row) < oldEnd) editing = null;
    const hadFocus = doc.activeElement && page.contains(doc.activeElement);
    for (const r of shown.slice(prefix, oldEnd)) r.el.remove();
    if (empty.parentNode) empty.remove();
    const before = shown[oldEnd]?.el ?? null;
    /** @type {Row[]} */
    const made = [];
    for (let i = prefix; i < newEnd; i++) {
      const node = nodes[i];
      const tmp = doc.createElement("div");
      drawNodes(doc, tmp, [node]);
      const el = /** @type {HTMLElement} */ (tmp.firstElementChild);
      const row = { sig: rowSigs[i], el };
      el.tabIndex = -1;
      el.dataset.index = String(node.index);
      el.dataset.span = String(node.span ?? 1);
      decorate(row, entries[node.index ?? 0]);
      page.insertBefore(el, before);
      made.push(row);
    }
    shown = [...shown.slice(0, prefix), ...made, ...shown.slice(oldEnd)];
    shown.forEach((r, i) => {
      r.el.dataset.index = String(nodes[i].index);
    });
    if (!shown.length) page.append(empty);
    if (!shown.some((r) => r.el.tabIndex === 0)) setTabStop(shown[0] ?? null);
    markSelection();
    // The keyboard stays where it was, even when the row it was on has been drawn again.
    if (hadFocus && !doc.activeElement?.closest("#work-guide") && !editing) focusRow(rowNear(guide.cursor()));
  }

  /** @param {Row | null} row */
  function focusRow(row) {
    if (!row) return;
    setTabStop(row);
    lastRow = row;
    row.el.focus();
    placeCursor(row);
  }

  /** Draw a row's own words again, after they were being changed. @param {Row} row */
  function redrawRow(row) {
    const entry = entryOf(row);
    const [node] = describeGuide([entry], parseXml, { blank: true });
    row.el.classList.remove("g-editing");
    if (!node) {
      row.el.replaceChildren();
      return;
    }
    const tmp = doc.createElement("div");
    drawNodes(doc, tmp, [node]);
    const fresh = tmp.firstElementChild;
    row.el.replaceChildren(...(fresh ? Array.from(fresh.childNodes) : []));
  }

  /** Stop changing words (they are already in the document), and draw the row as it reads. */
  function stopEdit() {
    if (!editing) return;
    const e = editing;
    editing = null;
    redrawRow(e.row);
  }

  /**
   * Change the words of a row.
   * @param {Row} row @param {number | "start" | "end"} caret
   */
  function startEdit(row, caret) {
    const entry = entryOf(row);
    const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!info.editable) {
      stopEdit();
      focusRow(row);
      ui.say(info.reason ?? "This part is changed with the tools beside the guide.");
      return;
    }
    if (editing?.row === row) {
      editing.ta.focus();
      return;
    }
    stopEdit();
    const ta = doc.createElement("textarea");
    ta.className = "g-edit";
    ta.rows = 1;
    ta.value = info.text;
    ta.spellcheck = true;
    ta.setAttribute("aria-label", wordsLabel(entry.style ?? "Normal"));
    ta.setAttribute("aria-describedby", "edit-hint");
    /** @type {HTMLElement[]} */
    const lead = [];
    if (info.prefix) {
      const s = doc.createElement("span");
      s.className = "g-lead";
      s.textContent = info.prefix.trim();
      lead.push(s);
    }
    for (const part of info.lead) {
      if (part.kind !== "label" || !part.text?.trim()) continue;
      const s = doc.createElement("span");
      s.className = "g-lead g-label";
      s.textContent = part.text.trim();
      lead.push(s);
    }
    row.el.replaceChildren(...lead, ta);
    row.el.classList.add("g-editing");
    setTabStop(row);
    editing = { row, ta, group: `type:${++groupCounter}` };
    lastRow = row;
    clearSelection();
    ta.addEventListener("input", () => onInput(editing && editing.ta === ta ? editing : null));
    ta.addEventListener("keydown", (e) => onWordsKey(e, row, ta));
    ta.addEventListener("paste", (e) => onWordsPaste(e, row, ta));
    ta.addEventListener("focusout", () => {
      if (editing && editing.ta === ta) stopEdit();
    });
    ta.focus();
    const at = caret === "start" ? 0 : caret === "end" ? ta.value.length : Math.max(0, Math.min(caret, ta.value.length));
    ta.setSelectionRange(at, at);
    grow(ta);
    placeCursor(row);
  }

  /** @param {HTMLTextAreaElement} ta */
  function grow(ta) {
    ta.style.height = "auto";
    ta.style.height = `${ta.scrollHeight}px`;
  }

  /** The words in the box are the words in the document: a change goes through the queue so none is applied out of order. @param {Editing | null} e */
  function onInput(e) {
    if (!e) return;
    // A paragraph is one line of words: line ends that were pasted in become spaces.
    if (/[\r\n]/.test(e.ta.value)) {
      const at = e.ta.selectionStart;
      e.ta.value = e.ta.value.replace(/\r\n|[\r\n]/g, " ");
      e.ta.setSelectionRange(at, at);
    }
    grow(e.ta);
    run(() => commit(e));
    tellTools();
  }

  /** @param {() => Promise<void>} job */
  function run(job) {
    queue = queue.then(job).catch((err) => {
      ui.say(messageOf(err), true);
      render();
    });
    return queue;
  }

  /** Put the words that are in the box into the document, as part of the run of typing it is in. @param {Editing} e */
  async function commit(e) {
    if (gone || !shown.includes(e.row)) return;
    const entry = entryOf(e.row);
    const xml = setWords(entry.xml ?? "", entry.style ?? "Normal", e.ta.value, io);
    if (xml === entry.xml) return;
    const at = idx(e.row);
    await guide.replaceParagraphs({ start: at, count: 1, guard: entry.text ?? "", xmls: [xml], select: at, group: e.group });
    entries = guide.exportState().paragraphs;
    entrySigs = entries.map(rowSignature);
    e.row.sig = rowSignature(entries[at]);
  }

  /**
   * @param {KeyboardEvent} ev @param {Row} row @param {HTMLTextAreaElement} ta
   */
  function onWordsKey(ev, row, ta) {
    const action = wordsKey({
      key: ev.key,
      shift: ev.shiftKey,
      ctrl: ev.ctrlKey,
      alt: ev.altKey,
      meta: ev.metaKey,
      start: ta.selectionStart,
      end: ta.selectionEnd,
      length: ta.value.length,
      composing: ev.isComposing,
    });
    if (!action) return;
    ev.preventDefault();
    const e = editing;
    if (!e || e.ta !== ta) return;
    switch (action.do) {
      case "stop":
        run(async () => {
          await commit(e);
          stopEdit();
          focusRow(row);
        });
        break;
      case "split":
        run(() => split(e));
        break;
      case "join":
        run(() => join(e));
        break;
      case "previous":
      case "next": {
        const caret = action.caret ?? "start";
        const step = action.do === "next" ? 1 : -1;
        run(async () => {
          await commit(e);
          move(row, step, caret, true);
        });
        break;
      }
      case "undo":
      case "redo":
        run(() => history(action.do === "redo", true));
        break;
    }
  }

  /**
   * Go to the row before or after: into its words when it has some the author can change and the author was in words, else onto the row.
   * @param {Row} from @param {number} step @param {"start" | "end"} caret @param {boolean} wasEditing
   */
  function move(from, step, caret, wasEditing) {
    const to = shown[shown.indexOf(from) + step];
    if (!to) return;
    if (wasEditing && to.el.dataset.editable === "true") startEdit(to, caret);
    else {
      stopEdit();
      focusRow(to);
    }
  }

  /** Enter: the words after the caret become the next paragraph (a heading just goes on to the next row). @param {Editing} e */
  async function split(e) {
    const row = e.row;
    const entry = entryOf(row);
    if (isHeading(entry.style ?? "")) {
      await commit(e);
      move(row, 1, "start", true);
      return;
    }
    const start = e.ta.selectionStart;
    const end = e.ta.selectionEnd;
    // What is selected is replaced by the Enter, as in any editor.
    if (end > start) {
      e.ta.value = e.ta.value.slice(0, start) + e.ta.value.slice(end);
      e.ta.setSelectionRange(start, start);
    }
    await commit(e);
    const now = entryOf(row);
    const cut = splitWords(now.xml ?? "", now.style ?? "Normal", start, io);
    if (!cut) {
      ui.say("This part cannot be split.");
      return;
    }
    const at = idx(row);
    // The row it was is drawn as it reads (a row whose paragraph did not change is kept, so it must not be left as a text box).
    stopEdit();
    await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls: [cut.first, cut.second], select: at + 1 });
    render();
    const next = rowAt(at + 1);
    if (next) startEdit(next, 0);
  }

  /** Backspace at the start: the paragraph joins the one before; an empty one that cannot join is removed. @param {Editing} e */
  async function join(e) {
    const row = e.row;
    await commit(e);
    const at = idx(row);
    const current = entries[at];
    const previous = entries[at - 1];
    if (!previous) return;
    const joined = joinWords(previous, current, io);
    stopEdit();
    if (joined) {
      await guide.replaceParagraphs({ start: at - 1, count: 2, guard: previous.text ?? "", xmls: [joined.xml], select: at - 1 });
      render();
      const target = rowAt(at - 1);
      if (target) startEdit(target, joined.offset);
      return;
    }
    const info = readParagraph(current.xml ?? "", current.style ?? "Normal", io);
    if (info.editable && info.text === "" && !info.lead.length) {
      // Nothing is lost by taking an empty paragraph away.
      await guide.replaceParagraphs({ start: at, count: 1, guard: current.text ?? "", xmls: [], select: at - 1 });
      render();
      const target = rowAt(at - 1);
      if (target) {
        if (target.el.dataset.editable === "true") startEdit(target, "end");
        else focusRow(target);
      }
      return;
    }
    // It stays as it is: the words go back into the box.
    render();
    const again = rowAt(at);
    if (again) startEdit(again, 0);
  }

  /** @param {boolean} redo @param {boolean} wasEditing */
  async function history(redo, wasEditing) {
    stopEdit();
    // The paragraphs that were chosen may not be the same paragraphs after this: nothing stays chosen.
    selection = null;
    const done = redo ? guide.redo() : guide.undo();
    if (!done) {
      ui.say(redo ? "There is nothing to redo." : "There is nothing to undo.");
      if (wasEditing) {
        const r = rowNear(guide.cursor());
        if (r) startEdit(r, "end");
      }
      return;
    }
    render();
    const r = rowNear(guide.cursor());
    if (!r) return;
    if (wasEditing && r.el.dataset.editable === "true") startEdit(r, "end");
    else focusRow(r);
    ui.say(redo ? "Redone." : "Undone.");
  }

  // ---- rows chosen together, and what is done to them

  /** @param {Row} row How many entries of the document the row covers: a table covers its cells. */
  const spanOf = (row) => Number(row.el.dataset.span || 1);

  function selectedRows() {
    if (!selection) return [];
    const lo = Math.min(selection.anchor, selection.head);
    const hi = Math.max(selection.anchor, selection.head);
    return shown.filter((r) => idx(r) >= lo && idx(r) <= hi);
  }

  function markSelection() {
    const chosen = new Set(selectedRows());
    if (selection && !chosen.size) selection = null;
    for (const r of shown) {
      if (chosen.has(r)) r.el.dataset.selected = "true";
      else delete r.el.dataset.selected;
    }
  }

  function clearSelection() {
    if (!selection) return;
    selection = null;
    markSelection();
  }

  /** @param {number} anchor @param {number} head */
  function select(anchor, head) {
    selection = { anchor, head };
    markSelection();
    const n = selectedRows().length;
    ui.say(`${n} paragraph${n === 1 ? "" : "s"} selected.`);
  }

  /** Shift+Down and Shift+Up: the range grows or shrinks by a row, and the keyboard goes with it. @param {Row} row @param {number} step */
  function extend(row, step) {
    const to = shown[shown.indexOf(row) + step];
    if (!to) return;
    select(selection ? selection.anchor : idx(row), idx(to));
    focusRow(to);
  }

  /** The rows a command is for: the ones chosen together, or the one the author is on. */
  function targetRows() {
    const chosen = selectedRows();
    if (chosen.length) return chosen;
    const one = editing?.row ?? (lastRow && shown.includes(lastRow) ? lastRow : null);
    return one ? [one] : [];
  }

  async function settle() {
    if (editing) await commit(editing);
  }

  /** @param {string} name */
  async function command(name) {
    switch (name) {
      case "delete":
        return remove();
      case "moveUp":
        return move1(-1);
      case "moveDown":
        return move1(1);
      case "add":
        return addBelow();
      case "kind":
        return changeKind();
      case "copy":
        return copy(false);
      case "cut":
        return copy(true);
      case "paste":
        return paste(null);
      case "undo":
        return history(false, false);
      case "redo":
        return history(true, false);
    }
  }

  async function remove() {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const start = idx(rows[0]);
    const last = rows[rows.length - 1];
    const end = idx(last) + spanOf(last) - 1;
    stopEdit();
    await guide.deleteRange({ start, end, guard: entries[start]?.text ?? "" });
    selection = null;
    render();
    focusRow(rowNear(Math.max(0, start - 1)) ?? shown[0] ?? null);
    ui.say(rows.length === 1 ? "Paragraph deleted. Ctrl+Z brings it back." : `${rows.length} paragraphs deleted. Ctrl+Z brings them back.`);
  }

  /** @param {number} step */
  async function move1(step) {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const first = rows[0];
    const last = rows[rows.length - 1];
    const start = idx(first);
    const end = idx(last) + spanOf(last) - 1;
    let after;
    if (step < 0) {
      const above = shown[shown.indexOf(first) - 1];
      if (!above) return ui.say("It is already at the top.");
      after = idx(above) - 1;
    } else {
      const below = shown[shown.indexOf(last) + 1];
      if (!below) return ui.say("It is already at the end.");
      after = idx(below) + spanOf(below) - 1;
    }
    stopEdit();
    const moved = await guide.moveRange({ start, end, guard: entries[start]?.text ?? "", after });
    render();
    select(moved.at, moved.at + (end - start));
    focusRow(rowNear(moved.at));
  }

  async function addBelow() {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const last = rows[rows.length - 1];
    const entry = entryOf(last);
    const at = idx(last) + spanOf(last);
    const spec = spanOf(last) === 1 ? nextParagraphSpec(entry.xml ?? "", entry.style ?? "Normal", io) : null;
    const xml = paragraphWith(spec ?? { styleId: null, numPr: null }, "", io);
    stopEdit();
    await guide.replaceParagraphs({ start: at, count: 0, guard: "", xmls: [xml], select: at });
    selection = null;
    render();
    const row = rowAt(at);
    if (row) startEdit(row, 0);
  }

  async function changeKind() {
    const rows = targetRows();
    if (rows.length !== 1) return ui.say("Choose one paragraph to change its kind.");
    const row = rows[0];
    await settle();
    const entry = entryOf(row);
    const can = kindChange(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!can.ok) return ui.say(can.reason ?? "This kind cannot be changed.");
    stopEdit();
    const chosen = await ui.choose({
      title: "Kind of paragraph",
      intro: "What this paragraph is. The tools and the participant workbook treat each kind in its own way.",
      confirm: "Change",
      options: KINDS.map((k) => ({ id: k.id, label: k.label, note: k.note })),
      current: KINDS.find((k) => k.style === entry.style)?.id,
    });
    const kind = KINDS.find((k) => k.id === chosen);
    if (!kind || kind.style === entry.style) {
      focusRow(row);
      return;
    }
    const at = idx(row);
    const now = entryOf(row);
    const xml = withStyle(now.xml ?? "", now.style ?? "Normal", kind.styleId, io);
    await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls: [xml], select: at });
    render();
    focusRow(rowAt(at));
    ui.say(`Now ${kind.label.toLowerCase()}.`);
  }

  /** @param {boolean} cut */
  async function copy(cut) {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    /** @type {{ xml: string, style: string }[]} */
    const paragraphs = [];
    for (const r of rows) {
      const e = entryOf(r);
      if (spanOf(r) !== 1 || !/^<w:p[\s>]/.test(e.xml ?? "")) return ui.say("Tables are moved and deleted, not copied, in this version.");
      paragraphs.push({ xml: e.xml ?? "", style: e.style ?? "Normal" });
    }
    const text = rows.map((r) => (entryOf(r).text ?? "").trim()).join("\n");
    clip = { paragraphs, text };
    // The system clipboard is a courtesy (it may refuse): the guide keeps its own copy.
    navigator.clipboard?.writeText(text).catch(() => {});
    if (cut) await remove();
    else ui.say(rows.length === 1 ? "Copied." : `${rows.length} paragraphs copied.`);
  }

  /** Paste after the rows the author is on. @param {string | null} text what the system clipboard held, or null from a menu */
  async function paste(text) {
    const rows = targetRows();
    const last = rows[rows.length - 1] ?? shown[shown.length - 1];
    if (!last) return;
    await settle();
    const plain = (text ?? "").replace(/\r\n?/g, "\n").trim();
    /** @type {string[]} */
    let xmls = [];
    if (clip && (text === null || plain === "" || plain === clip.text.trim())) xmls = clip.paragraphs.map((p) => p.xml);
    else if (plain) {
      const entry = entryOf(last);
      const spec = (spanOf(last) === 1 ? nextParagraphSpec(entry.xml ?? "", entry.style ?? "Normal", io) : null) ?? { styleId: null, numPr: null };
      xmls = plain
        .split("\n")
        .filter((l) => l.trim())
        .map((l) => paragraphWith(spec, l, io));
    }
    if (!xmls.length) return ui.say("There is nothing to paste. Copy or cut some paragraphs first.");
    const at = idx(last) + spanOf(last);
    stopEdit();
    await guide.replaceParagraphs({ start: at, count: 0, guard: "", xmls, select: at });
    render();
    select(at, at + xmls.length - 1);
    focusRow(rowNear(at));
  }

  /**
   * Several lines pasted into a paragraph's words: the first goes in at the caret, the others become paragraphs after it, the end of the
   * last one is the text that was after the caret. In a heading, which is one line, they are joined with spaces.
   * @param {ClipboardEvent} ev @param {Row} row @param {HTMLTextAreaElement} ta
   */
  function onWordsPaste(ev, row, ta) {
    const raw = ev.clipboardData?.getData("text/plain") ?? "";
    const lines = raw.split(/\r\n|[\r\n]/).filter((l) => l.trim());
    if (lines.length < 2) return;
    ev.preventDefault();
    const e = editing;
    if (!e || e.ta !== ta) return;
    const style = entryOf(row).style ?? "Normal";
    if (isHeading(style)) {
      ta.setRangeText(lines.join(" "), ta.selectionStart, ta.selectionEnd, "end");
      onInput(e);
      return;
    }
    run(async () => {
      await commit(e);
      const now = entryOf(row);
      const spec = nextParagraphSpec(now.xml ?? "", now.style ?? "Normal", io);
      if (!spec) return;
      const before = ta.value.slice(0, ta.selectionStart);
      const after = ta.value.slice(ta.selectionEnd);
      const xmls = [
        setWords(now.xml ?? "", now.style ?? "Normal", before + lines[0], io),
        ...lines.slice(1, -1).map((l) => paragraphWith(spec, l, io)),
        paragraphWith(spec, lines[lines.length - 1] + after, io),
      ];
      const at = idx(row);
      stopEdit();
      await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls, select: at + xmls.length - 1 });
      render();
      const last = rowAt(at + xmls.length - 1);
      if (last) startEdit(last, lines[lines.length - 1].length);
    });
  }

  /**
   * Put a block (a form, a table, lines to write on) after the row the author is on, choose what was put in, and say how to take it back.
   * @param {(ctx: any) => string} build @param {string} what
   */
  async function insertBlock(build, what) {
    const rows = targetRows();
    const last = rows[rows.length - 1] ?? shown[shown.length - 1] ?? null;
    await settle();
    stopEdit();
    const after = last ? idx(last) + spanOf(last) - 1 : -1;
    const before = entries.length;
    await guide.insert(build, last ? { where: "after", after, spare: false, select: "last" } : { where: "start", spare: false, select: "last" });
    selection = null;
    render();
    const added = entries.length - before;
    if (added > 0) {
      select(after + 1, after + added);
      focusRow(rowNear(after + 1));
    }
    ui.say(`${what} inserted. Ctrl+Z takes it back.`);
  }

  /** @param {EventTarget | null} target @returns {Row | null} */
  function rowOf(target) {
    const el = target instanceof Element ? target.closest("[data-index]") : null;
    return el ? shown.find((r) => r.el === el) ?? null : null;
  }

  // The keys of a row that has the keyboard, and a click on one.
  page.addEventListener("keydown", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    const row = rowOf(ev.target);
    if (!row || ev.target !== row.el) return;
    const action = rowKey({ key: ev.key, shift: ev.shiftKey, ctrl: ev.ctrlKey, alt: ev.altKey, meta: ev.metaKey });
    if (!action) return;
    ev.preventDefault();
    const at = shown.indexOf(row);
    if (action === "down") {
      clearSelection();
      focusRow(shown[at + 1] ?? row);
    } else if (action === "up") {
      clearSelection();
      focusRow(shown[at - 1] ?? row);
    } else if (action === "first") {
      clearSelection();
      focusRow(shown[0]);
    } else if (action === "last") {
      clearSelection();
      focusRow(shown[shown.length - 1]);
    } else if (action === "edit") startEdit(row, "end");
    else if (action === "extendDown" || action === "extendUp") extend(row, action === "extendDown" ? 1 : -1);
    else if (action === "clear") clearSelection();
    else if (action === "selectAll") {
      if (shown.length) select(idx(shown[0]), idx(shown[shown.length - 1]));
    } else if (action === "undo" || action === "redo") run(() => history(action === "redo", false));
    else run(() => command(action));
  });
  page.addEventListener("click", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    const row = rowOf(ev.target);
    if (!row) return;
    if (ev.shiftKey && lastRow && lastRow !== row && shown.includes(lastRow)) {
      select(selection ? selection.anchor : idx(lastRow), idx(row));
      focusRow(row);
      return;
    }
    clearSelection();
    startEdit(row, "end");
  });
  // Ctrl+V on a row: paragraphs cut or copied in the guide come back as they were, other text becomes a paragraph for each line.
  page.addEventListener("paste", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    ev.preventDefault();
    const text = ev.clipboardData?.getData("text/plain") ?? "";
    run(() => paste(text));
  });
  page.addEventListener("focusin", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    const row = rowOf(ev.target);
    if (row && ev.target === row.el) {
      // The keyboard going to a row that is not one of the chosen rows ends the choice (going along with the choice, as Shift+Down does, keeps it).
      if (selection && !selectedRows().includes(row)) clearSelection();
      setTabStop(row);
      lastRow = row;
      placeCursor(row);
    }
  });

  render();
  return {
    render,
    /** The document is about to go: stop everything. */
    destroy() {
      gone = true;
      if (selectionTimer !== null) window.clearTimeout(selectionTimer);
      editing = null;
    },
    /** A command of the Edit menu, for the rows the author is on or has chosen. @param {string} name */
    command(name) {
      return run(() => (name === "undo" || name === "redo" ? history(name === "redo", false) : command(name)));
    },
    /** Put a block after the row the author is on. @param {(ctx: any) => string} build @param {string} what */
    insert(build, what) {
      return run(() => insertBlock(build, what));
    },
    /** What the Edit menu needs to know to dim what cannot be done now. */
    state() {
      const h = guide.historyState();
      return { canUndo: h.undo > 0, canRedo: h.redo > 0, hasRows: targetRows().length > 0 };
    },
    /** Put the keyboard on the row of the cursor. */
    focusCursor() {
      focusRow(rowNear(guide.cursor()));
    },
  };
}
