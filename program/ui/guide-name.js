// @ts-check
/**
 * What a guide is called on screen. A guide is kept in a file whose name ends in `.guide-project.json`; the person who made it named the guide,
 * not the ending, so lists and dialogs show only the name. The file's whole name is still what the program asks the server for.
 * @param {string} fileName
 */
export function guideName(fileName) {
  const name = String(fileName ?? "");
  const shown = name.replace(/\.guide-project\.json$/i, "");
  return shown || name;
}
