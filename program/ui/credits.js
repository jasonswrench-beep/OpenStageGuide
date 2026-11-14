// @ts-check
/**
 * What the About window says about the program's version, its license and what it is made with. The NOTICE file at the top of the program's folder
 * has the full text; a test keeps what is said here in step with it.
 */
export const LICENSE_LINE = "OpenStage Guide is free software under the Apache License 2.0. Copyright 2026 Jason S. Wrench and the OpenStage Guide contributors.";

/** What the program is made with, each with the license it is used under. */
export const MADE_WITH = Object.freeze([
  { name: "JSZip", use: "makes the Word files", license: "MIT License" },
  { name: "Lucide icons", use: "the icons on the menus and buttons", license: "ISC License" },
  { name: "Atkinson Hyperlegible, Lexend, Source Sans 3 and Source Serif 4", use: "the fonts that come with the program", license: "SIL Open Font License 1.1" },
]);

export const ICONS_LINE = "The training icons are by Jason S. Wrench and are licensed under Creative Commons Attribution-NonCommercial-ShareAlike 4.0, not under the Apache License.";
export const NOTICE_LINE = "The full notices are in the NOTICE file in the program's folder.";

/**
 * The line saying which version this is, from what the server reports, or null when it did not say (a made-up version is never shown).
 * @param {{ version?: unknown, build?: unknown } | null | undefined} info
 */
export function aboutVersion(info) {
  const version = String(info?.version ?? "").trim();
  if (!/^\d+\.\d+\.\d+$/.test(version)) return null;
  const build = String(info?.build ?? "").trim();
  return `Version ${version}${/^[0-9a-f]{6,64}$/.test(build) ? `, build ${build}` : ""}.`;
}
