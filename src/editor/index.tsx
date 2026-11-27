// Public entry for the editor area.
//
// <EmailEditor>    TipTap rich-text editor for email bodies (and sender signatures).
// <HtmlCodeEditor> CodeMirror HTML source editor (html mode).
// serializeEmail() Compile saved editor JSON to the email-safe body fragment.
// prepareHtmlForEditor() / normalizePastedHtml() Clean legacy / clipboard HTML.
export { EmailEditor } from './EmailEditor';
export type { EmailEditorProps, EmailEditorValue } from './EmailEditor';
export { HtmlCodeEditor, formatHtml } from './HtmlCodeEditor';
export type { HtmlCodeEditorProps } from './HtmlCodeEditor';
export { serializeEmail } from './serialize';
export { normalizePastedHtml, prepareHtmlForEditor, plainTextToHtml } from './paste';
