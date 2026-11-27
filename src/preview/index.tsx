// Public entry for the email preview.
//
// <EmailPreview> calls POST /api/render/preview (debounced) and shows the result in
// client/device tabs (Gmail desktop, Gmail mobile, Apple Mail, Outlook approximation,
// dark-mode simulation, plain text) plus the lint/compatibility panel.
export { EmailPreview, type EmailPreviewProps } from './EmailPreview';
export { renderRequestKey } from './useRenderPreview';
export { ChecksPanel } from './ChecksPanel';
export { compatibilityReport, supportFor, type CompatRow } from './caniemail';
export { detectFeatures } from './features';
export { htmlForClient, outlookTransform, gmailTransform, type PreviewClientId } from './transforms';
