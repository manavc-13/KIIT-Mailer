// Markup contract for custom content blocks. The editor serializer (src/editor) emits
// exactly this markup; shared/themes themeCss() styles these classes; the server inlines
// the CSS with juice. Keep all three in sync.
//
// Button (bulletproof, works in Outlook):
//   <table role="presentation" class="kb-button" data-variant="primary|secondary" align="left|center"
//          cellpadding="0" cellspacing="0" border="0"><tr><td class="kb-button-cell">
//     <a class="kb-button-link" href="https://..." target="_blank">Label</a>
//   </td></tr></table>
//
// Callout (highlighted note; variant info|warn|success):
//   <table role="presentation" class="kb-callout" data-variant="info" width="100%" cellpadding="0" cellspacing="0" border="0">
//     <tr><td class="kb-callout-cell"> ...paragraphs... </td></tr></table>
//
// Details panel (label/value rows, e.g. Date / Venue / Time):
//   <table role="presentation" class="kb-details" width="100%" cellpadding="0" cellspacing="0" border="0">
//     <tr><td class="kb-details-label">Date</td><td class="kb-details-value">12 Dec 2026</td></tr> ...</table>
//
// Divider:
//   <table role="presentation" class="kb-divider" width="100%" cellpadding="0" cellspacing="0" border="0">
//     <tr><td class="kb-divider-cell">&nbsp;</td></tr></table>
//
// Images: <img src="https://..." alt="..." width="NNN" style="display:block;max-width:100%;height:auto;border:0;">
// User formatting (font family / size / colour / highlight / alignment / line-height) is always
// emitted as INLINE styles on spans/paragraphs, which take precedence over theme CSS.
export const BLOCK_CLASSES = {
  body: 'kb-body',
  button: 'kb-button', buttonCell: 'kb-button-cell', buttonLink: 'kb-button-link',
  callout: 'kb-callout', calloutCell: 'kb-callout-cell',
  details: 'kb-details', detailsLabel: 'kb-details-label', detailsValue: 'kb-details-value',
  divider: 'kb-divider', dividerCell: 'kb-divider-cell',
} as const;
