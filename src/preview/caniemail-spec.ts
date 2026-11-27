// Which caniemail features and platforms the preview cares about. Shared by
// scripts/build-caniemail.ts (to slice the vendored data) and the runtime lookup.
// No imports: this file is loaded by the Node build script as well as the browser.

/** Compact platform key -> [caniemail family, caniemail platform]. */
export const CANIEMAIL_PLATFORMS = {
  'gmail-desktop': ['gmail', 'desktop-webmail'],
  'gmail-ios': ['gmail', 'ios'],
  'gmail-android': ['gmail', 'android'],
  'outlook-windows': ['outlook', 'windows'],
  'apple-mail-macos': ['apple-mail', 'macos'],
  'apple-mail-ios': ['apple-mail', 'ios'],
} as const satisfies Record<string, readonly [string, string]>;

export type PlatformKey = keyof typeof CANIEMAIL_PLATFORMS;

/** The four clients the preview targets, and which platforms each one stands for. */
export const TARGET_CLIENTS = [
  { id: 'gmail-desktop', label: 'Gmail (desktop)', short: 'Gmail web', platforms: ['gmail-desktop'] },
  { id: 'gmail-mobile', label: 'Gmail (mobile apps)', short: 'Gmail app', platforms: ['gmail-ios', 'gmail-android'] },
  { id: 'apple-mail', label: 'Apple Mail (iPhone, Mac)', short: 'Apple Mail', platforms: ['apple-mail-ios', 'apple-mail-macos'] },
  { id: 'outlook-windows', label: 'Outlook (Windows desktop)', short: 'Outlook Win', platforms: ['outlook-windows'] },
] as const satisfies readonly { id: string; label: string; short: string; platforms: readonly PlatformKey[] }[];

export type TargetClientId = (typeof TARGET_CLIENTS)[number]['id'];

/** Every caniemail feature slug the detector in features.ts can report. */
export const FEATURE_SLUGS = [
  'html-style',
  'css-at-media',
  'css-at-media-prefers-color-scheme',
  'css-at-font-face',
  'css-at-import',
  'css-variables',
  'css-border-radius',
  'css-box-shadow',
  'css-text-shadow',
  'css-background-image',
  'css-background-size',
  'css-linear-gradient',
  'css-radial-gradient',
  'css-max-width',
  'css-min-width',
  'css-max-height',
  'css-display-flex',
  'css-display-grid',
  'css-display-none',
  'css-gap',
  'css-position',
  'css-z-index',
  'css-float',
  'css-margin',
  'css-padding',
  'css-opacity',
  'css-overflow',
  'css-object-fit',
  'css-transform',
  'css-transition',
  'css-animation',
  'css-at-keyframes',
  'css-pseudo-class-hover',
  'css-unit-rem',
  'css-unit-vw',
  'css-unit-calc',
  'css-function-clamp',
  'css-sytem-ui',
  'html-background',
  'html-svg',
  'image-svg',
  'html-video',
  'html-audio',
  'html-picture',
  'html-srcset',
  'html-form',
  'html-semantics',
  'image-webp',
  'image-avif',
  'image-gif',
  'image-base64',
] as const;

export type FeatureSlug = (typeof FEATURE_SLUGS)[number];
