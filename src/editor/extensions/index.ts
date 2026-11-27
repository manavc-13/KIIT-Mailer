// Extension set for the email editor. Also used headlessly in tests.
import type { AnyExtension } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { TextStyle, FontFamily, FontSize, Color, BackgroundColor } from '@tiptap/extension-text-style';
import TextAlign from '@tiptap/extension-text-align';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import Placeholder from '@tiptap/extension-placeholder';
import { BlockStyle } from './block-style';
import { Field } from './field';
import { KbSpan } from './kb-span';
import { tableExtensions } from './table';
import { EmailImage } from './image';
import { blockExtensions } from './blocks';
import { ImageUploads, type ImageUploadsOptions } from './image-uploads';

export interface BuildOptions {
  variant?: 'body' | 'signature';
  placeholder?: string;
  uploads?: Partial<ImageUploadsOptions>;
  /** Headless/tests: skip React node views and the upload plugin. */
  headless?: boolean;
}

export function buildExtensions(opts: BuildOptions = {}): AnyExtension[] {
  const signature = opts.variant === 'signature';
  const exts: AnyExtension[] = [
    StarterKit.configure({
      heading: { levels: [1, 2, 3] },
      code: false,
      codeBlock: false,
      link: {
        openOnClick: false,
        autolink: true,
        defaultProtocol: 'https',
        protocols: ['mailto', 'tel'],
        HTMLAttributes: { target: '_blank', rel: 'noopener noreferrer', class: null },
        isAllowedUri: (url, ctx) => /^\{[^{}]+\}/.test(url) || ctx.defaultValidate(url),
      },
    }),
    TextStyle,
    FontFamily,
    FontSize,
    Color,
    BackgroundColor,
    TextAlign.configure({ types: ['heading', 'paragraph'], alignments: ['left', 'center', 'right', 'justify'] }),
    Subscript,
    Superscript,
    BlockStyle,
    Field,
    KbSpan,
    EmailImage,
    ...tableExtensions,
  ];
  // Block nodes are always in the schema (so pasted/loaded blocks survive); the signature
  // toolbar simply doesn't offer them.
  exts.push(...blockExtensions);
  if (!opts.headless) {
    exts.push(Placeholder.configure({ placeholder: opts.placeholder ?? (signature ? 'Your name, title and contact details…' : 'Write your message…') }));
    exts.push(ImageUploads.configure(opts.uploads || {}));
  }
  return exts;
}
