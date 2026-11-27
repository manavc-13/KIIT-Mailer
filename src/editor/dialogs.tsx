// Dialogs used by the toolbar and node views: link, button block, image insert, alt text.
import { useEffect, useRef, useState } from 'react';
import { AlertTriangle, ImageUp, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { uploadFile } from '@/lib/uploads';
import { errorMessage } from '@/lib/api';
import { checkLink } from './style-utils';
import { MAX_IMAGE_WIDTH, type ButtonAlign, type ButtonVariant } from './constants';

function Field({ id, label, children, hint }: { id: string; label: string; children: React.ReactNode; hint?: React.ReactNode }) {
  return (
    <div className="grid gap-1.5">
      <Label htmlFor={id}>{label}</Label>
      {children}
      {hint}
    </div>
  );
}

function Note({ tone, children }: { tone: 'warn' | 'error'; children: React.ReactNode }) {
  return (
    <p role={tone === 'error' ? 'alert' : 'status'} className={tone === 'error' ? 'text-xs text-destructive' : 'flex items-start gap-1.5 text-xs text-warning'}>
      {tone === 'warn' && <AlertTriangle className="mt-px size-3.5 shrink-0" aria-hidden />}
      {children}
    </p>
  );
}

function Segmented<T extends string>({ value, options, onChange, label }: { value: T; options: { id: T; label: string }[]; onChange: (v: T) => void; label: string }) {
  return (
    <div role="radiogroup" aria-label={label} className="inline-flex rounded-md border border-input p-0.5">
      {options.map(o => (
        <button
          key={o.id}
          type="button"
          role="radio"
          aria-checked={value === o.id}
          onClick={() => onChange(o.id)}
          className={`rounded px-3 py-1 text-sm ${value === o.id ? 'bg-primary text-primary-foreground' : 'text-foreground hover:bg-muted'}`}
        >
          {o.label}
        </button>
      ))}
    </div>
  );
}

// ------------------------------------------------------------------ link

export interface LinkDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  initialHref: string;
  initialText: string;
  /** Text field is shown when there is no selection or the selection is a single link. */
  showText: boolean;
  isEditing: boolean;
  onSubmit: (v: { href: string; text: string }) => void;
  onRemove: () => void;
}

export function LinkDialog(p: LinkDialogProps) {
  const [href, setHref] = useState(p.initialHref);
  const [text, setText] = useState(p.initialText);
  const [touched, setTouched] = useState(false);
  useEffect(() => { if (p.open) { setHref(p.initialHref); setText(p.initialText); setTouched(false); } }, [p.open, p.initialHref, p.initialText]);
  const check = href.trim() ? checkLink(href) : null;
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    setTouched(true);
    const c = checkLink(href);
    if (!c.ok) return;
    p.onSubmit({ href: c.href, text: text.trim() });
  };
  return (
    <Dialog open={p.open} onOpenChange={p.onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <form onSubmit={submit} className="grid gap-4">
          <DialogHeader>
            <DialogTitle>{p.isEditing ? 'Edit link' : 'Insert link'}</DialogTitle>
            <DialogDescription>Links open in a new tab. Use https:// addresses or an email address.</DialogDescription>
          </DialogHeader>
          <Field id="kb-link-url" label="Address">
            <Input id="kb-link-url" autoFocus value={href} onChange={e => setHref(e.target.value)} placeholder="https://kiit.ac.in/… or name@kiit.ac.in" aria-invalid={touched && check?.ok === false} />
          </Field>
          {touched && check && !check.ok && <Note tone="error">{check.error}</Note>}
          {check?.ok && check.warning && <Note tone="warn">{check.warning}</Note>}
          {p.showText && (
            <Field id="kb-link-text" label="Text to display">
              <Input id="kb-link-text" value={text} onChange={e => setText(e.target.value)} placeholder="Defaults to the address" />
            </Field>
          )}
          <DialogFooter className="gap-2">
            {p.isEditing && <Button type="button" variant="outline" onClick={p.onRemove}>Remove link</Button>}
            <Button type="button" variant="ghost" onClick={() => p.onOpenChange(false)}>Cancel</Button>
            <Button type="submit">{p.isEditing ? 'Update' : 'Insert'}</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ button block

export interface ButtonAttrs { label: string; href: string; variant: ButtonVariant; align: ButtonAlign }

export function ButtonDialog({ open, onOpenChange, initial, onSubmit, isEditing }: {
  open: boolean; onOpenChange: (o: boolean) => void; initial: ButtonAttrs; onSubmit: (a: ButtonAttrs) => void; isEditing?: boolean;
}) {
  const [a, setA] = useState(initial);
  const [touched, setTouched] = useState(false);
  useEffect(() => { if (open) { setA(initial); setTouched(false); } }, [open]); // eslint-disable-line react-hooks/exhaustive-deps
  const check = a.href.trim() ? checkLink(a.href) : ({ ok: false, error: 'Enter the address the button opens.' } as const);
  const labelOk = a.label.trim().length > 0;
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    setTouched(true);
    if (!check.ok || !labelOk) return;
    onSubmit({ ...a, label: a.label.trim(), href: check.href });
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <form onSubmit={submit} className="grid gap-4">
          <DialogHeader>
            <DialogTitle>{isEditing ? 'Edit button' : 'Insert button'}</DialogTitle>
            <DialogDescription>A call-to-action button that renders in every mail client, including Outlook.</DialogDescription>
          </DialogHeader>
          <Field id="kb-btn-label" label="Button text">
            <Input id="kb-btn-label" autoFocus maxLength={60} value={a.label} onChange={e => setA({ ...a, label: e.target.value })} placeholder="Register now" aria-invalid={touched && !labelOk} />
          </Field>
          {touched && !labelOk && <Note tone="error">Enter the button text.</Note>}
          <Field id="kb-btn-href" label="Link address">
            <Input id="kb-btn-href" value={a.href} onChange={e => setA({ ...a, href: e.target.value })} placeholder="https://" aria-invalid={touched && !check.ok} />
          </Field>
          {touched && !check.ok && <Note tone="error">{check.error}</Note>}
          {check.ok && check.warning && <Note tone="warn">{check.warning}</Note>}
          <div className="flex flex-wrap gap-6">
            <div className="grid gap-1.5">
              <span className="text-sm font-medium">Style</span>
              <Segmented label="Button style" value={a.variant} onChange={v => setA({ ...a, variant: v })} options={[{ id: 'primary', label: 'Primary' }, { id: 'secondary', label: 'Secondary' }]} />
            </div>
            <div className="grid gap-1.5">
              <span className="text-sm font-medium">Alignment</span>
              <Segmented label="Button alignment" value={a.align} onChange={v => setA({ ...a, align: v })} options={[{ id: 'left', label: 'Left' }, { id: 'center', label: 'Centre' }]} />
            </div>
          </div>
          <DialogFooter className="gap-2">
            <Button type="button" variant="ghost" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button type="submit">{isEditing ? 'Update' : 'Insert'}</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ image insert

export interface InsertImageValue { src: string; alt: string; width: number | null }

const WIDTH_PRESETS = [25, 50, 75, 100];

function naturalWidth(url: string): Promise<number | null> {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => resolve(img.naturalWidth || null);
    img.onerror = () => resolve(null);
    img.src = url;
  });
}

export function ImageDialog({ open, onOpenChange, onInsert }: { open: boolean; onOpenChange: (o: boolean) => void; onInsert: (v: InsertImageValue) => void }) {
  const [file, setFile] = useState<File | null>(null);
  const [preview, setPreview] = useState<string | null>(null);
  const [natural, setNatural] = useState<number | null>(null);
  const [alt, setAlt] = useState('');
  const [pct, setPct] = useState(100);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    if (!open) {
      abortRef.current?.abort();
      setFile(null); setAlt(''); setPct(100); setBusy(false); setProgress(null); setError(null); setNatural(null);
    }
  }, [open]);
  useEffect(() => {
    if (!file) { setPreview(null); return; }
    const url = URL.createObjectURL(file);
    setPreview(url);
    naturalWidth(url).then(setNatural);
    if (!alt) setAlt(file.name.replace(/\.[a-z0-9]+$/i, '').replace(/[-_]+/g, ' ').trim());
    return () => URL.revokeObjectURL(url);
  }, [file]); // eslint-disable-line react-hooks/exhaustive-deps

  const pick = (f: File | undefined) => {
    setError(null);
    if (!f) return;
    if (!f.type.startsWith('image/')) { setError('Choose an image file (PNG, JPEG, GIF or WebP).'); return; }
    if (f.size > 10 * 1024 * 1024) { setError('Images must be under 10 MB. Compress it and try again.'); return; }
    setFile(f);
  };

  const base = Math.min(MAX_IMAGE_WIDTH, natural || MAX_IMAGE_WIDTH);
  const width = Math.round((MAX_IMAGE_WIDTH * pct) / 100) > base ? base : Math.round((MAX_IMAGE_WIDTH * pct) / 100);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) { setError('Choose an image first.'); return; }
    setBusy(true); setError(null); setProgress(0);
    const ctrl = new AbortController();
    abortRef.current = ctrl;
    try {
      const ref = await uploadFile(file, 'image', { signal: ctrl.signal, onProgress: p => setProgress(p.total ? p.loaded / p.total : null) });
      if (!ref.publicUrl) throw new Error('The server did not return an image address.');
      onInsert({ src: ref.publicUrl, alt: alt.trim(), width });
      onOpenChange(false);
    } catch (err) {
      if ((err as Error).name !== 'AbortError') setError(errorMessage(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={o => { if (!busy || !o) onOpenChange(o); }}>
      <DialogContent className="sm:max-w-md">
        <form onSubmit={submit} className="grid gap-4">
          <DialogHeader>
            <DialogTitle>Insert image</DialogTitle>
            <DialogDescription>The image is uploaded and linked from the email, so it is never sent as an embedded data URI.</DialogDescription>
          </DialogHeader>
          <input ref={inputRef} type="file" accept="image/png,image/jpeg,image/gif,image/webp" className="sr-only" onChange={e => pick(e.target.files?.[0])} aria-label="Image file" />
          <button
            type="button"
            onClick={() => inputRef.current?.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); pick(e.dataTransfer.files?.[0]); }}
            className="flex min-h-36 flex-col items-center justify-center gap-2 rounded-md border border-dashed border-input bg-muted/40 p-4 text-sm text-muted-foreground hover:bg-muted"
          >
            {preview ? (
              <img src={preview} alt="" className="max-h-40 max-w-full rounded border border-border bg-white object-contain" />
            ) : (
              <>
                <ImageUp className="size-6" aria-hidden />
                <span>Choose an image or drop it here</span>
              </>
            )}
          </button>
          {file && <p className="-mt-2 truncate text-xs text-muted-foreground">{file.name} · {(file.size / 1024).toFixed(0)} KB{natural ? ` · ${natural}px wide` : ''}</p>}
          <Field id="kb-img-alt" label="Alt text" hint={<p className="text-xs text-muted-foreground">Describes the image for screen readers and when images are blocked.</p>}>
            <Input id="kb-img-alt" value={alt} onChange={e => setAlt(e.target.value)} placeholder="e.g. Convocation 2026 poster" />
          </Field>
          <div className="grid gap-1.5">
            <span className="text-sm font-medium">Width</span>
            <Segmented label="Image width" value={String(pct)} onChange={v => setPct(Number(v))} options={WIDTH_PRESETS.map(w => ({ id: String(w), label: `${w}%` }))} />
            <p className="text-xs text-muted-foreground">{width}px of the {MAX_IMAGE_WIDTH}px email width{natural && natural < (MAX_IMAGE_WIDTH * pct) / 100 ? ' (limited to the image’s own size)' : ''}.</p>
          </div>
          {error && <Note tone="error">{error}</Note>}
          {busy && progress !== null && (
            <div className="h-1.5 w-full overflow-hidden rounded bg-muted" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(progress * 100)}>
              <div className="h-full bg-primary transition-[width]" style={{ width: `${Math.round(progress * 100)}%` }} />
            </div>
          )}
          <DialogFooter className="gap-2">
            <Button type="button" variant="ghost" onClick={() => onOpenChange(false)} disabled={busy}>Cancel</Button>
            <Button type="submit" disabled={!file || busy}>{busy ? <><Loader2 className="animate-spin" aria-hidden /> Uploading…</> : 'Insert'}</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ------------------------------------------------------------------ alt text

export function AltTextDialog({ open, onOpenChange, initial, onSubmit }: { open: boolean; onOpenChange: (o: boolean) => void; initial: string; onSubmit: (alt: string) => void }) {
  const [alt, setAlt] = useState(initial);
  useEffect(() => { if (open) setAlt(initial); }, [open, initial]);
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <form onSubmit={e => { e.preventDefault(); onSubmit(alt.trim()); onOpenChange(false); }} className="grid gap-4">
          <DialogHeader>
            <DialogTitle>Alt text</DialogTitle>
            <DialogDescription>Shown when images are blocked and read aloud by screen readers. Leave empty only for decorative images.</DialogDescription>
          </DialogHeader>
          <Input autoFocus aria-label="Alt text" value={alt} onChange={e => setAlt(e.target.value)} placeholder="Describe the image" />
          <DialogFooter className="gap-2">
            <Button type="button" variant="ghost" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button type="submit">Save</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
