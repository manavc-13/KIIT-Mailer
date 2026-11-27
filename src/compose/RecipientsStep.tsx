// Step 1: who gets the mail. Single recipient / CSV upload / manual grid, column mapping,
// validation, global CC/BCC, sending mode and de-duplication.
import { useRef, useState, type ClipboardEvent, type DragEvent } from 'react';
import Papa from 'papaparse';
import { toast } from 'sonner';
import {
  User, FileSpreadsheet, Table2, Plus, X, Upload, Download, Trash2, ClipboardPaste, Users, Send, Info, RotateCcw, Columns3,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Switch } from '@/components/ui/switch';
import { Checkbox } from '@/components/ui/checkbox';
import { Textarea } from '@/components/ui/textarea';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter,
  AlertDialogHeader, AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import { cn } from '@/lib/utils';
import { invalidEmails, parseTabular, toCSV, MAX_RECIPIENTS_PER_MESSAGE } from '../../shared/text';
import { RECIPIENT_ROLES, ROLE_LABELS, nonEmptyRows, type RecipientRow, type RoleMap } from '../../shared/recipients';
import type { JobMode } from '../../shared/types';
import { useComposeStore, type RecipientSource } from './store';
import { useRecipientModel, type RecipientModel } from './model';
import { Section, Segmented, StatPill, Chip, downloadText, pluralize } from './ui';

const NONE = '__none__';

export function RecipientsStep() {
  const source = useComposeStore(s => s.source);
  const set = useComposeStore(s => s.set);
  const model = useRecipientModel();

  return (
    <div className="space-y-5">
      <Section
        id="recipients-source"
        title="Recipients"
        description="Send to one person, upload a CSV, or type a list into the grid. Every column becomes a {Placeholder}."
        actions={
          <Segmented<RecipientSource>
            ariaLabel="Recipient source"
            value={source}
            onChange={v => set({ source: v, previewIndex: 0 })}
            options={[
              { value: 'single', label: 'Single', icon: User },
              { value: 'csv', label: 'CSV upload', icon: FileSpreadsheet },
              { value: 'manual', label: 'Manual grid', icon: Table2 },
            ]}
          />
        }
      >
        {source === 'single' && <SingleRecipient />}
        {source === 'csv' && <CsvRecipients model={model} />}
        {source === 'manual' && <ManualGrid model={model} />}
      </Section>

      {source !== 'single' && model.headers.length > 0 && <MappingPanel model={model} />}
      {model.rows.length > 0 && source !== 'single' && <ValidationCard model={model} />}

      <CopiesSection />
      <SendingModeSection model={model} />
    </div>
  );
}

// ------------------------------------------------------------------ single
function SingleRecipient() {
  const { singleName, singleEmail, singleExtras } = useComposeStore();
  const { set, addExtra, updateExtra, removeExtra } = useComposeStore();
  const bad = invalidEmails(singleEmail);

  return (
    <div className="space-y-4">
      <div className="grid gap-4 sm:grid-cols-2">
        <div className="space-y-1.5">
          <Label htmlFor="single-name">Name</Label>
          <Input id="single-name" value={singleName} placeholder="Dr. A. Rao" autoComplete="off"
            onChange={e => set({ singleName: e.target.value })} />
        </div>
        <div className="space-y-1.5">
          <Label htmlFor="single-email">Email <span className="text-destructive" aria-hidden>*</span></Label>
          <Input id="single-email" type="text" inputMode="email" value={singleEmail} placeholder="name@kiit.ac.in"
            aria-invalid={bad.length > 0} aria-describedby="single-email-hint" autoComplete="off"
            onChange={e => set({ singleEmail: e.target.value })} />
          <p id="single-email-hint" className={cn('text-xs', bad.length ? 'text-destructive' : 'text-muted-foreground')}>
            {bad.length ? `Not a valid address: ${bad.join(', ')}` : 'Several addresses separated by commas all go in To of one message.'}
          </p>
        </div>
      </div>

      <div className="space-y-2">
        <div className="flex items-center justify-between">
          <div>
            <div className="text-sm font-medium">Custom fields</div>
            <p className="text-xs text-muted-foreground">Optional placeholders for this message, e.g. Department → {'{Department}'}.</p>
          </div>
          <Button type="button" variant="outline" size="sm" onClick={addExtra}><Plus /> Add field</Button>
        </div>
        {singleExtras.map((f, i) => (
          <div key={f.id} className="flex items-center gap-2">
            <Input aria-label={`Field ${i + 1} name`} placeholder="Field name" value={f.key} className="max-w-[14rem]"
              onChange={e => updateExtra(f.id, { key: e.target.value.replace(/[{}]/g, '') })} />
            <Input aria-label={`Field ${i + 1} value`} placeholder="Value" value={f.value}
              onChange={e => updateExtra(f.id, { value: e.target.value })} />
            <Button type="button" variant="ghost" size="icon-sm" aria-label={`Remove field ${f.key || i + 1}`} onClick={() => removeExtra(f.id)}>
              <X />
            </Button>
          </div>
        ))}
      </div>
    </div>
  );
}

// ------------------------------------------------------------------ csv
const SAMPLE_CSV = 'Name,Email,CC,BCC,Department\n'
  + 'Dr. A Rao,a.rao@kiit.ac.in,"hod.cse@kiit.ac.in; dean@kiit.ac.in",audit@kiit.ac.in,CSE\n'
  + 'John Doe,john.doe@example.com,,,Admin\n';

function CsvRecipients({ model }: { model: RecipientModel }) {
  const { csvFileName, csvRows, csvNeedsReupload } = useComposeStore();
  const { loadCsv, clearCsv } = useComposeStore();
  const inputRef = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  const [parsing, setParsing] = useState(false);

  const parse = (file: File) => {
    if (!/\.(csv|tsv|txt)$/i.test(file.name) && !/csv|text/.test(file.type)) {
      toast.error(`"${file.name}" doesn't look like a CSV file.`);
      return;
    }
    setParsing(true);
    Papa.parse<RecipientRow>(file, {
      header: true,
      skipEmptyLines: 'greedy',
      transformHeader: h => h.replace(/^﻿/, '').trim(),
      complete: res => {
        setParsing(false);
        const headers = (res.meta.fields ?? []).filter(Boolean);
        const rows = (res.data ?? []).map(r => Object.fromEntries(headers.map(h => [h, String(r[h] ?? '').trim()])));
        if (!rows.length || !headers.length) { toast.error('The CSV is empty or has no header row.'); return; }
        loadCsv(file.name, headers, rows);
        const detected = useComposeStore.getState().csvRoleMap;
        if (!detected.email) toast.warning('Could not find an Email column. Pick it in the mapping panel below.');
        else toast.success(`Loaded ${pluralize(rows.length, 'row')} from ${file.name}`);
      },
      error: err => { setParsing(false); toast.error(`Could not read the CSV: ${err.message}`); },
    });
  };

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) parse(f);
  };

  return (
    <div className="space-y-3">
      {csvNeedsReupload && (
        <div role="status" className="flex items-start gap-2 rounded-md border border-warning/30 bg-warning-soft px-3 py-2 text-sm text-warning">
          <Info className="mt-0.5 size-4 shrink-0" />
          <span>Your draft used <strong>{csvFileName ?? 'a CSV file'}</strong>, which was too large to keep in the browser. Upload it again to continue.</span>
        </div>
      )}
      <div
        role="button"
        tabIndex={0}
        aria-label="Upload a CSV file: drop it here or press Enter to browse"
        onClick={() => inputRef.current?.click()}
        onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputRef.current?.click(); } }}
        onDragOver={e => { e.preventDefault(); e.stopPropagation(); e.dataTransfer.dropEffect = 'copy'; setOver(true); }}
        onDragLeave={() => setOver(false)}
        onDrop={onDrop}
        className={cn(
          'flex cursor-pointer flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed px-6 py-8 text-center transition-colors outline-none focus-visible:ring-2 focus-visible:ring-ring',
          over ? 'border-gold bg-gold-soft' : 'border-border bg-muted/40 hover:bg-muted',
        )}
      >
        <Upload className="size-6 text-muted-foreground" strokeWidth={1.5} />
        <div className="text-sm font-medium">{parsing ? 'Reading…' : 'Drop a CSV here, or click to browse'}</div>
        <div className="text-xs text-muted-foreground">First row must be the column names. Excel: File → Save As → CSV UTF-8.</div>
        <input ref={inputRef} type="file" accept=".csv,.tsv,.txt,text/csv" className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) parse(f); e.target.value = ''; }} />
      </div>
      <div className="flex flex-wrap items-center justify-between gap-2 text-sm">
        <span className="text-muted-foreground">
          {csvRows
            ? <>Loaded <strong className="text-foreground">{pluralize(csvRows.length, 'row')}</strong> from <strong className="text-foreground">{csvFileName}</strong> · columns: {model.headers.join(', ')}</>
            : 'No file loaded yet.'}
        </span>
        <div className="flex gap-2">
          <Button type="button" variant="ghost" size="sm" onClick={() => downloadText(SAMPLE_CSV, 'kiit_mailer_sample.csv', 'text/csv')}>
            <Download /> Sample CSV
          </Button>
          {csvRows && <Button type="button" variant="ghost" size="sm" onClick={clearCsv}><X /> Remove file</Button>}
        </div>
      </div>
      {csvRows && csvRows.length > 0 && <CsvPreviewTable headers={model.headers} rows={csvRows.slice(0, 5)} total={csvRows.length} />}
    </div>
  );
}

function CsvPreviewTable({ headers, rows, total }: { headers: string[]; rows: RecipientRow[]; total: number }) {
  return (
    <div className="overflow-x-auto rounded-md border">
      <table className="w-full text-left text-xs">
        <caption className="sr-only">First rows of the uploaded CSV</caption>
        <thead className="bg-muted text-muted-foreground">
          <tr>{headers.map(h => <th key={h} scope="col" className="whitespace-nowrap px-3 py-1.5 font-medium">{h}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} className="border-t">
              {headers.map(h => <td key={h} className="max-w-[16rem] truncate px-3 py-1.5">{r[h]}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
      {total > rows.length && <div className="border-t bg-muted/40 px-3 py-1.5 text-xs text-muted-foreground">…and {(total - rows.length).toLocaleString()} more rows</div>}
    </div>
  );
}

// ------------------------------------------------------------------ manual grid
function ManualGrid({ model }: { model: RecipientModel }) {
  const { gridColumns, gridRows, gridRoleMap } = useComposeStore();
  const {
    setGridCell, addGridRow, removeGridRows, addGridColumn, renameGridColumn, removeGridColumn, importGrid, pasteIntoGrid, resetGrid,
  } = useComposeStore();
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [newCol, setNewCol] = useState('');
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteText, setPasteText] = useState('');
  const [pasteAppend, setPasteAppend] = useState(true);
  const [confirm, setConfirm] = useState<null | 'clear' | 'delete'>(null);
  const tableRef = useRef<HTMLTableElement>(null);

  const sel = new Set([...selected].filter(i => i < gridRows.length));
  const allSelected = gridRows.length > 0 && sel.size === gridRows.length;
  const filled = nonEmptyRows(gridRows);
  const withEmail = gridRoleMap.email ? filled.filter(r => String(r[gridRoleMap.email!] ?? '').trim()).length : 0;

  const onAddColumn = () => {
    if (!newCol.trim()) return;
    if (!addGridColumn(newCol.replace(/[{}]/g, ''))) toast.error('That column already exists.');
    else setNewCol('');
  };

  const onCellPaste = (e: ClipboardEvent<HTMLInputElement>, ri: number, col: string) => {
    const text = e.clipboardData.getData('text');
    if (!text || !/[\t\n]/.test(text)) return;
    e.preventDefault();
    const n = pasteIntoGrid(ri, col, text);
    toast.success(`Pasted ${pluralize(n, 'row')}`);
  };

  const applyImport = () => {
    const { columns, rows } = parseTabular(pasteText.trim());
    if (!rows.length) { toast.warning('No rows detected. Include a header row, then one row per recipient.'); return; }
    importGrid(columns, rows, pasteAppend);
    setSelected(new Set());
    setPasteOpen(false); setPasteText('');
    toast.success(`Imported ${pluralize(rows.length, 'row')}`);
  };

  const exportCsv = () => {
    if (!filled.length) { toast.warning('The grid is empty.'); return; }
    downloadText(toCSV(gridColumns, filled), 'kiit_mailer_recipients.csv', 'text/csv');
  };

  // Enter in the last row adds a row; Enter elsewhere moves down a row (spreadsheet feel).
  const focusCell = (ri: number, ci: number) => {
    requestAnimationFrame(() => {
      tableRef.current?.querySelector<HTMLInputElement>(`input[data-ri="${ri}"][data-ci="${ci}"]`)?.focus();
    });
  };

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <Button type="button" variant="outline" size="sm" onClick={() => { addGridRow(); focusCell(gridRows.length, 0); }}><Plus /> Row</Button>
        <div className="flex items-center gap-1">
          <Input aria-label="New column name" placeholder="New column" value={newCol} className="h-8 w-36"
            onChange={e => setNewCol(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); onAddColumn(); } }} />
          <Button type="button" variant="outline" size="sm" onClick={onAddColumn} disabled={!newCol.trim()}><Columns3 /> Add column</Button>
        </div>
        <Button type="button" variant="outline" size="sm" aria-expanded={pasteOpen} onClick={() => setPasteOpen(o => !o)}><ClipboardPaste /> Paste from Excel</Button>
        <Button type="button" variant="ghost" size="sm" onClick={exportCsv}><Download /> Export CSV</Button>
        <Button type="button" variant="ghost" size="sm" onClick={() => setConfirm('clear')}><RotateCcw /> Clear grid</Button>
        {sel.size > 0 && (
          <Button type="button" variant="destructive" size="sm" className="ml-auto" onClick={() => setConfirm('delete')}>
            <Trash2 /> Delete {sel.size} selected
          </Button>
        )}
      </div>

      {pasteOpen && (
        <div className="space-y-2 rounded-md border bg-muted/40 p-3">
          <Label htmlFor="grid-paste">Paste rows copied from Excel or Google Sheets (first line = column names)</Label>
          <Textarea id="grid-paste" rows={5} value={pasteText} autoFocus className="font-mono text-xs"
            placeholder={'Name\tEmail\tDepartment\nDr. A Rao\ta.rao@kiit.ac.in\tCSE'}
            onChange={e => setPasteText(e.target.value)} />
          <div className="flex flex-wrap items-center gap-3">
            <label className="flex items-center gap-2 text-sm">
              <Checkbox checked={pasteAppend} onCheckedChange={v => setPasteAppend(v === true)} /> Append to existing rows
            </label>
            <div className="ml-auto flex gap-2">
              <Button type="button" variant="ghost" size="sm" onClick={() => { setPasteOpen(false); setPasteText(''); }}>Cancel</Button>
              <Button type="button" size="sm" onClick={applyImport} disabled={!pasteText.trim()}>Import rows</Button>
            </div>
          </div>
        </div>
      )}

      <div className="max-h-[28rem] overflow-auto rounded-md border">
        <table ref={tableRef} className="w-full border-collapse text-sm">
          <caption className="sr-only">Recipient grid. Paste multiple cells from a spreadsheet into any cell.</caption>
          <thead className="sticky top-0 z-10 bg-muted">
            <tr>
              <th scope="col" className="w-9 border-b px-2 py-1.5">
                <Checkbox aria-label="Select all rows" checked={allSelected ? true : sel.size ? 'indeterminate' : false}
                  onCheckedChange={v => setSelected(v === true ? new Set(gridRows.map((_, i) => i)) : new Set())} />
              </th>
              {gridColumns.map((col, ci) => (
                <th key={`${ci}-${col}`} scope="col" className="min-w-[9rem] border-b border-l px-1 py-1 text-left font-medium">
                  <div className="flex items-center gap-1">
                    <input
                      aria-label={`Column ${ci + 1} name`}
                      defaultValue={col}
                      readOnly={col === gridRoleMap.email}
                      title={col === gridRoleMap.email ? 'Email column (required)' : 'Rename column'}
                      className="h-7 w-full min-w-0 rounded bg-transparent px-1.5 text-[13px] font-semibold outline-none focus-visible:bg-card focus-visible:ring-1 focus-visible:ring-ring read-only:cursor-default"
                      onBlur={e => { if (!renameGridColumn(ci, e.target.value.replace(/[{}]/g, ''))) { toast.error('Column names must be unique and not empty.'); e.target.value = col; } }}
                      onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); }}
                    />
                    {col !== gridRoleMap.email && (
                      <Button type="button" variant="ghost" size="icon-xs" aria-label={`Delete column ${col}`} onClick={() => removeGridColumn(ci)}><X /></Button>
                    )}
                  </div>
                </th>
              ))}
              <th scope="col" className="w-9 border-b border-l"><span className="sr-only">Row actions</span></th>
            </tr>
          </thead>
          <tbody>
            {gridRows.map((row, ri) => (
              <tr key={ri} className={cn('border-b last:border-b-0', sel.has(ri) && 'bg-gold-soft')}>
                <td className="px-2 text-center">
                  <Checkbox aria-label={`Select row ${ri + 1}`} checked={sel.has(ri)}
                    onCheckedChange={v => setSelected(prev => { const n = new Set(prev); if (v === true) n.add(ri); else n.delete(ri); return n; })} />
                </td>
                {gridColumns.map((col, ci) => (
                  <td key={col} className="border-l p-0">
                    <input
                      data-ri={ri} data-ci={ci}
                      aria-label={`Row ${ri + 1} ${col}`}
                      type="text"
                      inputMode={col === gridRoleMap.email ? 'email' : undefined}
                      value={row[col] ?? ''}
                      placeholder={col}
                      className={cn(
                        'h-9 w-full bg-transparent px-2.5 outline-none placeholder:text-muted-foreground/50 focus:bg-card focus:ring-2 focus:ring-inset focus:ring-ring',
                        col === gridRoleMap.email && row[col] && invalidEmails(row[col]).length > 0 && 'text-destructive',
                      )}
                      onChange={e => setGridCell(ri, col, e.target.value)}
                      onPaste={e => onCellPaste(e, ri, col)}
                      onKeyDown={e => {
                        if (e.key !== 'Enter') return;
                        e.preventDefault();
                        if (ri === gridRows.length - 1) addGridRow();
                        focusCell(ri + 1, ci);
                      }}
                    />
                  </td>
                ))}
                <td className="border-l text-center">
                  <Button type="button" variant="ghost" size="icon-xs" aria-label={`Delete row ${ri + 1}`}
                    onClick={() => { removeGridRows([ri]); setSelected(new Set()); }}>
                    <X />
                  </Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="text-xs text-muted-foreground">
        {pluralize(gridRows.length, 'row')} · {withEmail} with email · {model.headers.length} columns. Tip: paste a block of cells
        from Excel straight into any cell; press Enter to move down.
      </p>

      <AlertDialog open={confirm !== null} onOpenChange={o => { if (!o) setConfirm(null); }}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>{confirm === 'clear' ? 'Clear the grid?' : `Delete ${pluralize(sel.size, 'row')}?`}</AlertDialogTitle>
            <AlertDialogDescription>
              {confirm === 'clear' ? 'All rows and custom columns will be removed. This cannot be undone.' : 'The selected rows will be removed from the grid.'}
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Keep</AlertDialogCancel>
            <AlertDialogAction variant="destructive" onClick={() => {
              if (confirm === 'clear') resetGrid(); else removeGridRows([...sel]);
              setSelected(new Set());
            }}>
              {confirm === 'clear' ? 'Clear grid' : 'Delete rows'}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}

// ------------------------------------------------------------------ mapping + validation
function MappingPanel({ model }: { model: RecipientModel }) {
  const source = useComposeStore(s => s.source);
  const setCsvRole = useComposeStore(s => s.setCsvRole);
  const setGridRole = useComposeStore(s => s.setGridRole);
  const setRole = (role: keyof RoleMap, v: string | null) => (source === 'csv' ? setCsvRole(role, v) : setGridRole(role, v));

  return (
    <Section id="recipients-mapping" title="Column mapping" description="Which column holds each part of the address. Detected automatically; change it if needed.">
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        {RECIPIENT_ROLES.map(role => {
          const value = model.roleMap[role];
          return (
            <div key={role} className="space-y-1.5">
              <Label htmlFor={`map-${role}`} className="flex items-center justify-between">
                <span>{ROLE_LABELS[role]}{role === 'email' && <span className="text-destructive"> *</span>}</span>
                <span className={cn('text-xs', value ? 'text-success' : role === 'email' ? 'text-destructive' : 'text-muted-foreground')}>
                  {value ? 'Mapped' : role === 'email' ? 'Required' : 'Not used'}
                </span>
              </Label>
              <Select value={value ?? NONE} onValueChange={v => setRole(role, v === NONE ? null : v)}>
                <SelectTrigger id={`map-${role}`} className="w-full"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value={NONE}>(none)</SelectItem>
                  {model.headers.map(h => <SelectItem key={h} value={h}>{h}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
          );
        })}
      </div>
      <p className="mt-3 text-xs text-muted-foreground">
        CC / BCC cells may hold several addresses separated by <code>;</code>. They are merged with the global CC/BCC below.
      </p>
    </Section>
  );
}

function ValidationCard({ model }: { model: RecipientModel }) {
  const v = model.validation;
  const dedupe = useComposeStore(s => s.dedupe);
  const mode = useComposeStore(s => s.mode);
  const list = (items: string[], tone: 'danger' | 'warn') => (
    <div className="flex flex-wrap gap-1">
      {items.slice(0, 12).map(e => <Chip key={e} tone={tone}>{e}</Chip>)}
      {items.length > 12 && <span className="text-xs text-muted-foreground">+{items.length - 12} more</span>}
    </div>
  );
  return (
    <Section id="recipients-validation" title="Validation" description="Rows without a valid email are skipped when sending.">
      <div className="flex flex-wrap gap-2">
        <StatPill label="Rows" value={v.total.toLocaleString()} tone={v.total ? 'neutral' : 'warn'} />
        <StatPill label="Valid emails" value={v.valid.toLocaleString()} tone={v.valid ? 'ok' : 'danger'} />
        <StatPill label="Invalid" value={v.invalid.length} tone={v.invalid.length ? 'danger' : 'ok'} />
        <StatPill label="Duplicates" value={v.duplicateRows} tone={v.duplicateRows ? 'warn' : 'ok'} />
        <StatPill label="Missing email" value={v.missingEmail} tone={v.missingEmail ? 'warn' : 'ok'} />
        <StatPill label="Missing name" value={v.missingName} tone={v.missingName ? 'warn' : 'ok'} />
      </div>
      {!model.roleMap.email && <p className="mt-3 text-sm text-destructive">Map the Email column to continue.</p>}
      {v.invalid.length > 0 && <div className="mt-3 space-y-1"><div className="text-xs font-medium">Invalid addresses</div>{list(v.invalid, 'danger')}</div>}
      {v.duplicates.length > 0 && (
        <div className="mt-3 space-y-1">
          <div className="text-xs font-medium">
            Duplicates {mode === 'bcc_broadcast' || dedupe ? '(only the first will be sent)' : '(each copy will be sent: turn on de-duplicate below to avoid this)'}
          </div>
          {list(v.duplicates, 'warn')}
        </div>
      )}
      {model.build.invalidCopies.length > 0 && (
        <div className="mt-3 space-y-1"><div className="text-xs font-medium">Invalid CC/BCC addresses (dropped)</div>{list(model.build.invalidCopies, 'warn')}</div>
      )}
    </Section>
  );
}

// ------------------------------------------------------------------ CC / BCC
function CopiesSection() {
  const { globalCc, globalBcc, set, mode } = useComposeStore();
  const badCc = invalidEmails(globalCc).filter(e => !/^\{.+\}$/.test(e));
  const badBcc = invalidEmails(globalBcc).filter(e => !/^\{.+\}$/.test(e));
  return (
    <Section id="recipients-copies" title="Copies for every message"
      description={mode === 'personal'
        ? 'Added to each email, merged with any CC/BCC columns. Never repeats the To address. Placeholders such as {Supervisor} work.'
        : 'Added to every broadcast message. Row placeholders are not available in broadcast mode.'}>
      <div className="grid gap-4 sm:grid-cols-2">
        <div className="space-y-1.5">
          <Label htmlFor="global-cc">CC</Label>
          <Input id="global-cc" value={globalCc} placeholder="registrar@kiit.ac.in; dean@kiit.ac.in" aria-invalid={badCc.length > 0}
            onChange={e => set({ globalCc: e.target.value })} />
          {badCc.length > 0 && <p className="text-xs text-destructive">Not valid: {badCc.join(', ')}</p>}
        </div>
        <div className="space-y-1.5">
          <Label htmlFor="global-bcc">BCC</Label>
          <Input id="global-bcc" value={globalBcc} placeholder="iqac.records@kiit.ac.in" aria-invalid={badBcc.length > 0}
            onChange={e => set({ globalBcc: e.target.value })} />
          {badBcc.length > 0 && <p className="text-xs text-destructive">Not valid: {badBcc.join(', ')}</p>}
        </div>
      </div>
    </Section>
  );
}

// ------------------------------------------------------------------ mode + dedupe
function SendingModeSection({ model }: { model: RecipientModel }) {
  const { mode, dedupe, source, set } = useComposeStore();
  const options: { value: JobMode; title: string; icon: typeof Users; body: string; badge?: string }[] = [
    {
      value: 'personal', icon: Send, title: 'Personalized', badge: 'Recommended',
      body: 'Each person gets their own email. Best inbox placement, and placeholders such as {Name} work for every recipient.',
    },
    {
      value: 'bcc_broadcast', icon: Users, title: 'One email, everyone in BCC',
      body: `One message to the whole list; recipients can't see each other. Sent in groups of ${MAX_RECIPIENTS_PER_MESSAGE} because Gmail limits recipients per message. Only global placeholders work.`,
    },
  ];
  return (
    <Section id="recipients-mode" title="Sending mode">
      <div role="radiogroup" aria-label="Sending mode" className="grid gap-3 md:grid-cols-2">
        {options.map(o => {
          const active = mode === o.value;
          const Icon = o.icon;
          return (
            <button
              key={o.value} type="button" role="radio" aria-checked={active}
              onClick={() => set({ mode: o.value })}
              className={cn(
                'flex gap-3 rounded-md border p-4 text-left transition-colors outline-none focus-visible:ring-2 focus-visible:ring-ring',
                active ? 'border-primary bg-accent/60 ring-1 ring-primary dark:border-gold dark:ring-gold' : 'bg-card hover:bg-muted/60',
              )}
            >
              <span className={cn('mt-0.5 grid size-5 shrink-0 place-items-center rounded-full border-2', active ? 'border-primary dark:border-gold' : 'border-input')}>
                {active && <span className="size-2.5 rounded-full bg-primary dark:bg-gold" />}
              </span>
              <span className="min-w-0">
                <span className="flex flex-wrap items-center gap-2 font-medium">
                  <Icon className="size-4 text-muted-foreground" strokeWidth={1.75} /> {o.title}
                  {o.badge && <Chip tone="ok">{o.badge}</Chip>}
                </span>
                <span className="mt-1 block text-[13px] text-muted-foreground">{o.body}</span>
              </span>
            </button>
          );
        })}
      </div>

      {mode === 'bcc_broadcast' && model.build.recipientCount > 0 && (
        <p className="mt-3 text-sm text-muted-foreground">
          {pluralize(model.build.recipientCount, 'address', 'addresses')} → {pluralize(model.build.messageCount, 'message')}. Each message is addressed To your sender account.
          {model.build.ignoredRowCopies && ' Per-row CC/BCC columns are ignored in this mode.'}
        </p>
      )}

      <div className="mt-4 flex flex-wrap items-start justify-between gap-3 border-t pt-4">
        <div className="flex items-start gap-3">
          <Switch id="dedupe" checked={mode === 'bcc_broadcast' || dedupe} disabled={mode === 'bcc_broadcast' || source === 'single'}
            onCheckedChange={v => set({ dedupe: v })} />
          <div>
            <Label htmlFor="dedupe">Remove duplicate addresses</Label>
            <p className="text-xs text-muted-foreground">Send once per address even if it appears on several rows.</p>
          </div>
        </div>
        <p className="flex max-w-md items-start gap-1.5 text-xs text-muted-foreground">
          <Info className="mt-0.5 size-3.5 shrink-0" />
          Addresses that unsubscribed or bounced earlier are skipped automatically when the job is created.
        </p>
      </div>
    </Section>
  );
}
