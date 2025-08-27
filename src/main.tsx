// src/main.tsx
import React, { useEffect, useMemo, useState } from 'react';
import ReactDOM from 'react-dom/client';
import {
  connect,
  type RenderFieldExtensionCtx,
} from 'datocms-plugin-sdk';
import {
  Canvas,
  Button,
  TextField,
  Spinner,
} from 'datocms-react-ui';
import { buildClient } from '@datocms/cma-client-browser';
import * as XLSX from 'xlsx';
import { AgGridReact } from 'ag-grid-react';

// AG Grid v34 modules (Theming API + editors)
import {
  ModuleRegistry,
  ClientSideRowModelModule,
  TextEditorModule,
  themeQuartz,
} from 'ag-grid-community';

ModuleRegistry.registerModules([
  ClientSideRowModelModule,
  TextEditorModule,
]);

// Dato UI styles
import 'datocms-react-ui/styles.css';

// ======= Config / fallbacks =======
const DEFAULT_SOURCE_FILE_API_KEY = 'sourcefile'; // your file field API key

type TableRow = Record<string, unknown>;

type FieldParams = {
  sourceFileApiKey?: string;
  columnsMetaApiKey?: string;
  rowCountApiKey?: string;
};

// ======= Helpers =======
function getUrlOverrideApiKey(): string | undefined {
  try {
    const u = new URL(window.location.href);
    return u.searchParams.get('fileApiKey') || undefined;
  } catch {
    return undefined;
  }
}

function getEditorParams(ctx: RenderFieldExtensionCtx): FieldParams {
  const direct = (ctx.parameters as any) || {};
  if (direct && Object.keys(direct).length) return direct;

  const appearance =
    (ctx.field as any)?.attributes?.appearance?.parameters ||
    (ctx as any)?.fieldAppearance?.parameters ||
    {};
  return appearance;
}

function resolveFieldId(
  ctx: RenderFieldExtensionCtx,
  preferred?: string | null,
): string | null {
  const fields = Object.values(ctx.fields) as any[];

  if (preferred) {
    const byId = (ctx.fields as any)[preferred];
    if (byId?.id) return byId.id;

    const match = fields.find((f) => (f.apiKey ?? f.attributes?.api_key) === preferred);
    if (match) return match.id;
  }

  const firstFile = fields.find(
    (f) => (f.fieldType ?? f.attributes?.field_type) === 'file',
  );
  return firstFile?.id ?? null;
}

function pickLocalizedValue(raw: any, locale?: string | null) {
  if (
    raw &&
    typeof raw === 'object' &&
    raw !== null &&
    locale &&
    Object.prototype.hasOwnProperty.call(raw, locale)
  ) {
    return raw[locale];
  }
  return raw ?? null;
}

async function fetchUploadUrlFromValue(
  fileFieldValue: any,
  cmaToken: string,
): Promise<string | null> {
  const uploadId = fileFieldValue?.upload_id;
  if (!uploadId) return null;
  const client = buildClient({ apiToken: cmaToken });
  const upload = await client.uploads.find(String(uploadId));
  return (upload as any)?.url || null;
}

function toSheetJSRows(
  binary: ArrayBuffer,
  preferredSheet?: string,
): { rows: TableRow[]; sheetNames: string[] } {
  const wb = XLSX.read(binary, { type: 'array' });
  const names = wb.SheetNames;
  const target =
    preferredSheet && names.includes(preferredSheet) ? preferredSheet : names[0];
  const ws = wb.Sheets[target];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null }) as TableRow[];
  return { rows, sheetNames: names };
}

function findFirstUploadInObject(obj: any, locale?: string | null): any | null {
  if (!obj || typeof obj !== 'object') return null;

  const localized = pickLocalizedValue(obj, locale);
  const cur = localized ?? obj;

  if (Array.isArray(cur)) {
    for (const item of cur) {
      const hit = findFirstUploadInObject(item, locale);
      if (hit) return hit;
    }
    return null;
  }

  if (cur?.upload_id || cur?.upload?.id || (typeof cur === 'string' && cur.startsWith('http'))) {
    return cur;
  }

  for (const key of Object.keys(cur)) {
    const hit = findFirstUploadInObject(cur[key], locale);
    if (hit) return hit;
  }
  return null;
}

// ---- Strict/safe normalization: keys and string values ----
function sanitizeKey(raw: string, index: number): string {
  const base = (raw || '').toString().trim();
  const candidate = base && base !== '__EMPTY' ? base : `column_${index + 1}`;
  const cleaned = candidate.replace(/[^a-zA-Z0-9_ ]+/g, ' ').replace(/\s+/g, ' ').trim();
  return cleaned || `column_${index + 1}`;
}
function uniqueKeys(keys: string[]): string[] {
  const seen = new Map<string, number>();
  return keys.map((k) => {
    const c = seen.get(k) ?? 0;
    seen.set(k, c + 1);
    return c === 0 ? k : `${k}_${c + 1}`;
  });
}
function toStringValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' && Number.isNaN(v)) return '';
  return String(v);
}
/** Normalize to safe column names and **string** cell values */
function normalizeSheetRowsStrings(rows: TableRow[]): { rows: TableRow[]; columns: string[] } {
  if (!rows || rows.length === 0) return { rows: [], columns: [] };
  const firstRow = rows[0] as Record<string, unknown>;
  const candidateKeys = Object.keys(firstRow);
  const safe = uniqueKeys(candidateKeys.map((k, i) => sanitizeKey(k, i)));

  const keyMap = new Map<string, string>();
  candidateKeys.forEach((orig, i) => keyMap.set(orig, safe[i]));

  const normalizedRows = rows.map((r) => {
    const obj = r as Record<string, unknown>;
    const out: Record<string, string> = {};
    for (const [origKey, val] of Object.entries(obj)) {
      const mapped = keyMap.get(origKey) ?? sanitizeKey(origKey, 0);
      out[mapped] = toStringValue(val);
    }
    // ensure all columns present
    safe.forEach((k) => { if (!(k in out)) out[k] = ''; });
    return out;
  });

  return { rows: normalizedRows, columns: safe };
}

// resolve id/apiKey to the correct path (handles locale)
function getFieldPath(ctx: RenderFieldExtensionCtx, apiKeyOrId: string): string | null {
  const fields = Object.values(ctx.fields) as any[];
  const byId = (ctx.fields as any)[apiKeyOrId];
  const id =
    byId?.id ?? (fields.find(f => (f.apiKey ?? f.attributes?.api_key) === apiKeyOrId)?.id);
  if (!id) return null;
  return ctx.locale ? `${id}.${ctx.locale}` : id;
}

async function setFieldByApiOrId(ctx: RenderFieldExtensionCtx, apiKeyOrId: string, value: unknown) {
  const path = getFieldPath(ctx, apiKeyOrId);
  if (path) await ctx.setFieldValue(path, value);
}

function Alert({ children }: { children: React.ReactNode }) {
  return (
    <div
      role="alert"
      style={{
        padding: '8px 12px',
        border: '1px solid var(--border-color)',
        borderRadius: 6,
        marginBottom: 8,
      }}
    >
      {children}
    </div>
  );
}

// ======= Editor =======
function Editor({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = getEditorParams(ctx);
  const preferredApiKey =
    getUrlOverrideApiKey() || params.sourceFileApiKey || DEFAULT_SOURCE_FILE_API_KEY;

  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);

  const [sheet, setSheet] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);

  const [rows, setRows] = useState<TableRow[]>(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    // Load rows from existing object wrapper if present
    if (initial && typeof initial === 'object' && !Array.isArray(initial) && (initial as any).rows) {
      return ((initial as any).rows as TableRow[]) || [];
    }
    return Array.isArray(initial) ? (initial as TableRow[]) : [];
  });

  const columnDefs = useMemo(() => {
    const cols = new Set<string>();
    rows.forEach((r) => Object.keys(r as object).forEach((k) => cols.add(k)));
    if (cols.size === 0) cols.add('column1');
    return Array.from(cols).map((c) => ({ field: c, editable: true }));
  }, [rows]);

  function getFileFieldValue() {
    const fileFieldId = resolveFieldId(ctx, preferredApiKey);
    let raw = fileFieldId ? (ctx.formValues as any)[fileFieldId] : undefined;

    raw = pickLocalizedValue(raw, ctx.locale);
    if (!raw) {
      const found = findFirstUploadInObject(ctx.formValues, ctx.locale);
      raw = found || null;
    }
    if (!raw) return null;

    if (Array.isArray(raw)) raw = raw[0];

    if (raw?.upload_id) return raw;
    if (raw?.upload?.id) return { upload_id: raw.upload.id };
    if (typeof raw === 'string' && raw.startsWith('http')) return { __direct_url: raw };
    return null;
  }

  async function importFromSource() {
    try {
      setBusy(true);
      setNotice(null);

      const fileVal = getFileFieldValue();
      if (!fileVal) {
        setNotice(
          'No file in the configured file field. Check the "Source File API key" or upload a file (locale-aware).',
        );
        return;
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      let url: string | null = null;

      if ((fileVal as any).__direct_url) {
        url = (fileVal as any).__direct_url as string;
      } else {
        if (!token) {
          setNotice(
            'Missing CMA token in plugin configuration (Settings → Plugins → this plugin → Configuration).',
          );
          return;
        }
        url = await fetchUploadUrlFromValue(fileVal, token);
      }

      if (!url) {
        setNotice('Could not resolve upload URL from the file field value.');
        return;
      }

      const res = await fetch(url);
      const buf = await res.arrayBuffer();
      const { rows: parsed, sheetNames: names } = toSheetJSRows(buf, sheet || undefined);

      const normalized = normalizeSheetRowsStrings(parsed as TableRow[]);
      setSheetNames(names);
      setRows(normalized.rows);
      setSheet(names[0] || null);

      // write ONLY { rows } into the field to satisfy strict JSON validation
      await ctx.setFieldValue(ctx.fieldPath, { rows: normalized.rows });

      // Optional meta fields
      if (params.columnsMetaApiKey) {
        await setFieldByApiOrId(ctx, params.columnsMetaApiKey, { columns: normalized.columns });
      }
      if (params.rowCountApiKey) {
        await setFieldByApiOrId(ctx, params.rowCountApiKey, Number(normalized.rows.length));
      }
    } catch (e: any) {
      setNotice(`Import failed: ${e?.message || e}`);
    } finally {
      setBusy(false);
    }
  }

  async function saveJson() {
    try {
      setBusy(true);
      setNotice(null);

      const normalized = normalizeSheetRowsStrings(rows as TableRow[]);
      await ctx.setFieldValue(ctx.fieldPath, { rows: normalized.rows });

      if (params.columnsMetaApiKey) {
        await setFieldByApiOrId(ctx, params.columnsMetaApiKey, { columns: normalized.columns });
      }
      if (params.rowCountApiKey) {
        await setFieldByApiOrId(ctx, params.rowCountApiKey, Number(normalized.rows.length));
      }

      ctx.notice('Saved table JSON to field.');
    } catch (e: any) {
      setNotice(`Save failed: ${e?.message || e}`);
    } finally {
      setBusy(false);
    }
  }

  // Inline edits: update state and immediately push normalized { rows } to the form
  const handleCellValueChanged = (e: any) => {
    const { rowIndex, colDef, newValue } = e;
    const key = colDef.field as string;

    setRows(prev => {
      const next = [...prev];
      const row = { ...(next[rowIndex] || {}) } as Record<string, unknown>;
      row[key] = toStringValue(newValue);
      next[rowIndex] = row;

      const normalized = normalizeSheetRowsStrings(next as TableRow[]);
      void ctx.setFieldValue(ctx.fieldPath, { rows: normalized.rows });

      return next;
    });
  };

  function addRow() {
    setRows(prev => {
      const next = [...prev, {}];
      const normalized = normalizeSheetRowsStrings(next as TableRow[]);
      void ctx.setFieldValue(ctx.fieldPath, { rows: normalized.rows });
      return next;
    });
  }
  function addColumn() {
    const newKey = `column_${Date.now().toString().slice(-4)}`;
    setRows(prev => {
      const next = prev.map(r => ({ ...r, [newKey]: (r as any)[newKey] ?? '' }));
      const normalized = normalizeSheetRowsStrings(next as TableRow[]);
      void ctx.setFieldValue(ctx.fieldPath, { rows: normalized.rows });
      return next;
    });
  }

  // If external changes write back to this field, reflect them
  useEffect(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    if (initial && typeof initial === 'object' && !Array.isArray(initial) && (initial as any).rows) {
      setRows(((initial as any).rows as TableRow[]) || []);
    } else if (Array.isArray(initial)) {
      setRows(initial as TableRow[]);
    }
  }, [ctx.fieldPath, ctx.formValues]);

  return (
    <Canvas ctx={ctx}>
      <div style={{ display: 'flex', gap: 8, marginBottom: 8, flexWrap: 'wrap' }}>
        <Button onClick={importFromSource} disabled={busy} buttonType="primary">
          Import from source file
        </Button>
        <Button onClick={saveJson} disabled={busy} buttonType="primary">
          Save JSON to field
        </Button>
        <Button onClick={addRow} disabled={busy} buttonType="muted" buttonSize="s">
          + Row
        </Button>
        <Button onClick={addColumn} disabled={busy} buttonType="muted" buttonSize="s">
          + Column
        </Button>
      </div>

      {busy && <Spinner />}

      <div style={{ height: 420, width: '100%' }}>
        <AgGridReact
          theme={themeQuartz}
          rowData={rows as any[]}
          columnDefs={columnDefs as any}
          onCellValueChanged={handleCellValueChanged}
        />
      </div>

      {notice && <Alert>{notice}</Alert>}
    </Canvas>
  );
}

// ======= Config screen (CMA token) =======
function Config({ ctx }: { ctx: any }) {
  const [token, setToken] = useState<string>(
    (ctx.plugin.attributes.parameters as any)?.cmaToken || '',
  );
  return (
    <Canvas ctx={ctx}>
      <TextField
        id="cmaToken"
        name="cmaToken"
        label="CMA API Token (Uploads: Read)"
        value={token}
        onChange={setToken}
      />
      <div style={{ marginTop: 8 }}>
        <Button
          buttonType="primary"
          onClick={async () => {
            await ctx.updatePluginParameters({ cmaToken: token });
            ctx.notice('Saved plugin configuration.');
          }}
        >
          Save configuration
        </Button>
      </div>
    </Canvas>
  );
}

// ======= Wire up the plugin =======
connect({
  renderConfigScreen(ctx) {
    ReactDOM.createRoot(document.getElementById('root')!).render(<Config ctx={ctx} />);
  },

  manualFieldExtensions() {
    return [
      {
        id: 'excelJsonEditor',
        name: 'Excel → Editable JSON Table',
        type: 'editor',
        fieldTypes: ['json'],
        parameters: [
          { id: 'sourceFileApiKey', name: 'Source File API key', type: 'string', required: true },
          { id: 'columnsMetaApiKey', name: 'Columns Meta API key', type: 'string' },
          { id: 'rowCountApiKey', name: 'Row Count API key', type: 'string' },
        ],
      },
    ];
  },

  renderFieldExtension(id, ctx) {
    if (id === 'excelJsonEditor') {
      ReactDOM.createRoot(document.getElementById('root')!).render(<Editor ctx={ctx} />);
    }
  },
});

// Optional: dev harness if opened directly (not inside Dato)
if (window.self === window.top) {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <div style={{ padding: 16 }}>
      <h3>Plugin dev harness</h3>
      <p>
        Embed this in Dato as a Private Plugin and attach it as the Field editor for your
        <code> dataJson</code> field (Presentation tab).
      </p>
      <p>
        You can override the file field via URL:
        <code>?fileApiKey=another_file_field</code>
      </p>
    </div>,
  );
}
