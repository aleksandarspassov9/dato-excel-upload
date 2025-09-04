// src/main.tsx
import React, { useEffect, useState } from 'react';
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
): { rows: TableRow[]; sheetNames: string[] } {
  const wb = XLSX.read(binary, { type: 'array' });
  const names = wb.SheetNames;
  const ws = wb.Sheets[names[0]];
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

function toStringValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' && Number.isNaN(v)) return '';
  return String(v);
}

/** Normalize to safe column names and **string** cell values */
function normalizeSheetRowsStrings(rows: TableRow[]): { rows: TableRow[]; columns: string[] } {
  if (!rows || rows.length === 0) return { rows: [], columns: [] };

  const firstRow = rows[0] as Record<string, unknown>;
  const colCount = Math.max(1, Object.keys(firstRow).length);
  const safe = Array.from({ length: colCount }, (_, i) => `column_${i + 1}`);

  const normalizedRows = rows.map((r) => {
    const values = Object.values(r);
    const out: Record<string, string> = {};
    safe.forEach((col, i) => {
      out[col] = toStringValue(values[i]);
    });
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

function fieldExpectsJsonObject(ctx: RenderFieldExtensionCtx) {
  return (ctx.field as any)?.attributes?.field_type === 'json';
}

/** Writes RAW JSON correctly depending on the field type:
 *  - JSON field → object
 *  - Text field (or other) → stringified JSON
 */
async function writePayload(ctx: RenderFieldExtensionCtx, payloadObj: any) {
  const value = fieldExpectsJsonObject(ctx) ? payloadObj : JSON.stringify(payloadObj);

  // Clear first to guarantee a diff, then write
  await ctx.setFieldValue(ctx.fieldPath, null);
  await Promise.resolve();
  await ctx.setFieldValue(ctx.fieldPath, value);
}

function Alert({ children }: { children: React.ReactNode }) {
  return (
    <div
      role="alert"
      style={{
        padding: '8px 12px',
        border: '1px solid var(--border-color)',
        borderRadius: 6,
        marginTop: 8,
      }}
    >
      {children}
    </div>
  );
}

// ======= Minimal uploader (no preview/edit) =======
function Uploader({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = getEditorParams(ctx);
  const preferredApiKey =
    getUrlOverrideApiKey() || params.sourceFileApiKey || DEFAULT_SOURCE_FILE_API_KEY;

  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [rows, setRows] = useState<TableRow[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [sheetNames, setSheetNames] = useState<string[]>([]);

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
        setNotice('No file in the configured file field. Upload one and try again.');
        return;
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      let url: string | null = null;

      if ((fileVal as any).__direct_url) {
        url = (fileVal as any).__direct_url as string;
      } else {
        if (!token) {
          setNotice('Missing CMA token in plugin configuration (Uploads: read).');
          return;
        }
        url = await fetchUploadUrlFromValue(fileVal, token);
      }
      if (!url) {
        setNotice('Could not resolve upload URL from the file field value.');
        return;
      }

      // Cache-bust to avoid stale Excel downloads
      const bust = Date.now();
      const res = await fetch(url + (url.includes('?') ? '&' : '?') + `cb=${bust}`, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);
      const buf = await res.arrayBuffer();

      const { rows: parsed, sheetNames: names } = toSheetJSRows(buf);
      const normalized = normalizeSheetRowsStrings(parsed as TableRow[]);

      setRows(normalized.rows);
      setColumns(normalized.columns);
      setSheetNames(names);

      const payloadObj = {
        columns: normalized.columns,
        data: normalized.rows.map(r =>
          normalized.columns.map(c => (r as any)[c] ?? '')
        ),
        meta: { nonce: bust }, // ensures value changes even if content is identical
      };

      await writePayload(ctx, payloadObj);

      // Optional meta fields
      if (params.columnsMetaApiKey) {
        await setFieldByApiOrId(ctx, params.columnsMetaApiKey, { columns: normalized.columns });
      }
      if (params.rowCountApiKey) {
        await setFieldByApiOrId(ctx, params.rowCountApiKey, Number(normalized.rows.length));
      }

      // Persist immediately if possible
      if (typeof (ctx as any).saveCurrentItem === 'function') {
        await (ctx as any).saveCurrentItem();
      }

      ctx.notice('Imported and saved JSON to field.');
    } catch (e: any) {
      setNotice(`Import failed: ${e?.message || e}`);
    } finally {
      setBusy(false);
    }
  }

  async function saveAndPublish() {
    try {
      setBusy(true);
      setNotice(null);

      const payloadObj = {
        columns,
        data: (rows as Array<Record<string, string>>).map(r => columns.map(c => r[c] ?? '')),
        meta: { nonce: Date.now() },
      };

      await writePayload(ctx, payloadObj);

      if (typeof (ctx as any).saveCurrentItem === 'function') {
        await (ctx as any).saveCurrentItem();
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      const itemId = (ctx as any).itemId || (ctx as any).item?.id || null;

      if (token && itemId) {
        const client = buildClient({ apiToken: token });
        await client.items.publish(itemId);
        ctx.notice('Saved & published!');
      } else {
        setNotice('Saved JSON. Click “Publish” in Dato, or add a CMA token with Items: write + publish.');
      }
    } catch (e: any) {
      setNotice(`Save/Publish failed: ${e?.message || e}`);
    } finally {
      setBusy(false);
    }
  }

  // Keep local state in sync if something else writes to this field (optional)
  useEffect(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    try {
      const value = fieldExpectsJsonObject(ctx) ? initial : (initial ? JSON.parse(initial) : null);
      if (value?.columns && value?.data) {
        setColumns(value.columns);
        const asRows: TableRow[] = value.data.map((arr: string[]) =>
          Object.fromEntries(value.columns.map((c: string, i: number) => [c, arr[i] ?? ''])),
        );
        setRows(asRows);
      }
    } catch {
      // ignore invalid JSON in string fields
    }
  }, [ctx.fieldPath, ctx.formValues]);

  return (
    <Canvas ctx={ctx}>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <Button onClick={importFromSource} disabled={busy} buttonType="primary">
          Import from Excel file
        </Button>
        <Button onClick={saveAndPublish} disabled={busy} buttonType="primary">
          Save & Publish
        </Button>
        <Button
          onClick={async () => {
            await ctx.setFieldValue(ctx.fieldPath, null);
            await Promise.resolve();
            await importFromSource();
          }}
          disabled={busy}
          buttonType="muted"
        >
          Force re-import
        </Button>
      </div>

      {busy && <Spinner />}

      {sheetNames.length > 0 && (
        <div style={{ marginTop: 8, fontSize: 12, opacity: 0.8 }}>
          Imported sheets detected: {sheetNames.join(', ')} (first sheet used)
        </div>
      )}

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
        label="CMA API Token (Uploads: read; Items: write + publish for Save & Publish)"
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
        id: 'excelJsonUploader',
        name: 'Excel → JSON (Upload & Publish)',
        type: 'editor',
        fieldTypes: ['json', 'text'], // allow attaching to JSON or Text fields
        parameters: [
          { id: 'sourceFileApiKey', name: 'Source File API key', type: 'string', required: true },
          { id: 'columnsMetaApiKey', name: 'Columns Meta API key', type: 'string' },
          { id: 'rowCountApiKey', name: 'Row Count API key', type: 'string' },
        ],
      },
    ];
  },

  renderFieldExtension(id, ctx) {
    if (id === 'excelJsonUploader') {
      ReactDOM.createRoot(document.getElementById('root')!).render(<Uploader ctx={ctx} />);
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
        <code> dataJson</code> (JSON or Text) field (Presentation tab).
      </p>
      <p>
        You can override the file field via URL:
        <code>?fileApiKey=another_file_field</code>
      </p>
    </div>,
  );
}
