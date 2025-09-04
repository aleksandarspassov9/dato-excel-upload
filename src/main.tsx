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

const IMAGE_MIME_PREFIX = 'image/';

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
  return null;
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

async function fetchUploadMeta(
  fileFieldValue: any,
  cmaToken: string,
): Promise<{ url: string; mime_type: string | null; filename: string | null } | null> {
  if (fileFieldValue?.upload_id) {
    const client = buildClient({ apiToken: cmaToken });
    const upload: any = await client.uploads.find(String(fileFieldValue.upload_id));
    return {
      url: upload?.url || null,
      mime_type: upload?.mime_type ?? null,
      filename: upload?.filename ?? null,
    };
  }
  if (fileFieldValue?.__direct_url) {
    // Best-effort for direct URLs (no mime available before fetch)
    const url: string = fileFieldValue.__direct_url;
    const filename = (() => {
      try {
        const u = new URL(url);
        return decodeURIComponent(u.pathname.split('/').pop() || '');
      } catch {
        return null;
      }
    })();
    return { url, mime_type: null, filename };
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
  const [selectedMeta, setSelectedMeta] = useState<{ filename: string | null, mime: string | null } | null>(null);

  function getFileFieldValueStrict() {
    const fileFieldId = resolveFieldId(ctx, preferredApiKey);
    if (!fileFieldId) return null;

    let raw = (ctx.formValues as any)[fileFieldId];
    raw = pickLocalizedValue(raw, ctx.locale);

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

      // 1) read ONLY from the configured file field (no global fallback)
      const fileVal = getFileFieldValueStrict();
      if (!fileVal) {
        setNotice(`No file found in the configured field: "${preferredApiKey}". Check the plugin param or upload a spreadsheet there.`);
        return;
      }

      // 2) resolve URL + mime using CMA when possible
      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      const meta = await fetchUploadMeta(fileVal, token);
      if (!meta?.url) {
        setNotice('Could not resolve upload URL from the file field value.');
        return;
      }
      setSelectedMeta({ filename: meta.filename, mime: meta.mime_type });

      // 3) quick guard: obvious wrong mime types (images)
      if (meta.mime_type && meta.mime_type.startsWith(IMAGE_MIME_PREFIX)) {
        setNotice(`The selected file appears to be an image (${meta.mime_type}${meta.filename ? `, ${meta.filename}` : ''}). Please choose an Excel/CSV file in "${preferredApiKey}".`);
        return;
      }

      // 4) fetch with hard cache-busting
      const bust = Date.now();
      const url = meta.url + (meta.url.includes('?') ? '&' : '?') + `cb=${bust}`;
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);

      // 5) choose parser by Content-Type (fallback to XLSX array)
      const ct = res.headers.get('content-type') || meta.mime_type || '';
      let rowsParsed: TableRow[] = [];
      let sheetList: string[] = [];

      if (ct.includes('csv')) {
        const text = await res.text();
        const wb = XLSX.read(text, { type: 'string' });
        const names = wb.SheetNames;
        const ws = wb.Sheets[names[0]];
        rowsParsed = XLSX.utils.sheet_to_json(ws, { defval: null }) as TableRow[];
        sheetList = names;
      } else {
        // assume XLSX/XLS or similar
        const buf = await res.arrayBuffer();
        try {
          const wb = XLSX.read(buf, { type: 'array' });
          const names = wb.SheetNames;
          const ws = wb.Sheets[names[0]];
          rowsParsed = XLSX.utils.sheet_to_json(ws, { defval: null }) as TableRow[];
          sheetList = names;
        } catch (err: any) {
          // Common SheetJS message when the buffer is a PNG:
          // "PNG Image File is not a spreadsheet"
          throw new Error(
            `${err?.message || err}. Make sure the configured field "${preferredApiKey}" contains an Excel/CSV file, not an image.`,
          );
        }
      }

      const normalized = normalizeSheetRowsStrings(rowsParsed);
      setRows(normalized.rows);
      setColumns(normalized.columns);
      setSheetNames(sheetList);

      const payloadObj = {
        columns: normalized.columns,
        data: normalized.rows.map(r =>
          normalized.columns.map(c => (r as any)[c] ?? '')
        ),
        meta: {
          filename: meta.filename ?? null,
          mime_type: meta.mime_type ?? null,
          imported_at: new Date().toISOString(),
          nonce: bust,
        },
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
        meta: {
          ...(selectedMeta || {}),
          saved_at: new Date().toISOString(),
          nonce: Date.now(),
        },
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
          Import from Excel/CSV
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

      {selectedMeta && (
        <div style={{ marginTop: 8, fontSize: 12, opacity: 0.8 }}>
          Selected: {selectedMeta.filename || 'unknown filename'}
          {selectedMeta.mime ? ` • ${selectedMeta.mime}` : ''}
        </div>
      )}

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
        You must configure the correct file field via the "Source File API key" parameter.
      </p>
    </div>,
  );
}
