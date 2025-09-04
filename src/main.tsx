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

import 'datocms-react-ui/styles.css';

// ======= Config / fallbacks =======
const DEFAULT_SOURCE_FILE_API_KEY = 'sourcefile'; // change to your usual file field key

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

/** Field IDs that belong to the **current item type** */
function fieldIdsOnCurrentItem(ctx: RenderFieldExtensionCtx): Set<string> {
  // Prefer itemType relationship if available (most accurate)
  const idsFromItemType: string[] =
    (ctx as any)?.itemType?.relationships?.fields?.data?.map((d: any) => String(d.id)) || [];
  if (idsFromItemType.length) return new Set(idsFromItemType);

  // Fallback: derive from formValues keys
  const keys = Object.keys((ctx as any).formValues || {});
  // Keys are typically plain field IDs; if they ever include ".locale", strip it:
  return new Set(keys.map((k) => String(k).split('.')[0]));
}

/** File fields on the **current item type** only */
function listFileFieldsOnItem(ctx: RenderFieldExtensionCtx) {
  const allowed = fieldIdsOnCurrentItem(ctx);
  const all = Object.values(ctx.fields) as any[];
  return all
    .filter((f) => allowed.has(String(f.id)))
    .filter((f) => (f.fieldType ?? f.attributes?.field_type) === 'file')
    .map((f) => ({
      id: String(f.id),
      apiKey: f.apiKey ?? f.attributes?.api_key,
      label: f.label ?? f.attributes?.label,
    }));
}

/** Resolve a field ID for a given apiKey **on the current item type only** */
function resolveFieldIdOnItem(
  ctx: RenderFieldExtensionCtx,
  preferred?: string | null,
): string | null {
  if (!preferred) return null;
  const list = listFileFieldsOnItem(ctx);

  // If the caller passed a numeric/string ID directly and it's on this item, accept it
  const asId = String(preferred);
  if (list.some((f) => f.id === asId)) return asId;

  // Else match by apiKey among current item's fields
  const match = list.find((f) => f.apiKey === preferred);
  return match?.id ?? null;
}

function pickAnyLocaleValue(raw: any, locale?: string | null) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return raw ?? null;
  if (locale && Object.prototype.hasOwnProperty.call(raw, locale) && raw[locale]) {
    return raw[locale];
  }
  for (const k of Object.keys(raw)) {
    if (raw[k]) return raw[k];
  }
  return null;
}

async function fetchUploadMeta(
  fileFieldValue: any,
  cmaToken: string,
): Promise<{ url: string; mime_type: string | null; filename: string | null } | null> {
  if (fileFieldValue?.upload_id) {
    if (!cmaToken) return null;
    const client = buildClient({ apiToken: cmaToken });
    const upload: any = await client.uploads.find(String(fileFieldValue.upload_id));
    return {
      url: upload?.url || null,
      mime_type: upload?.mime_type ?? null,
      filename: upload?.filename ?? null,
    };
  }
  if (fileFieldValue?.__direct_url) {
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

function fieldExpectsJsonObject(ctx: RenderFieldExtensionCtx) {
  return (ctx.field as any)?.attributes?.field_type === 'json';
}

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

// ======= Minimal uploader (with on-item field filtering + assist UI) =======
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

  // File fields **on this item** and whether they currently have a value (any locale)
  const detectedFileFields = useMemo(() => {
    const candidates = listFileFieldsOnItem(ctx);
    return candidates.map((f) => {
      const raw = (ctx.formValues as any)[f.id];
      const val = pickAnyLocaleValue(raw, ctx.locale);
      let hasValue = false;
      let preview = '';
      if (val) {
        hasValue = true;
        if (Array.isArray(val) && val.length > 0) {
          const v0 = val[0];
          preview = v0?.upload_id ?? v0?.upload?.id ?? (typeof v0 === 'string' ? v0 : '[array]');
        } else if (val?.upload_id || val?.upload?.id) {
          preview = val.upload_id ?? val.upload?.id;
        } else if (typeof val === 'string') {
          preview = val;
        } else {
          preview = '[object]';
        }
      }
      return { ...f, hasValue, preview };
    });
  }, [ctx.fields, ctx.formValues, ctx.locale]);

  function getFileFieldValueFrom(apiKey: string | null) {
    if (!apiKey) return null;
    const fileFieldId = resolveFieldIdOnItem(ctx, apiKey);
    if (!fileFieldId) return null;

    let raw = (ctx.formValues as any)[fileFieldId];
    raw = pickAnyLocaleValue(raw, ctx.locale);
    if (!raw) return null;

    if (Array.isArray(raw)) raw = raw[0];
    if (raw?.upload_id) return raw;
    if (raw?.upload?.id) return { upload_id: raw.upload.id };
    if (typeof raw === 'string' && raw.startsWith('http')) return { __direct_url: raw };
    return null;
  }

  async function importFromSource(overrideApiKey?: string) {
    try {
      setBusy(true);
      setNotice(null);

      const effectiveApiKey = overrideApiKey ?? preferredApiKey;
      const fileVal = getFileFieldValueFrom(effectiveApiKey);

      if (!fileVal) {
        const onItem = listFileFieldsOnItem(ctx);
        const extras = onItem.length
          ? `Detected file fields on this item: ${onItem
              .map((f) => `${f.apiKey}${detectedFileFields.find(df => df.id === f.id)?.hasValue ? ' (has value)' : ''}`)
              .join(', ')}.`
          : 'No file fields detected on this item.';
        setNotice(
          `No file found in the field "${effectiveApiKey}". ${extras} ` +
          'Upload your spreadsheet to that exact field (and locale), or click "Import from this field" next to a field that has a value.',
        );
        return;
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      const meta = await fetchUploadMeta(fileVal, token);
      if (!meta?.url) {
        setNotice('Could not resolve upload URL from the file field value. Add a CMA token with "Uploads: read" in the plugin config, or use a direct URL value.');
        return;
      }
      setSelectedMeta({ filename: meta.filename, mime: meta.mime_type });

      if (meta.mime_type && meta.mime_type.startsWith(IMAGE_MIME_PREFIX)) {
        setNotice(`The selected file appears to be an image (${meta.mime_type}${meta.filename ? `, ${meta.filename}` : ''}). Please choose an Excel/CSV file.`);
        return;
      }

      const bust = Date.now();
      const url = meta.url + (meta.url.includes('?') ? '&' : '?') + `cb=${bust}`;
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);

      const ct = res.headers.get('content-type') || meta.mime_type || '';
      let rowsParsed: TableRow[] = [];
      let names: string[] = [];

      if (ct.includes('csv')) {
        const text = await res.text();
        const wb = XLSX.read(text, { type: 'string' });
        names = wb.SheetNames;
        const ws = wb.Sheets[names[0]];
        rowsParsed = XLSX.utils.sheet_to_json(ws, { defval: null }) as TableRow[];
      } else {
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        names = wb.SheetNames;
        const ws = wb.Sheets[names[0]];
        rowsParsed = XLSX.utils.sheet_to_json(ws, { defval: null }) as TableRow[];
      }

      const normalized = normalizeSheetRowsStrings(rowsParsed);
      setRows(normalized.rows);
      setColumns(normalized.columns);
      setSheetNames(names);

      const payloadObj = {
        columns: normalized.columns,
        data: normalized.rows.map((r) =>
          normalized.columns.map((c) => (r as any)[c] ?? ''),
        ),
        meta: {
          filename: meta.filename ?? null,
          mime_type: meta.mime_type ?? null,
          imported_at: new Date().toISOString(),
          nonce: bust,
          source_field_api_key: effectiveApiKey,
        },
      };

      await writePayload(ctx, payloadObj);

      if (params.columnsMetaApiKey) {
        await setFieldByApiOrId(ctx, params.columnsMetaApiKey, { columns: normalized.columns });
      }
      if (params.rowCountApiKey) {
        await setFieldByApiOrId(ctx, params.rowCountApiKey, Number(normalized.rows.length));
      }

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

  async function setFieldByApiOrId(ctx: RenderFieldExtensionCtx, apiKeyOrId: string, value: unknown) { const path = getFieldPath(ctx, apiKeyOrId); if (path) await ctx.setFieldValue(path, value); }
  
  function getFieldPath(ctx: RenderFieldExtensionCtx, apiKeyOrId: string): string | null {
    const fields = Object.values(ctx.fields) as any[];
    const byId = (ctx.fields as any)[apiKeyOrId];
    const id =
      byId?.id ?? (fields.find(f => (f.apiKey ?? f.attributes?.api_key) === apiKeyOrId)?.id);
    if (!id) return null;
    return ctx.locale ? `${id}.${ctx.locale}` : id;
  }

  async function saveAndPublish() {
    try {
      setBusy(true);
      setNotice(null);

      const payloadObj = {
        columns,
        data: (rows as Array<Record<string, string>>).map((r) => columns.map((c) => r[c] ?? '')),
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

  const onItemFileFields = listFileFieldsOnItem(ctx);
  const configuredFieldId = resolveFieldIdOnItem(ctx, preferredApiKey);
  const configuredFieldExists = !!configuredFieldId;

  return (
    <Canvas ctx={ctx}>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <Button onClick={() => importFromSource()} disabled={busy} buttonType="primary">
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

      {/* Assist panel: shows only file fields that belong to this item */}
      <div style={{ marginTop: 10, fontSize: 12, opacity: 0.9 }}>
        <div>
          Configured file field API key: <code>{preferredApiKey}</code>{' '}
          {!configuredFieldExists && <strong>(not found on this item type)</strong>}
        </div>
        <div>Current locale: <code>{String(ctx.locale || 'default')}</code></div>
        {onItemFileFields.length > 0 ? (
          <div style={{ marginTop: 6 }}>
            File fields on this item:
            <ul style={{ margin: '6px 0 0 16px' }}>
              {detectedFileFields.map((f) => (
                <li key={f.id}>
                  <code>{f.apiKey}</code> — {f.hasValue ? 'has value ✓' : 'empty'}
                  {f.preview ? ` (preview: ${String(f.preview).slice(0, 40)}…)` : ''}
                  {f.hasValue && (
                    <Button
                      buttonSize="xs"
                      buttonType="muted"
                      style={{ marginLeft: 8 }}
                      onClick={() => importFromSource(f.apiKey)}
                      disabled={busy}
                    >
                      Import from this field
                    </Button>
                  )}
                </li>
              ))}
            </ul>
          </div>
        ) : (
          <div style={{ marginTop: 6 }}>
            No file fields detected on this item. Add a single-asset file field to this model and set its API key in the editor parameters.
          </div>
        )}
      </div>

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
        fieldTypes: ['json', 'text'],
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
        Attach this as the Field editor for your <code>dataJson</code> (JSON or Text) field.
      </p>
      <p>
        Configure the correct file field via the "Source File API key" parameter (or add <code>?fileApiKey=…</code>).
      </p>
    </div>,
  );
}
