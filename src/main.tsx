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

const DEFAULT_SOURCE_FILE_API_KEY = 'sourcefile';

type TableRow = Record<string, unknown>;
type FieldParams = {
  sourceFileApiKey?: string;
  columnsMetaApiKey?: string;
  rowCountApiKey?: string;
};

const IMAGE_MIME_PREFIX = 'image/';

/* ---------- Small helpers ---------- */
function getUrlOverrideApiKey(): string | undefined {
  try {
    const u = new URL(window.location.href);
    return u.searchParams.get('fileApiKey') || undefined;
  } catch { return undefined; }
}
function getEditorParams(ctx: RenderFieldExtensionCtx): FieldParams {
  const direct = (ctx.parameters as any) || {};
  if (direct && Object.keys(direct).length) return direct;
  const appearance =
    (ctx.field as any)?.attributes?.appearance?.parameters ||
    (ctx as any)?.fieldAppearance?.parameters || {};
  return appearance;
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
  const columns = Array.from({ length: colCount }, (_, i) => `column_${i + 1}`);
  const normalizedRows = rows.map((r) => {
    const values = Object.values(r);
    const out: Record<string, string> = {};
    columns.forEach((col, i) => { out[col] = toStringValue(values[i]); });
    return out;
  });
  return { rows: normalizedRows, columns };
}
function pickAnyLocaleValue(raw: any, locale?: string | null) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return raw ?? null;
  if (locale && Object.prototype.hasOwnProperty.call(raw, locale) && raw[locale]) return raw[locale];
  for (const k of Object.keys(raw)) if (raw[k]) return raw[k];
  return null;
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
async function fetchUploadMeta(
  fileFieldValue: any,
  cmaToken: string,
): Promise<{ url: string; mime_type: string | null; filename: string | null } | null> {
  if (fileFieldValue?.upload_id) {
    if (!cmaToken) return null;
    const client = buildClient({ apiToken: cmaToken });
    const upload: any = await client.uploads.find(String(fileFieldValue.upload_id));
    return { url: upload?.url || null, mime_type: upload?.mime_type ?? null, filename: upload?.filename ?? null };
  }
  if (fileFieldValue?.__direct_url) {
    const url: string = fileFieldValue.__direct_url;
    let filename: string | null = null;
    try { const u = new URL(url); filename = decodeURIComponent(u.pathname.split('/').pop() || ''); } catch {}
    return { url, mime_type: null, filename };
  }
  return null;
}

/* ---------- Block container discovery ---------- */
function splitPath(path: string): string[] { return path.split('.').filter(Boolean); }
function parentPath(path: string): string { const segs = splitPath(path); return segs.slice(0, -1).join('.'); }
function getAtPath(root: any, path: string) { return splitPath(path).reduce((acc: any, seg) => (acc ? acc[seg] : undefined), root); }

/** Field info for current (dataJson) field */
function currentFieldInfo(ctx: RenderFieldExtensionCtx) {
  const f: any = ctx.field;
  return { id: String(f?.id), apiKey: f?.apiKey ?? f?.attributes?.api_key };
}

/** Find the nearest object in formValues that contains the current field */
function findBlockContainerWithCurrentField(ctx: RenderFieldExtensionCtx): { container: any; containerPath: string } | null {
  // Fast path: parent of ctx.fieldPath is typically the block container
  const pPath = parentPath(ctx.fieldPath);
  const container = getAtPath((ctx as any).formValues, pPath);
  const { id, apiKey } = currentFieldInfo(ctx);
  if (container && (Object.prototype.hasOwnProperty.call(container, id) || Object.prototype.hasOwnProperty.call(container, apiKey ?? ''))) {
    return { container, containerPath: pPath };
  }

  // Fallback: deep search (rare)
  const root = (ctx as any).formValues;
  function walk(node: any, trail: string[]): { container: any; containerPath: string } | null {
    if (!node || typeof node !== 'object') return null;
    if (Object.prototype.hasOwnProperty.call(node, id) || (apiKey && Object.prototype.hasOwnProperty.call(node, apiKey))) {
      return { container: node, containerPath: trail.join('.') };
    }
    if (Array.isArray(node)) {
      for (let i = 0; i < node.length; i++) { const r = walk(node[i], trail.concat(String(i))); if (r) return r; }
      return null;
    }
    for (const k of Object.keys(node)) { const r = walk(node[k], trail.concat(k)); if (r) return r; }
    return null;
  }
  return walk(root, []) ;
}

/** Find a field definition by apiKey or id */
function findFieldDef(ctx: RenderFieldExtensionCtx, apiKeyOrId: string): any | null {
  const byId = (ctx.fields as any)[apiKeyOrId];
  if (byId) return byId;
  const all = Object.values(ctx.fields) as any[];
  return all.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === apiKeyOrId) ?? null;
}
function isFieldLocalized(def: any): boolean {
  return Boolean(def?.localized ?? def?.attributes?.localized);
}

/** Resolve sibling path (inside the same block) for a given apiKey/id */
function resolveSiblingPathInBlock(ctx: RenderFieldExtensionCtx, apiKeyOrId: string): string | null {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) return null;
  const { container, containerPath } = hit;

  const def = findFieldDef(ctx, apiKeyOrId);
  if (!def) return null;

  const keyCandidates = [String(def.id), def.apiKey ?? def.attributes?.api_key].filter(Boolean);
  const keyInContainer = keyCandidates.find((k) => Object.prototype.hasOwnProperty.call(container, k));
  if (!keyInContainer) return null;

  const locSuffix = isFieldLocalized(def) && ctx.locale ? `.${ctx.locale}` : '';
  return `${containerPath}.${keyInContainer}${locSuffix}`;
}

/** Resolve top-level path (page item) for a given apiKey/id */
function resolveTopLevelPath(ctx: RenderFieldExtensionCtx, apiKeyOrId: string): string | null {
  const def = findFieldDef(ctx, apiKeyOrId);
  if (!def) return null;
  const locSuffix = isFieldLocalized(def) && ctx.locale ? `.${ctx.locale}` : '';
  return `${def.id}${locSuffix}`;
}

/** Sets a field by apiKey/id. Priority: sibling in same block → top level */
async function setFieldByApiOrId(ctx: RenderFieldExtensionCtx, apiKeyOrId: string, value: any) {
  const sib = resolveSiblingPathInBlock(ctx, apiKeyOrId);
  const path = sib ?? resolveTopLevelPath(ctx, apiKeyOrId);
  if (!path) return;
  await ctx.setFieldValue(path, value);
}

/** Get sibling file value from this block by API key */
function getSiblingFileFromBlock(ctx: RenderFieldExtensionCtx, siblingApiKey: string) {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) return null;
  const { container } = hit;

  // try both id and apiKey
  const def = findFieldDef(ctx, siblingApiKey);
  const keyCandidates = def
    ? [String(def.id), def.apiKey ?? def.attributes?.api_key].filter(Boolean)
    : [siblingApiKey];
  const key = keyCandidates.find((k) => Object.prototype.hasOwnProperty.call(container, k));
  if (!key) return null;

  let raw = container[key];
  raw = pickAnyLocaleValue(raw, ctx.locale);
  if (!raw) return null;
  if (Array.isArray(raw)) raw = raw[0];
  if (raw?.upload_id) return raw;
  if (raw?.upload?.id) return { upload_id: raw.upload.id };
  if (typeof raw === 'string' && raw.startsWith('http')) return { __direct_url: raw };
  return null;
}

/** List file siblings inside this block (for assist UI) */
function listBlockFileSiblings(ctx: RenderFieldExtensionCtx) {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) return [] as Array<{ key: string; apiKey: string; hasValue: boolean; preview: string }>;
  const { container } = hit;

  const out: Array<{ key: string; apiKey: string; hasValue: boolean; preview: string }> = [];
  for (const k of Object.keys(container)) {
    const def = (ctx.fields as any)[k] ||
      (Object.values(ctx.fields) as any[]).find((f: any) => (f.apiKey ?? f.attributes?.api_key) === k);
    if (!def) continue;
    const type = def.fieldType ?? def.attributes?.field_type;
    if (type !== 'file') continue;
    const apiKey = def.apiKey ?? def.attributes?.api_key;

    const val = pickAnyLocaleValue(container[k], ctx.locale);
    let hasValue = false; let preview = '';
    if (val) {
      hasValue = true;
      const v0 = Array.isArray(val) ? val[0] : val;
      preview = v0?.upload_id ?? v0?.upload?.id ?? (typeof v0 === 'string' ? v0 : '[object]');
    }
    out.push({ key: k, apiKey, hasValue, preview });
  }
  return out;
}

/** Top-level file fields (assist fallback) */
function listTopLevelFileFields(ctx: RenderFieldExtensionCtx) {
  const all = Object.values(ctx.fields) as any[];
  return all
    .filter((f: any) => (f.fieldType ?? f.attributes?.field_type) === 'file')
    .map((f: any) => ({
      id: String(f.id),
      apiKey: f.apiKey ?? f.attributes?.api_key,
      label: f.label ?? f.attributes?.label,
    }));
}
function getTopLevelFileValueByApiKey(ctx: RenderFieldExtensionCtx, apiKey: string) {
  const list = listTopLevelFileFields(ctx);
  const match = list.find((f) => f.apiKey === apiKey);
  if (!match) return null;
  let raw = (ctx.formValues as any)[match.id];
  raw = pickAnyLocaleValue(raw, ctx.locale);
  if (!raw) return null;
  if (Array.isArray(raw)) raw = raw[0];
  if (raw?.upload_id) return raw;
  if (raw?.upload?.id) return { upload_id: raw.upload.id };
  if (typeof raw === 'string' && raw.startsWith('http')) return { __direct_url: raw };
  return null;
}

function Alert({ children }: { children: React.ReactNode }) {
  return (
    <div role="alert" style={{ padding: '8px 12px', border: '1px solid var(--border-color)', borderRadius: 6, marginTop: 8 }}>
      {children}
    </div>
  );
}

/* ---------- Editor (block-aware) ---------- */
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

  const blockSiblings = useMemo(() => listBlockFileSiblings(ctx), [ctx.fieldPath, ctx.formValues, ctx.locale, ctx.fields]);
  const topLevelFiles = useMemo(() => listTopLevelFileFields(ctx), [ctx.fields]);

  async function importFromSource(opts?: { fromApiKey?: string; preferTopLevel?: boolean }) {
    try {
      setBusy(true);
      setNotice(null);

      const effectiveKey = opts?.fromApiKey ?? preferredApiKey;

      // Prefer the sibling inside the same block; optionally force top-level
      let fileVal = opts?.preferTopLevel
        ? getTopLevelFileValueByApiKey(ctx, effectiveKey)
        : getSiblingFileFromBlock(ctx, effectiveKey) || getTopLevelFileValueByApiKey(ctx, effectiveKey);

      if (!fileVal) {
        const blockList = blockSiblings.length
          ? `Block file fields: ${blockSiblings.map(b => `${b.apiKey}${b.hasValue ? ' (has value)' : ''}`).join(', ')}. `
          : 'No file fields found in this block. ';
        const pageList = topLevelFiles.length
          ? `Page file fields: ${topLevelFiles.map(f => f.apiKey).join(', ')}.`
          : 'No file fields found on the page.';
        setNotice(
          `No file found for API key "${effectiveKey}" in this block or page. ` +
          blockList + pageList +
          ' Upload your spreadsheet to the block’s file field (same block as this dataJson), or use one of the import buttons below.',
        );
        return;
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      const meta = await fetchUploadMeta(fileVal, token);
      if (!meta?.url) {
        setNotice('Could not resolve upload URL. Add a CMA token with "Uploads: read" in the plugin config, or use a direct URL value.');
        return;
      }
      setSelectedMeta({ filename: meta.filename, mime: meta.mime_type });

      if (meta.mime_type && meta.mime_type.startsWith(IMAGE_MIME_PREFIX)) {
        setNotice(`The selected file appears to be an image (${meta.mime_type}${meta.filename ? `, ${meta.filename}` : ''}). Please choose an Excel/CSV file.`);
        return;
      }

      // fetch + parse
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
        data: normalized.rows.map((r) => normalized.columns.map((c) => (r as any)[c] ?? '')),
        meta: {
          filename: meta.filename ?? null,
          mime_type: meta.mime_type ?? null,
          imported_at: new Date().toISOString(),
          nonce: bust,
          source_field_api_key: effectiveKey,
        },
      };

      await writePayload(ctx, payloadObj);

      // Optional meta fields (tries block sibling first, then top-level)
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

  async function saveAndPublish() {
    try {
      setBusy(true);
      setNotice(null);

      const payloadObj = {
        columns,
        data: (rows as Array<Record<string, string>>).map((r) => columns.map((c) => r[c] ?? '')),
        meta: { ...(selectedMeta || {}), saved_at: new Date().toISOString(), nonce: Date.now() },
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

  // reflect external updates
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
    } catch { /* ignore */ }
  }, [ctx.fieldPath, ctx.formValues]);

  return (
    <Canvas ctx={ctx}>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <Button onClick={() => importFromSource()} disabled={busy} buttonType="primary">
          Import from Excel/CSV
        </Button>
        <Button onClick={saveAndPublish} disabled={busy} buttonType="primary">
          Save & Publish
        </Button>
      </div>

      {busy && <Spinner />}

      {/* Assist panel */}
      <div style={{ marginTop: 12, fontSize: 12, opacity: 0.9 }}>
        <div>Configured file field API key for this block: <code>{preferredApiKey}</code></div>
        <div>Current locale: <code>{String(ctx.locale || 'default')}</code></div>

        <div style={{ marginTop: 8 }}>
          <strong>Block file fields (siblings in this block):</strong>
          {blockSiblings.length ? (
            <ul style={{ margin: '6px 0 0 16px' }}>
              {blockSiblings.map((b) => (
                <li key={b.key}>
                  <code>{b.apiKey}</code> — {b.hasValue ? 'has value ✓' : 'empty'}
                  {b.preview ? ` (preview: ${String(b.preview).slice(0, 40)}…)` : ''}
                  {b.hasValue && (
                    <Button
                      buttonSize="xs"
                      buttonType="muted"
                      style={{ marginLeft: 8 }}
                      onClick={() => importFromSource({ fromApiKey: b.apiKey })}
                      disabled={busy}
                    >
                      Import from this block field
                    </Button>
                  )}
                </li>
              ))}
            </ul>
          ) : (
            <div style={{ marginTop: 4 }}>No file fields found inside this block.</div>
          )}
        </div>

        <div style={{ marginTop: 8 }}>
          <strong>Page file fields (top-level fallback):</strong>
          {topLevelFiles.length ? (
            <ul style={{ margin: '6px 0 0 16px' }}>
              {topLevelFiles.map((f) => (
                <li key={f.id}>
                  <code>{f.apiKey}</code>
                  <Button
                    buttonSize="xs"
                    buttonType="muted"
                    style={{ marginLeft: 8 }}
                    onClick={() => importFromSource({ fromApiKey: f.apiKey, preferTopLevel: true })}
                    disabled={busy}
                  >
                    Import from page field
                  </Button>
                </li>
              ))}
            </ul>
          ) : (
            <div style={{ marginTop: 4 }}>No file fields detected on the page item.</div>
          )}
        </div>
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

/* ---------- Config screen ---------- */
function Config({ ctx }: { ctx: any }) {
  const [token, setToken] = useState<string>((ctx.plugin.attributes.parameters as any)?.cmaToken || '');
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
        <Button buttonType="primary" onClick={async () => {
          await ctx.updatePluginParameters({ cmaToken: token });
          ctx.notice('Saved plugin configuration.');
        }}>
          Save configuration
        </Button>
      </div>
    </Canvas>
  );
}

/* ---------- Wire up plugin ---------- */
connect({
  renderConfigScreen(ctx) {
    ReactDOM.createRoot(document.getElementById('root')!).render(<Config ctx={ctx} />);
  },
  manualFieldExtensions() {
    return [{
      id: 'excelJsonUploader',
      name: 'Excel → JSON (Upload & Publish)',
      type: 'editor',
      fieldTypes: ['json', 'text'],
      parameters: [
        { id: 'sourceFileApiKey', name: 'Source File API key (block sibling)', type: 'string', required: true },
        { id: 'columnsMetaApiKey', name: 'Columns Meta API key', type: 'string' },
        { id: 'rowCountApiKey', name: 'Row Count API key', type: 'string' },
      ],
    }];
  },
  renderFieldExtension(id, ctx) {
    if (id === 'excelJsonUploader') {
      ReactDOM.createRoot(document.getElementById('root')!).render(<Uploader ctx={ctx} />);
    }
  },
});

/* ---------- Dev harness (optional) ---------- */
if (window.self === window.top) {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <div style={{ padding: 16 }}>
      <h3>Plugin dev harness</h3>
      <p>Attach this as the Field editor for your block’s <code>dataJson</code> field, and set the parameter <strong>Source File API key</strong> to the block’s file field (e.g. <code>sourcefile</code>).</p>
    </div>,
  );
}
