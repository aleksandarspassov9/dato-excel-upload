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
import 'datocms-react-ui/styles.css';

/** CONFIG */
const DEFAULT_SOURCE_FILE_API_KEY = 'sourcefile';
type TableRow = Record<string, unknown>;
type FieldParams = {
  sourceFileApiKey?: string;       // sibling file field api key in this block (e.g. "sourcefile")
  columnsMetaApiKey?: string;      // optional sibling json/text to write { columns }
  rowCountApiKey?: string;         // optional sibling number/text to write row count
};
// If your JSON field expects { rows: [...] } instead of { columns, data }, set 'rows'
const PAYLOAD_SHAPE: 'matrix' | 'rows' = 'matrix';

/** ---------- Small utils ---------- */
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
  // clear first to guarantee a diff
  await ctx.setFieldValue(ctx.fieldPath, null);
  await Promise.resolve();
  await ctx.setFieldValue(ctx.fieldPath, value);
}

/** ---------- Block-only helpers ---------- */
function splitPath(p: string) { return p.split('.').filter(Boolean); }
function parentPath(p: string) { const s = splitPath(p); return s.slice(0, -1).join('.'); }
function getAtPath(root: any, path: string) {
  return splitPath(path).reduce((acc: any, seg) => (acc ? acc[seg] : undefined), root);
}
/** Get the container object for this block (the object that contains dataJson) */
function getBlockContainer(ctx: RenderFieldExtensionCtx) {
  const p = parentPath(ctx.fieldPath);
  const container = getAtPath((ctx as any).formValues, p);
  return { container, containerPath: p };
}
/** Read sibling file by apiKey inside this block only */
function normalizeUploadLike(raw: any) {
  if (!raw) return null;
  if (Array.isArray(raw)) raw = raw[0];
  if (!raw) return null;
  if (raw?.upload_id) return raw;
  if (raw?.upload?.id) return { upload_id: raw.upload.id };
  if (typeof raw === 'string' && raw.startsWith('http')) return { __direct_url: raw };
  return null;
}
function getSiblingFileFromBlock(ctx: RenderFieldExtensionCtx, apiKey: string) {
  const { container } = getBlockContainer(ctx);
  console.log(container, 'container')
  if (!container || typeof container !== 'object') return null;

  // Most blocks keep child values keyed by apiKey; try that first
  if (Object.prototype.hasOwnProperty.call(container, apiKey)) {
    const raw = pickAnyLocaleValue(container[apiKey], ctx.locale);
    const norm = normalizeUploadLike(raw);
    if (norm) return norm;
  }

  // If the block stores by field id, try to resolve by definition id
  const allDefs = Object.values(ctx.fields) as any[];
  const def = allDefs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === apiKey);
  if (def && Object.prototype.hasOwnProperty.call(container, String(def.id))) {
    const raw = pickAnyLocaleValue(container[String(def.id)], ctx.locale);
    const norm = normalizeUploadLike(raw);
    if (norm) return norm;
  }

  return null;
}
/** Set a sibling field (columns/rowCount) inside this block only */
async function setSiblingInBlock(ctx: RenderFieldExtensionCtx, apiKey: string, value: any) {
  const { container, containerPath } = getBlockContainer(ctx);
  if (!container || typeof container !== 'object') return;

  let key: string | null = null;
  if (Object.prototype.hasOwnProperty.call(container, apiKey)) key = apiKey;
  else {
    const defs = Object.values(ctx.fields) as any[];
    const def = defs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === apiKey);
    if (def && Object.prototype.hasOwnProperty.call(container, String(def.id))) key = String(def.id);
  }
  if (!key) return;

  // Check localization for target
  const defs = Object.values(ctx.fields) as any[];
  const def = defs.find((f: any) => String(f.id) === key || (f.apiKey ?? f.attributes?.api_key) === key);
  const isLocalized = Boolean(def?.localized ?? def?.attributes?.localized);
  const locSuffix = isLocalized && ctx.locale ? `.${ctx.locale}` : '';
  const path = `${containerPath}.${key}${locSuffix}`;
  await ctx.setFieldValue(path, value);
}

/** ---------- Robust parsing (works even without headers) ---------- */
function aoaFromWorksheet(ws: XLSX.WorkSheet): any[][] {
  // header:1 -> array-of-arrays, defval:'' -> keep empties
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
  // drop completely blank rows
  return aoa.filter(row => row.some(cell => String(cell ?? '').trim() !== ''));
}
function normalizeAoA(aoa: any[][]) {
  const maxCols = Math.max(1, ...aoa.map(r => r.length));
  const columns = Array.from({ length: maxCols }, (_, i) => `column_${i + 1}`);
  const rows: TableRow[] = aoa.map(r => {
    const padded = [...r];
    while (padded.length < maxCols) padded.push('');
    const obj: Record<string, string> = {};
    columns.forEach((c, i) => { obj[c] = toStringValue(padded[i]); });
    return obj;
  });
  return { rows, columns };
}

/** ---------- CMA helper ---------- */
async function fetchUploadMeta(
  fileFieldValue: any,
  cmaToken: string,
): Promise<{ url: string; mime: string | null; filename: string | null } | null> {
  if (fileFieldValue?.upload_id) {
    if (!cmaToken) return null;
    const client = buildClient({ apiToken: cmaToken });
    const upload: any = await client.uploads.find(String(fileFieldValue.upload_id));
    return { url: upload?.url || null, mime: upload?.mime_type ?? null, filename: upload?.filename ?? null };
  }
  if (fileFieldValue?.__direct_url) {
    const url: string = fileFieldValue.__direct_url;
    let filename: string | null = null;
    try { const u = new URL(url); filename = decodeURIComponent(u.pathname.split('/').pop() || ''); } catch {}
    return { url, mime: null, filename };
  }
  return null;
}

/** UI bits */
function Alert({ children }: { children: React.ReactNode }) {
  return (
    <div role="alert" style={{ padding: '8px 12px', border: '1px solid var(--border-color)', borderRadius: 6, marginTop: 8 }}>
      {children}
    </div>
  );
}

/** ---------- Editor (block-only) ---------- */
function Uploader({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = getEditorParams(ctx);
  const sourceApiKey = DEFAULT_SOURCE_FILE_API_KEY;

  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);

  async function importFromBlock() {
    try {
      setBusy(true);
      setNotice(null);

      const fileVal = getSiblingFileFromBlock(ctx, sourceApiKey);
      console.log(fileVal, 'fileVal')
      if (!fileVal) {
        setNotice(`No file found in this block’s "${sourceApiKey}" field. Upload an .xlsx/.xls/.csv there and try again.`);
        return;
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      const meta = await fetchUploadMeta(fileVal, token);
      if (!meta?.url) { setNotice('Could not resolve upload URL. Add a CMA token with "Uploads: read" in the plugin configuration.'); return; }
      if (meta.mime && meta.mime.startsWith('image/')) { setNotice(`"${meta.filename ?? 'selected file'}" looks like an image (${meta.mime}). Please upload an Excel/CSV file.`); return; }

      // Fetch with cache-busting
      const bust = Date.now();
      const res = await fetch(meta.url + (meta.url.includes('?') ? '&' : '?') + `cb=${bust}`, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);

      // Parse as AoA (works with/without headers), then normalize
      const ct = res.headers.get('content-type') || meta.mime || '';
      let aoa: any[][];
      if (ct.includes('csv')) {
        const text = await res.text();
        const wb = XLSX.read(text, { type: 'string' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        aoa = aoaFromWorksheet(ws);
      } else {
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        aoa = aoaFromWorksheet(ws);
      }

      const norm = normalizeAoA(aoa);

      const payloadObj =
        PAYLOAD_SHAPE === 'matrix'
          ? {
              columns: norm.columns,
              data: norm.rows.map((r) => norm.columns.map((c) => (r as any)[c] ?? '')),
              meta: { filename: meta.filename ?? null, mime: meta.mime ?? null, imported_at: new Date().toISOString(), nonce: bust },
            }
          : {
              rows: norm.rows,
              meta: { filename: meta.filename ?? null, mime: meta.mime ?? null, imported_at: new Date().toISOString(), nonce: bust },
            };

      await writePayload(ctx, payloadObj);

      // Optional: write sibling meta fields inside this same block
      if (params.columnsMetaApiKey && PAYLOAD_SHAPE === 'matrix') {
        await setSiblingInBlock(ctx, params.columnsMetaApiKey, { columns: norm.columns });
      }
      if (params.rowCountApiKey) {
        await setSiblingInBlock(ctx, params.rowCountApiKey, Number(norm.rows.length));
      }

      if (typeof (ctx as any).saveCurrentItem === 'function') {
        await (ctx as any).saveCurrentItem();
      }

      ctx.notice(`Imported ${norm.rows.length} rows × ${norm.columns.length} columns.`);
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

      // Just persist/publish the current item (no rewrite)
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

  // we don't render a table, so nothing to sync visually
  useEffect(() => {}, [ctx.fieldPath, ctx.formValues]);

  return (
    <Canvas ctx={ctx}>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <Button onClick={importFromBlock} disabled={busy} buttonType="primary">
          Import from Excel/CSV (block)
        </Button>
        <Button onClick={saveAndPublish} disabled={busy} buttonType="primary">
          Save & Publish
        </Button>
      </div>
      {busy && <Spinner />}
      {notice && <Alert>{notice}</Alert>}
    </Canvas>
  );
}

/** ---------- Config screen ---------- */
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

/** ---------- Wire up plugin ---------- */
connect({
  renderConfigScreen(ctx) {
    ReactDOM.createRoot(document.getElementById('root')!).render(<Config ctx={ctx} />);
  },
  manualFieldExtensions() {
    return [{
      id: 'excelJsonUploaderBlockOnly',
      name: 'Excel → JSON (Block Only)',
      type: 'editor',
      fieldTypes: ['json', 'text'],
      parameters: [
        { id: 'sourceFileApiKey', name: 'Sibling file field API key', type: 'string', required: true, help_text: 'Usually "sourcefile".' },
        { id: 'columnsMetaApiKey', name: 'Sibling meta field for columns (optional)', type: 'string' },
        { id: 'rowCountApiKey', name: 'Sibling meta field for row count (optional)', type: 'string' },
      ],
    }];
  },
  renderFieldExtension(id, ctx) {
    if (id === 'excelJsonUploaderBlockOnly') {
      ReactDOM.createRoot(document.getElementById('root')!).render(<Uploader ctx={ctx} />);
    }
  },
});

/** ---------- Dev harness (optional) ---------- */
if (window.self === window.top) {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <div style={{ padding: 16 }}>
      <h3>Plugin dev harness</h3>
      <p>Attach as editor to the block’s <code>dataJson</code> field. The importer reads from the sibling <code>sourcefile</code> field in the same block.</p>
    </div>,
  );
}
