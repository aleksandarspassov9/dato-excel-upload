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

/** =================== Config =================== */
const DEFAULT_SOURCE_FILE_API_KEY = 'sourcefile';
const PAYLOAD_SHAPE: 'matrix' | 'rows' = 'matrix'; // set 'rows' if your field expects { rows: [...] }
const DEBUG = false; // set true to see console logs

type TableRow = Record<string, string>;
type FieldParams = {
  sourceFileApiKey?: string;   // e.g. "sourcefile"
  columnsMetaApiKey?: string;  // optional sibling json/text to write { columns }
  rowCountApiKey?: string;     // optional sibling number/text to write row count
};

/** =================== Tiny utils =================== */
const log = (...args: any[]) => { if (DEBUG) console.log('[excel-json-block]', ...args); };

function getEditorParams(ctx: RenderFieldExtensionCtx): FieldParams {
  const direct = (ctx.parameters as any) || {};
  if (direct && Object.keys(direct).length) return direct;
  const appearance =
    (ctx.field as any)?.attributes?.appearance?.parameters ||
    (ctx as any)?.fieldAppearance?.parameters ||
    {};
  return appearance;
}
function toStringValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' && Number.isNaN(v)) return '';
  return String(v);
}
function fieldExpectsJsonObject(ctx: RenderFieldExtensionCtx) {
  return (ctx.field as any)?.attributes?.field_type === 'json';
}
async function writePayload(ctx: RenderFieldExtensionCtx, payloadObj: any) {
  const value = fieldExpectsJsonObject(ctx) ? payloadObj : JSON.stringify(payloadObj);
  // Clear first to force a diff
  await ctx.setFieldValue(ctx.fieldPath, null);
  await Promise.resolve();
  await ctx.setFieldValue(ctx.fieldPath, value);
}
function pickAnyLocaleValue(raw: any, locale?: string | null) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return raw ?? null;
  if (locale && Object.prototype.hasOwnProperty.call(raw, locale) && raw[locale]) return raw[locale];
  for (const k of Object.keys(raw)) if (raw[k]) return raw[k];
  return null;
}

/** =================== Path helpers (block-only) =================== */
/**
 * Dissect the current fieldPath into: containerPath (array of segments),
 * the key of the current field (by id), and whether a trailing locale segment was present.
 */
function dissectCurrentPath(ctx: RenderFieldExtensionCtx) {
  const parts = String(ctx.fieldPath).split('.').filter(Boolean);
  let localeWasSuffix = false;
  if (ctx.locale && parts[parts.length - 1] === String(ctx.locale)) {
    localeWasSuffix = true;
    parts.pop(); // drop locale segment
  }
  const currentKey = parts.pop() || '';
  const containerPath = parts; // array of segments
  log('dissect', { fieldPath: ctx.fieldPath, containerPath, currentKey, localeWasSuffix });
  return { containerPath, currentKey, localeWasSuffix };
}

/** Read a value from formValues via dotted path (supports arrays) */
function getValueAt(root: any, dotted: string) {
  const parts = dotted.split('.').filter(Boolean);
  return parts.reduce((acc: any, seg) => (acc == null ? acc : acc[seg]), root);
}

/** Convert any upload-like to a normalized shape we can use */
function normalizeUploadLike(raw: any) {
  if (!raw) return null;
  if (Array.isArray(raw)) raw = raw[0];
  if (!raw) return null;
  if (raw?.upload_id) return raw;
  if (raw?.upload?.id) return { upload_id: raw.upload.id };
  if (typeof raw === 'string' && raw.startsWith('http')) return { __direct_url: raw };
  return null;
}

/**
 * Resolve the sibling file value by building a path using the sibling field **id** (preferred) and
 * optionally the apiKey as a fallback. This is the key change that avoids returning null.
 */
function getSiblingFileFromBlock(ctx: RenderFieldExtensionCtx, siblingApiKey: string) {
  const { containerPath } = dissectCurrentPath(ctx);
  const root = (ctx as any).formValues;

  // Find sibling field definition by apiKey (works both for block fields and globals)
  const allDefs = Object.values(ctx.fields) as any[];
  const siblingDef = allDefs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === siblingApiKey);
  const isLocalized = Boolean(siblingDef?.localized ?? siblingDef?.attributes?.localized);
  const locSuffix = isLocalized && ctx.locale ? `.${ctx.locale}` : '';

  // Prefer the ID path — blocks typically store by field id
  if (siblingDef?.id) {
    const idPath = [...containerPath, String(siblingDef.id)].join('.') + locSuffix;
    const rawId = pickAnyLocaleValue(getValueAt(root, idPath), ctx.locale);
    const normId = normalizeUploadLike(rawId);
    log('try idPath', idPath, '→', normId ? 'found' : 'miss');
    if (normId) return normId;
  }

  // Fallback: try apiKey path (some block editors keep apiKey)
  const ak = siblingDef?.apiKey ?? siblingDef?.attributes?.api_key ?? siblingApiKey;
  const akPath = [...containerPath, ak].join('.') + locSuffix;
  const rawAk = pickAnyLocaleValue(getValueAt(root, akPath), ctx.locale);
  const normAk = normalizeUploadLike(rawAk);
  log('try akPath', akPath, '→', normAk ? 'found' : 'miss');
  if (normAk) return normAk;

  // Nothing found
  return null;
}

/**
 * Set a sibling field inside the same block using its field **id** path (with locale when needed).
 * This avoids ambiguity and works even when the container stores children by id.
 */
async function setSiblingInBlock(ctx: RenderFieldExtensionCtx, apiKey: string, value: any) {
  const { containerPath } = dissectCurrentPath(ctx);
  const allDefs = Object.values(ctx.fields) as any[];
  const def = allDefs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === apiKey);
  if (!def?.id) return;

  const isLocalized = Boolean(def.localized ?? def.attributes?.localized);
  const locSuffix = isLocalized && ctx.locale ? `.${ctx.locale}` : '';
  const targetPath = [...containerPath, String(def.id)].join('.') + locSuffix;

  log('setSiblingInBlock', { apiKey, targetPath, isLocalized });
  await ctx.setFieldValue(targetPath, value);
}

/** =================== Robust parsing (headerless-safe) =================== */
function aoaFromWorksheet(ws: XLSX.WorkSheet): any[][] {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
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

/** =================== CMA (Uploads) helper =================== */
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

/** =================== UI =================== */
function Alert({ children }: { children: React.ReactNode }) {
  return (
    <div role="alert" style={{ padding: '8px 12px', border: '1px solid var(--border-color)', borderRadius: 6, marginTop: 8 }}>
      {children}
    </div>
  );
}

/** =================== Editor (block-only) =================== */
function Uploader({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = getEditorParams(ctx);
  const sourceApiKey = params.sourceFileApiKey || DEFAULT_SOURCE_FILE_API_KEY;

  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);

  async function importFromBlock() {
    try {
      setBusy(true);
      setNotice(null);

      const fileVal = getSiblingFileFromBlock(ctx, sourceApiKey);
      log('resolved sibling file', fileVal);
      if (!fileVal) {
        setNotice(`No file found in this block’s "${sourceApiKey}" field. Upload an .xlsx/.xls/.csv there and try again.`);
        return;
      }

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      const meta = await fetchUploadMeta(fileVal, token);
      if (!meta?.url) {
        setNotice('Could not resolve upload URL. Add a CMA token with "Uploads: read" in the plugin configuration.');
        return;
      }
      if (meta.mime && meta.mime.startsWith('image/')) {
        setNotice(`"${meta.filename ?? 'selected file'}" looks like an image (${meta.mime}). Please upload an Excel/CSV file.`);
        return;
      }

      // Fetch with cache-busting
      const bust = Date.now();
      const res = await fetch(meta.url + (meta.url.includes('?') ? '&' : '?') + `cb=${bust}`, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);

      // Parse → normalize
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

      // Optional sibling meta inside this block
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

/** =================== Config screen =================== */
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

/** =================== Wire up plugin =================== */
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

/** =================== Dev harness (optional) =================== */
if (window.self === window.top) {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <div style={{ padding: 16 }}>
      <h3>Plugin dev harness</h3>
      <p>Attach this as the editor for the block’s <code>dataJson</code> field. The importer reads from the sibling <code>sourcefile</code> field in the same block.</p>
    </div>,
  );
}
