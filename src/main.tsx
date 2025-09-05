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
const PAYLOAD_SHAPE: 'matrix' | 'rows' = 'matrix'; // set to 'rows' if your JSON field expects { rows: [...] }
const DEBUG = false; // set to true to see logs

type TableRow = Record<string, string>;
type FieldParams = {
  sourceFileApiKey?: string;
  columnsMetaApiKey?: string;
  rowCountApiKey?: string;
};

const log = (...a: any[]) => { if (DEBUG) console.log('[excel-json-block]', ...a); };

/** =================== Small utils =================== */
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

async function writePayload(ctx: RenderFieldExtensionCtx, payloadObj: any) {
  const value = JSON.stringify(payloadObj);
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

/** =================== Find current block container =================== */
function findBlockContainerWithCurrentField(
  ctx: RenderFieldExtensionCtx,
): { container: any; containerPath: string[] } | null {
  const root = (ctx as any).formValues;
  const cur: any = ctx.field;
  const curId = String(cur?.id);
  const curApi = cur?.apiKey ?? cur?.attributes?.api_key;

  function walk(node: any, path: string[]): { container: any; containerPath: string[] } | null {
    if (!node || typeof node !== 'object') return null;

    if (Object.prototype.hasOwnProperty.call(node, curId) ||
        (curApi && Object.prototype.hasOwnProperty.call(node, curApi))) {
      return { container: node, containerPath: path };
    }

    if (Array.isArray(node)) {
      for (let i = 0; i < node.length; i++) {
        const r = walk(node[i], path.concat(String(i)));
        if (r) return r;
      }
      return null;
    }

    for (const k of Object.keys(node)) {
      const r = walk(node[k], path.concat(k));
      if (r) return r;
    }
    return null;
  }

  return walk(root, []);
}

/** =================== Upload value helpers =================== */
function normalizeUploadLike(raw: any) {
  if (!raw) return null;
  const v = Array.isArray(raw) ? raw[0] : raw;
  if (!v) return null;
  if (v?.upload_id) return v;
  if (v?.upload?.id) return { upload_id: v.upload.id };
  if (typeof v === 'string' && v.startsWith('http')) return { __direct_url: v };
  return null;
}

/** Deeply search an object for any upload-like structure */
function findFirstUploadDeep(val: any): any | null {
  if (!val) return null;
  const candidate = normalizeUploadLike(val);
  if (candidate) return candidate;

  if (Array.isArray(val)) {
    for (const it of val) {
      const n = findFirstUploadDeep(it);
      if (n) return n;
    }
    return null;
  }
  if (typeof val === 'object') {
    for (const k of Object.keys(val)) {
      const n = findFirstUploadDeep(val[k]);
      if (n) return n;
    }
  }
  return null;
}

/** =================== Sibling file lookup (never leaves this block) =================== */
function getSiblingFileFromBlock(ctx: RenderFieldExtensionCtx, siblingApiKey: string) {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) { log('no block container found'); return null; }
  const { container } = hit;

  // Field definition by API key (if available)
  const allDefs = Object.values(ctx.fields) as any[];
  const sibDef = allDefs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === siblingApiKey);
  const isLocalized = Boolean(sibDef?.localized ?? sibDef?.attributes?.localized);

  // 1) Try by ID key (typical for blocks)
  if (sibDef?.id && Object.prototype.hasOwnProperty.call(container, String(sibDef.id))) {
    const raw = isLocalized ? pickAnyLocaleValue(container[String(sibDef.id)], ctx.locale)
                            : container[String(sibDef.id)];
    const norm = normalizeUploadLike(raw) || findFirstUploadDeep(raw);
    log('try by id', sibDef.id, '→', !!norm);
    if (norm) return norm;
  }

  // 2) Try by API key stored directly
  if (Object.prototype.hasOwnProperty.call(container, siblingApiKey)) {
    const raw = isLocalized ? pickAnyLocaleValue(container[siblingApiKey], ctx.locale)
                            : container[siblingApiKey];
    const norm = normalizeUploadLike(raw) || findFirstUploadDeep(raw);
    log('try by apiKey', siblingApiKey, '→', !!norm);
    if (norm) return norm;
  }

  // 3) Safety net: scan this container’s keys and pick an upload-like value.
  // Prefer a key whose apiKey maps to siblingApiKey; otherwise first upload-like.
  let fallback: any | null = null;
  for (const k of Object.keys(container)) {
    const defById = (ctx.fields as any)[k] || allDefs.find((f: any) => String(f.id) === String(k));
    const keyApi = defById ? (defById.apiKey ?? defById.attributes?.api_key) : k;
    const val = pickAnyLocaleValue(container[k], ctx.locale);
    const norm = normalizeUploadLike(val) || findFirstUploadDeep(val);
    if (!norm) continue;

    if (keyApi === siblingApiKey) {
      log('fallback scan matched by apiKey', k);
      return norm;
    }
    if (!fallback) fallback = norm;
  }
  if (fallback) {
    log('fallback scan returning first upload-like value');
    return fallback;
  }

  log('no upload-like value found in container');
  return null;
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
      log('resolved sibling file →', fileVal);
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

  /**
 * Set a sibling field inside the SAME block.
 * It builds the path using the sibling field's **id** (preferred) and appends the locale segment when needed.
 */
async function setSiblingInBlock(
  ctx: RenderFieldExtensionCtx,
  apiKey: string,
  value: any,
) {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) return;
  const { containerPath } = hit;

  // Find sibling field definition by apiKey
  const allDefs = Object.values(ctx.fields) as any[];
  const def = allDefs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === apiKey);

  // Prefer id-based path; fall back to apiKey if we can't resolve the def
  const key = def?.id ? String(def.id) : apiKey;

  // Respect localization if the target field is localized
  const isLocalized = Boolean(def?.localized ?? def?.attributes?.localized);
  const path = [...containerPath, key, ...(isLocalized && ctx.locale ? [ctx.locale] : [])].join('.');

  await ctx.setFieldValue(path, value);
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
