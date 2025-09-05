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

type TableRow = Record<string, unknown>;

type FieldParams = {
  /** Sibling file field api key inside the same block (e.g. "sourcefile") */
  sourceFileApiKey?: string;
  /** Optional sibling field (json/text) to store { columns } */
  columnsMetaApiKey?: string;
  /** Optional sibling field (number/text) to store row count */
  rowCountApiKey?: string;
};

/** If your JSON field expects { rows: [...] } instead of { columns, data }, set 'rows' */
const PAYLOAD_SHAPE: 'matrix' | 'rows' = 'matrix';

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
  // Clear first to ensure Dato detects a change
  await ctx.setFieldValue(ctx.fieldPath, null);
  await Promise.resolve();
  await ctx.setFieldValue(ctx.fieldPath, value);
}


/**
 * Find the container object (and its path) in formValues that actually holds this field.
 * We deep-scan formValues and stop at the first object that has our current field id/apiKey as a direct key.
 */
function findBlockContainerWithCurrentField(ctx: RenderFieldExtensionCtx): { container: any; containerPath: string[] } | null {
  const root = (ctx as any).formValues;
  const cur: any = ctx.field;
  const curId = String(cur?.id);
  const curApi = cur?.apiKey ?? cur?.attributes?.api_key;

  function walk(node: any, path: string[]): { container: any; containerPath: string[] } | null {
    if (!node || typeof node !== 'object') return null;

    if (Object.prototype.hasOwnProperty.call(node, curId) || (curApi && Object.prototype.hasOwnProperty.call(node, curApi))) {
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

/** Normalize Dato's upload-like value into a { upload_id } or { __direct_url } */
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
 * Read the sibling file value (by apiKey) from the SAME block as dataJson.
 * Works whether the block stores children by field id or by apiKey.
 */
function getSiblingFileFromBlock(ctx: RenderFieldExtensionCtx, siblingApiKey: string) {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) return null;
  const { container } = hit;

  // Scan actual keys present in this block instance.
  // For each key: resolve its apiKey via ctx.fields when the key is an id; otherwise use the key itself.
  let resolvedKey: string | null = null;
  for (const k of Object.keys(container)) {
    const def: any = (ctx.fields as any)[k];
    const keyApi = def ? (def.apiKey ?? def.attributes?.api_key) : k;
    if (keyApi === siblingApiKey) {
      resolvedKey = k;
      break;
    }
  }
  if (!resolvedKey) return null;

  const localized = pickAnyLocaleValue(container[resolvedKey], ctx.locale);
  return normalizeUploadLike(localized);
}

/**
 * Set a sibling field INSIDE the same block.
 * Tries apiKey first, then resolves by field id if needed, and respects localization.
 */
async function setSiblingInBlock(ctx: RenderFieldExtensionCtx, apiKey: string, value: any) {
  const hit = findBlockContainerWithCurrentField(ctx);
  if (!hit) return;
  const { container, containerPath } = hit;

  // Find the actual key in the container corresponding to this apiKey
  let foundKey: string | null = null;
  let defForTarget: any = null;

  // Try direct apiKey
  if (Object.prototype.hasOwnProperty.call(container, apiKey)) {
    foundKey = apiKey;
    defForTarget = (Object.values(ctx.fields) as any[]).find((f: any) =>
      (f.apiKey ?? f.attributes?.api_key) === apiKey
    );
  } else {
    // Resolve by field id
    const defs = Object.values(ctx.fields) as any[];
    const def = defs.find((f: any) => (f.apiKey ?? f.attributes?.api_key) === apiKey);
    if (def && Object.prototype.hasOwnProperty.call(container, String(def.id))) {
      foundKey = String(def.id);
      defForTarget = def;
    }
  }
  if (!foundKey) return;

  const isLocalized = Boolean(defForTarget?.localized ?? defForTarget?.attributes?.localized);
  const path = [...containerPath, foundKey, ...(isLocalized && ctx.locale ? [ctx.locale] : [])];

  // Write
  await ctx.setFieldValue(path.join('.'), value);
}

/** =================== Robust parsing (headerless-safe) =================== */
function aoaFromWorksheet(ws: XLSX.WorkSheet): any[][] {
  // header:1 -> array-of-arrays, defval:'' -> keep empty cells
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

/** =================== CMA (Uploads) helper =================== */
async function fetchUploadMeta(
  fileFieldValue: any,
  cmaToken: string,
): Promise<{ url: string; mime: string | null; filename: string | null } | null> {
  if (fileFieldValue?.upload_id) {
    if (!cmaToken) return null;
    const client = buildClient({ apiToken: cmaToken });
    const upload: any = await client.uploads.find(String(fileFieldValue.upload_id));
    return {
      url: upload?.url || null,
      mime: upload?.mime_type ?? null,
      filename: upload?.filename ?? null,
    };
  }
  if (fileFieldValue?.__direct_url) {
    const url: string = fileFieldValue.__direct_url;
    let filename: string | null = null;
    try { const u = new URL(url); filename = decodeURIComponent(u.pathname.split('/').pop() || ''); } catch {}
    return { url, mime: null, filename };
  }
  return null;
}

/** =================== UI bits =================== */
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

/** =================== Editor (block-only) =================== */
function Uploader({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = getEditorParams(ctx);
  const sourceApiKey = params.sourceFileApiKey || DEFAULT_SOURCE_FILE_API_KEY;

  console.log(sourceApiKey, 'sourceApiKey')

  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);

  async function importFromBlock() {
    try {
      setBusy(true);
      setNotice(null);

      const fileVal = getSiblingFileFromBlock(ctx, sourceApiKey);
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
              meta: {
                filename: meta.filename ?? null,
                mime: meta.mime ?? null,
                imported_at: new Date().toISOString(),
                nonce: bust,
              },
            }
          : {
              rows: norm.rows,
              meta: {
                filename: meta.filename ?? null,
                mime: meta.mime ?? null,
                imported_at: new Date().toISOString(),
                nonce: bust,
              },
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

      // Persist current item
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

  // we don't render a table preview; nothing to sync visually
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
