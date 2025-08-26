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

// AG Grid v31+ modular registration (Community)
import {
  ModuleRegistry,
  ClientSideRowModelModule,
  TextEditorModule,      // 👈 add this
  ValidationModule,      // (optional) just for clearer console errors
  themeQuartz,
} from 'ag-grid-community';

ModuleRegistry.registerModules([
  ClientSideRowModelModule,
  TextEditorModule,      // 👈 required for editable cells
  ValidationModule,   // (optional) helpful during dev
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

function inferColumns(rows: TableRow[]): string[] {
  const first = rows?.[0] || {};
  return Object.keys(first as object);
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
  const [showDebug, setShowDebug] = useState(false);

  const [sheet, setSheet] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [rows, setRows] = useState<TableRow[]>(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
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
const cleanParsed = sanitizeJSON(parsed);

setSheetNames(names);
setRows(cleanParsed);
setSheet(names[0] || null);


      if (params.columnsMetaApiKey) {
  await setFieldByApiOrId(ctx, params.columnsMetaApiKey, {
    columns: inferColumns(cleanParsed),
  });
}
if (params.rowCountApiKey) {
  await setFieldByApiOrId(ctx, params.rowCountApiKey, Number(cleanParsed.length));
}
    } catch (e: any) {
      setNotice(`Import failed: ${e?.message || e}`);
    } finally {
      setBusy(false);
    }
  }

  function getFieldPath(ctx: RenderFieldExtensionCtx, apiKeyOrId: string): string | null {
  const id = resolveFieldId(ctx, apiKeyOrId);
  if (!id) return null;
  return ctx.locale ? `${id}.${ctx.locale}` : id;
}

// Set a field value by API key or ID safely
async function setFieldByApiOrId(
  ctx: RenderFieldExtensionCtx,
  apiKeyOrId: string,
  value: unknown
) {
  const path = getFieldPath(ctx, apiKeyOrId);
  if (!path) return;
  await ctx.setFieldValue(path, value);
}

// Ensure we never send undefined/NaN (Dato rejects non-JSON values)
function sanitizeJSON(x: any): any {
  if (x === undefined || (typeof x === 'number' && Number.isNaN(x))) return null;
  if (Array.isArray(x)) return x.map(sanitizeJSON);
  if (x && typeof x === 'object') {
    const out: Record<string, any> = {};
    for (const [k, v] of Object.entries(x)) out[k] = sanitizeJSON(v);
    return out;
  }
  return x;
}

  async function saveJson() {
  try {
    setBusy(true);
    setNotice(null);

    const cleanRows = sanitizeJSON(rows);

    // Save into THIS JSON field (ctx.fieldPath already includes locale/id)
    await ctx.setFieldValue(ctx.fieldPath, cleanRows);

    // Optional meta fields — write using field IDs/paths, not API keys directly
    if (params.columnsMetaApiKey) {
      await setFieldByApiOrId(ctx, params.columnsMetaApiKey, {
        columns: inferColumns(cleanRows),
      });
    }
    if (params.rowCountApiKey) {
      await setFieldByApiOrId(ctx, params.rowCountApiKey, Number((cleanRows as any[]).length));
    }

    ctx.notice('Saved table JSON to field.');
  } catch (e: any) {
    setNotice(`Save failed: ${e?.message || e}`);
  } finally {
    setBusy(false);
  }
}


  function addRow() {
    setRows((r) => [...r, {}]);
  }
  function addColumn() {
    const newKey = `col_${Date.now().toString().slice(-4)}`;
    setRows((old) => old.map((r) => ({ ...r, [newKey]: (r as any)[newKey] ?? null })));
  }

  useEffect(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    if (Array.isArray(initial)) setRows(initial as TableRow[]);
  }, [ctx.fieldPath, ctx.formValues]);

  return (
    <Canvas ctx={ctx}>
      {busy && <Spinner />}
      {notice && <Alert>{notice}</Alert>}

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
        <Button
          onClick={() => setShowDebug((v) => !v)}
          disabled={busy}
          buttonType="muted"
          buttonSize="s"
        >
          {showDebug ? 'Hide debug' : 'Show debug'}
        </Button>

        {/* Optional: pick any upload directly (bypasses field resolution) */}
        <Button
          onClick={async () => {
            try {
              setBusy(true);
              setNotice(null);
              const picker = (ctx as any).selectUpload;
              if (!picker) {
                setNotice('Upload picker not available in this SDK version.');
                return;
              }
              const picked = await picker({ multiple: false });
              if (!picked) { setNotice('No file picked.'); return; }

              let url: string | null = (picked.url as string) || (picked.upload?.url as string) || null;
              if (!url) {
                const id =
                  (picked.id as string) ||
                  (picked.upload_id as string) ||
                  (picked.upload?.id as string);
                if (!id) { setNotice('Picked file has no URL or id; cannot resolve.'); return; }
                const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
                if (!token) { setNotice('CMA token required to resolve picked file URL (set it in Configuration).'); return; }
                const client = buildClient({ apiToken: token });
                const upload = await client.uploads.find(String(id));
                url = (upload as any)?.url || null;
              }
              if (!url) { setNotice('Could not resolve a URL for the picked file.'); return; }

              const res = await fetch(url);
              const buf = await res.arrayBuffer();
              const { rows: parsed, sheetNames: names } = toSheetJSRows(buf, sheet || undefined);
              setSheetNames(names);
              setRows(parsed);
              setSheet(names[0] || null);
            } catch (e: any) {
              setNotice(`Debug pick failed: ${e?.message || e}`);
            } finally {
              setBusy(false);
            }
          }}
          disabled={busy}
          buttonType="muted"
          buttonSize="s"
        >
          Pick file (debug)
        </Button>

        {sheetNames.length > 1 && (
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 12, opacity: 0.8 }}>Sheet</span>
            <select
              value={sheet ?? ''}
              onChange={(e) => setSheet(e.target.value)}
              style={{ padding: 6, borderRadius: 6, border: '1px solid var(--border-color)' }}
            >
              {sheetNames.map((n) => (
                <option key={n} value={n}>{n}</option>
              ))}
            </select>
          </label>
        )}
      </div>

      {showDebug && (
        <div style={{ fontSize: 12, opacity: 0.8, marginBottom: 8 }}>
          <strong>Debug</strong>
          <div>Current locale: {String(ctx.locale)}</div>
          <div>
            Params (ctx.parameters or appearance): {JSON.stringify(params)} | URL override:{' '}
            {String(getUrlOverrideApiKey())}
          </div>
          <div>
            Available fields:{' '}
            {JSON.stringify(
              Object.values(ctx.fields).map((f: any) => ({
                id: f.id,
                apiKey: f.apiKey ?? f.attributes?.api_key,
                type: f.fieldType ?? f.attributes?.field_type,
              })),
            )}
          </div>
          <div>
            Resolved file field id: {String(resolveFieldId(ctx, preferredApiKey))}
          </div>
        </div>
      )}

      {/* Theming API: pass theme object, no CSS theme class */}
      <div style={{ height: 420, width: '100%' }}>
        <AgGridReact
          theme={themeQuartz}
          rowData={rows as any[]}
          columnDefs={columnDefs as any}
          onCellValueChanged={(e: any) => {
            const { rowIndex, colDef, newValue } = e;
            const key = colDef.field as string;
            setRows((prev) => {
              const copy = [...prev];
              const row = { ...(copy[rowIndex] || {}) } as Record<string, unknown>;
              row[key] = newValue;
              copy[rowIndex] = row;
              return copy;
            });
          }}
        />
      </div>
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
