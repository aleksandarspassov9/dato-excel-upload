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
import { ModuleRegistry, ClientSideRowModelModule } from 'ag-grid-community';
ModuleRegistry.registerModules([ClientSideRowModelModule]);

// Legacy CSS themes (do NOT pass a theme object prop to AgGridReact)
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';

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

// Prefer ctx.parameters, else field appearance, but we can work without either
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
    // treat as field ID
    const byId = (ctx.fields as any)[preferred];
    if (byId?.id) return byId.id;

    // or as API key
    const match = fields.find((f) => (f.apiKey ?? f.attributes?.api_key) === preferred);
    if (match) return match.id;
  }

  // fallback: first file field on the model
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

  // derive preferred api key order: URL override > ctx/appearance params > fallback
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
    if (!fileFieldId) return null;

    // raw can be localized or not
    let raw = (ctx.formValues as any)[fileFieldId];
    raw = pickLocalizedValue(raw, ctx.locale);
    if (!raw) return null;

    // multiple files → take first
    if (Array.isArray(raw)) raw = raw[0];

    // normalize to an object with upload_id (preferred), or allow direct URL (edge)
    if (raw?.upload_id) return raw;
    if (raw?.upload?.id) return { upload_id: raw.upload.id };
    if (typeof raw === 'string' && raw.startsWith('http')) {
      return { __direct_url: raw };
    }
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

      setSheetNames(names);
      setRows(parsed);
      setSheet(names[0] || null);

      // optional meta
      if (params.columnsMetaApiKey) {
        await ctx.setFieldValue(params.columnsMetaApiKey, { columns: inferColumns(parsed) });
      }
      if (params.rowCountApiKey) {
        await ctx.setFieldValue(params.rowCountApiKey, parsed.length);
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
      await ctx.setFieldValue(ctx.fieldPath, rows);
      if (params.columnsMetaApiKey) {
        await ctx.setFieldValue(params.columnsMetaApiKey, { columns: inferColumns(rows) });
      }
      if (params.rowCountApiKey) {
        await ctx.setFieldValue(params.rowCountApiKey, rows.length);
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

  // keep in sync with external changes
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
              if (!picked) {
                setNotice('No file picked.');
                return;
              }
              const res = await fetch(picked.url);
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

        {/* Sheet selector (native select to avoid UI-kit typing differences) */}
        {sheetNames.length > 1 && (
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 12, opacity: 0.8 }}>Sheet</span>
            <select
              value={sheet ?? ''}
              onChange={(e) => setSheet(e.target.value)}
              style={{ padding: 6, borderRadius: 6, border: '1px solid var(--border-color)' }}
            >
              {sheetNames.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
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
            Resolved file field id:{' '}
            {String(resolveFieldId(ctx, preferredApiKey))}
          </div>
        </div>
      )}

      {/* Legacy CSS theme wrapper */}
      <div className="ag-theme-alpine" style={{ height: 420, width: '100%' }}>
        <AgGridReact
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
