// src/main.tsx
import React, { useEffect, useMemo, useState } from 'react';
import ReactDOM from 'react-dom/client';
import {
  connect,
  type RenderFieldExtensionCtx,
} from 'datocms-plugin-sdk';
import { Canvas, Button, TextField, Spinner } from 'datocms-react-ui';

import { buildClient } from '@datocms/cma-client-browser';
import * as XLSX from 'xlsx';
import { AgGridReact } from 'ag-grid-react';

// AG Grid v31+ modular API: register a row model (enables Community features)
import { ModuleRegistry, ClientSideRowModelModule } from 'ag-grid-community';
ModuleRegistry.registerModules([ClientSideRowModelModule]);

// AG Grid CSS File Themes (legacy approach; DO NOT use Theming API simultaneously)
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';

// Dato UI kit styles
import 'datocms-react-ui/styles.css';

// ===== Config =====
// TEMP hardcoded fallback so you can proceed even if ctx.parameters is empty
const DEFAULT_SOURCE_FILE_API_KEY = 'sourcefile';

type TableRow = Record<string, unknown>;

type FieldParams = {
  sourceFileApiKey?: string;
  columnsMetaApiKey?: string;
  rowCountApiKey?: string;
};

// ===== Helpers =====
function getEditorParams(ctx: RenderFieldExtensionCtx): FieldParams {
  const direct = (ctx.parameters as any) || {};

  console.log(ctx, 'params')
  if (direct && Object.keys(direct).length) return direct;

  // Fallback for some SDK versions: read from field appearance
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
    // Treat value as ID
    const byId = (ctx.fields as any)[preferred];
    if (byId?.id) return byId.id;

    // Or as API key
    const byApiKey = fields.find(
      f => (f.apiKey ?? f.attributes?.api_key) === preferred,
    );
    if (byApiKey) return byApiKey.id;
  }

  // Fallback: first file field in the model
  const firstFile = fields.find(
    f => (f.fieldType ?? f.attributes?.field_type) === 'file',
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
    preferredSheet && names.includes(preferredSheet)
      ? preferredSheet
      : names[0];
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

// ===== Editor =====
function Editor({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = getEditorParams(ctx);

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
    rows.forEach(r => Object.keys(r as object).forEach(k => cols.add(k)));
    if (cols.size === 0) cols.add('column1');
    return Array.from(cols).map(c => ({ field: c, editable: true }));
  }, [rows]);

  function getFileFieldValue() {
    const preferredKey = params.sourceFileApiKey || DEFAULT_SOURCE_FILE_API_KEY;
    const fileFieldId = resolveFieldId(ctx, preferredKey);
    if (!fileFieldId) return null;
    const raw = (ctx.formValues as any)[fileFieldId];
    return pickLocalizedValue(raw, ctx.locale);
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
      if (!token) {
        setNotice(
          'Missing CMA token in plugin configuration (Settings → Plugins → this plugin → Configuration).',
        );
        return;
      }

      const url = await fetchUploadUrlFromValue(fileVal, token);
      if (!url) {
        setNotice('Could not resolve upload URL from the file field value.');
        return;
      }

      const res = await fetch(url);
      const buf = await res.arrayBuffer();
      const { rows: parsed, sheetNames: names } = toSheetJSRows(
        buf,
        sheet || undefined,
      );

      setSheetNames(names);
      setRows(parsed);
      setSheet(names[0] || null);

      if (params.columnsMetaApiKey) {
        await ctx.setFieldValue(params.columnsMetaApiKey, {
          columns: inferColumns(parsed),
        });
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
        await ctx.setFieldValue(params.columnsMetaApiKey, {
          columns: inferColumns(rows),
        });
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
    setRows(r => [...r, {}]);
  }

  function addColumn() {
    const newKey = `col_${Date.now().toString().slice(-4)}`;
    setRows(old => old.map(r => ({ ...r, [newKey]: (r as any)[newKey] ?? null })));
  }

  // Keep in sync if formValues change externally
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
          onClick={() => setShowDebug(v => !v)}
          disabled={busy}
          buttonType="muted"
          buttonSize="s"
        >
          {showDebug ? 'Hide debug' : 'Show debug'}
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
          <div>Params (ctx.parameters or appearance): {JSON.stringify(params)}</div>
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
            {String(resolveFieldId(ctx, params.sourceFileApiKey || DEFAULT_SOURCE_FILE_API_KEY))}
          </div>
        </div>
      )}

      {/* Legacy CSS theme container (no Theming API prop) */}
      <div className="ag-theme-alpine" style={{ height: 420, width: '100%' }}>
        <AgGridReact
          rowData={rows as any[]}
          columnDefs={columnDefs as any}
          onCellValueChanged={(e: any) => {
            const { rowIndex, colDef, newValue } = e;
            const key = colDef.field as string;
            setRows(prev => {
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

// ===== Plugin config screen (CMA token) =====
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

// ===== Wiring =====
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
        This page is designed to be embedded in Dato. Add it as a Private Plugin, then attach it as the
        Field editor for your <code>dataJson</code> field (Presentation tab).
      </p>
    </div>,
  );
}
