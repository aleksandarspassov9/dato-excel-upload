
import { ModuleRegistry, AllCommunityModule } from 'ag-grid-community';
ModuleRegistry.registerModules([AllCommunityModule]);

import ReactDOM from 'react-dom/client';
import {
  connect,
  type RenderFieldExtensionCtx, // ✅ correct type
} from 'datocms-plugin-sdk';
import { Canvas, Button, TextField, SelectField, Spinner } from 'datocms-react-ui';

import { buildClient } from '@datocms/cma-client-browser';
import * as XLSX from 'xlsx';
import { AgGridReact } from 'ag-grid-react';

import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';
import 'datocms-react-ui/styles.css';
import React, { useEffect, useMemo, useState } from 'react';

type TableRow = Record<string, unknown>;

type FieldParams = {
  sourceFileApiKey: string;   // e.g., "sourceFile"
  columnsMetaApiKey?: string; // e.g., "columnsMeta"
  rowCountApiKey?: string;    // e.g., "rowCount"
};

function getFieldByApiKey(apiKey: string, ctx: RenderFieldExtensionCtx) {
  const all = Object.values(ctx.fields);
  return all.find((f: any) => f.attributes.api_key === apiKey) as
    | { id: string; attributes: { api_key: string } }
    | undefined;
}

async function fetchUploadUrlFromValue(
  fileFieldValue: any,
  cmaToken: string,
): Promise<string | null> {
  const uploadId = fileFieldValue?.upload_id;
  if (!uploadId) return null;
  const client = buildClient({ apiToken: cmaToken });
  const upload = await client.uploads.find(uploadId as string);
  return (upload as any)?.url || null;
}

function inferColumns(rows: TableRow[]): string[] {
  const first = rows?.[0] || {};
  return Object.keys(first as object);
}

function toSheetJSRows(
  binary: ArrayBuffer,
  desiredSheetName?: string,
): { rows: TableRow[]; sheetNames: string[] } {
  const wb = XLSX.read(binary, { type: 'array' });
  const names = wb.SheetNames;
  const target = desiredSheetName && names.includes(desiredSheetName)
    ? desiredSheetName
    : names[0];
  const ws = wb.Sheets[target];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null }) as TableRow[];
  return { rows, sheetNames: names };
}

function Alert({ children }: { children: React.ReactNode }) {
  return (
    <div role="alert" style={{
      padding: '8px 12px',
      border: '1px solid var(--border-color)',
      borderRadius: 6,
      marginBottom: 8,
    }}>
      {children}
    </div>
  );
}

function Editor({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  const params = (ctx.parameters as any) as FieldParams;

  const [busy, setBusy] = useState(false);
  const [sheet, setSheet] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [rows, setRows] = useState<TableRow[]>(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    return Array.isArray(initial) ? (initial as TableRow[]) : [];
  });
  const [notice, setNotice] = useState<string | null>(null);

  // Build Ag-Grid columns dynamically
  const columnDefs = useMemo(() => {
    const cols = new Set<string>();
    rows.forEach((r) => Object.keys(r as object).forEach((k) => cols.add(k)));
    if (cols.size === 0) cols.add('column1');
    return Array.from(cols).map((c) => ({ field: c, editable: true }));
  }, [rows]);

  function getFileFieldValue() {
    const field = getFieldByApiKey(params.sourceFileApiKey, ctx);
    if (!field) return null;
    return (ctx.formValues as any)[field.id] ?? null;
  }

  async function importFromSource() {
    try {
      setBusy(true);
      setNotice(null);

      const fileVal = getFileFieldValue();
      if (!fileVal) {
        setNotice('No file selected in the `sourceFile` field.');
        return;
      }

      const currentLocale = ctx.locale || undefined;
      const localized =
        typeof fileVal === 'object' &&
        fileVal !== null &&
        currentLocale &&
        fileVal[currentLocale]
          ? fileVal[currentLocale]
          : fileVal;

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      if (!token) {
        setNotice('Missing CMA token in plugin config.');
        return;
      }

      const url = await fetchUploadUrlFromValue(localized, token);
      if (!url) {
        setNotice('Could not resolve upload URL.');
        return;
      }

      const res = await fetch(url);
      const buf = await res.arrayBuffer();
      const { rows: parsed, sheetNames: names } = toSheetJSRows(buf, sheet || undefined);
      setSheetNames(names);
      setRows(parsed);
      setSheet(names[0] || null);

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

  useEffect(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    if (Array.isArray(initial)) setRows(initial as TableRow[]);
  }, [ctx.fieldPath, ctx.formValues]);

  return (
    <Canvas ctx={ctx}>
      {busy && <Spinner />}
      {notice && <Alert>{notice}</Alert>}

      <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
        <Button onClick={importFromSource} disabled={busy} buttonType="primary">Import from source file</Button>
        <Button onClick={saveJson} disabled={busy} buttonType="primary">Save JSON to field</Button>
        <Button onClick={addRow} disabled={busy} buttonType="muted" buttonSize="s">+ Row</Button>
        <Button onClick={addColumn} disabled={busy} buttonType="muted" buttonSize="s">+ Column</Button>
        {sheetNames.length > 1 && (
        <SelectField
            id="sheet"
            name="sheet"
            label="Sheet"
            value={sheet ?? ''}
            onChange={(v) => setSheet((v ?? '') as string)}
          />
        )}
      </div>

      <div className="ag-theme-alpine" style={{ height: 400, width: '100%' }}>
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

function Config({ ctx }: { ctx: any }) {
  const [token, setToken] = useState<string>((ctx.plugin.attributes.parameters as any)?.cmaToken || '');
  return (
    <Canvas ctx={ctx}>
      <TextField
        id="cmaToken"
        name="cmaToken"
        label="CMA API Token (Uploads: Read)"
        value={token}
        onChange={setToken}
      />
      <Button
        onClick={async () => {
          await ctx.updatePluginParameters({ cmaToken: token });
          ctx.notice('Saved plugin configuration.');
        }}
      >
        Save configuration
      </Button>
    </Canvas>
  );
}

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

// Optional dev harness if opened directly (not in Dato iframe)
if (window.self === window.top) {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <div style={{ padding: 16 }}>
      <h3>Plugin dev harness</h3>
      <p>This page is meant to be embedded in Dato.</p>
    </div>
  );
}
