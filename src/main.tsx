import React, { useEffect, useMemo, useRef, useState } from 'react';
import ReactDOM from 'react-dom/client';
import { connect, RenderManualFieldExtensionCtx } from 'datocms-plugin-sdk';
import { Canvas, Button, TextField, Dropdown, Spinner, Notice } from 'datocms-react-ui';
import { buildClient } from '@datocms/cma-client-browser';
import * as XLSX from 'xlsx';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';
import 'datocms-react-ui/styles.css';

type TableRow = Record<string, any>;

type FieldParams = {
  sourceFileApiKey: string;   // e.g., "sourceFile"
  columnsMetaApiKey?: string; // e.g., "columnsMeta"
  rowCountApiKey?: string;    // e.g., "rowCount"
};

function getFieldByApiKey(apiKey: string, ctx: RenderManualFieldExtensionCtx) {
  const all = Object.values(ctx.fields);
  return all.find((f: any) => f.attributes.api_key === apiKey);
}

async function fetchUploadUrlFromValue(
  fileFieldValue: any,
  cmaToken: string,
): Promise<string | null> {
  const uploadId = fileFieldValue?.upload_id;
  if (!uploadId) return null;
  const client = buildClient({ apiToken: cmaToken });
  const upload = await client.uploads.find(uploadId);
  return (upload as any)?.url || null; // CMA Upload has `url`
}

function inferColumns(rows: TableRow[]): string[] {
  const first = rows?.[0] || {};
  return Object.keys(first);
}

function toSheetJSRows(binary: ArrayBuffer, sheetName?: string): { rows: TableRow[], sheetNames: string[] } {
  const wb = XLSX.read(binary, { type: 'array' });
  const names = wb.SheetNames;
  const target = sheetName && names.includes(sheetName) ? sheetName : names[0];
  const ws = wb.Sheets[target];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null }); // array of objects
  return { rows, sheetNames: names };
}

function Editor({ ctx }: { ctx: RenderManualFieldExtensionCtx }) {
  const params = (ctx.parameters as any) as FieldParams;
  const [busy, setBusy] = useState(false);
  const [sheet, setSheet] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [rows, setRows] = useState<TableRow[]>(() => (ctx.formValues[ctx.fieldPath] as any) || []);
  const [notice, setNotice] = useState<string | null>(null);

  // Build Ag-Grid columns dynamically
  const columnDefs = useMemo(() => {
    const cols = new Set<string>();
    rows.forEach((r) => Object.keys(r).forEach((k) => cols.add(k)));
    if (cols.size === 0) cols.add('column1');
    return Array.from(cols).map((c) => ({ field: c, editable: true }));
  }, [rows]);

  // Helper: get the raw value of the file field from formValues (by field ID)
  function getFileFieldValue() {
    const field = getFieldByApiKey(params.sourceFileApiKey, ctx);
    if (!field) return null;
    // formValues keys are field IDs; for localized fields you'll get per-locale object
    return (ctx.formValues as any)[field.id] || null;
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

      // Handle localized file fields (take current locale if present)
      const currentLocale = ctx.locale || undefined;
      const val = typeof fileVal === 'object' && fileVal !== null && currentLocale && fileVal[currentLocale]
        ? fileVal[currentLocale]
        : fileVal;

      const token = (ctx.plugin.attributes.parameters as any)?.cmaToken || '';
      if (!token) {
        setNotice('Missing CMA token in plugin config.');
        return;
      }

      const url = await fetchUploadUrlFromValue(val, token);
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

      // Optionally update meta fields
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
      // Save table rows into the JSON field this editor is attached to
      await ctx.setFieldValue(ctx.fieldPath, rows);
      // Optional meta refresh
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
    setRows((old) => old.map((r) => ({ ...r, [newKey]: r[newKey] ?? null })));
  }

  // Keep grid state in sync with Dato form when editor opens with existing JSON
  useEffect(() => {
    const initial = (ctx.formValues as any)[ctx.fieldPath];
    if (Array.isArray(initial)) setRows(initial);
  }, [ctx.fieldPath]);

  return (
    <Canvas ctx={ctx}>
      {busy && <Spinner/>}
      {notice && <Notice>{notice}</Notice>}

      <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
        <Button onClick={importFromSource} disabled={busy}>Import from source file</Button>
        <Button onClick={saveJson} disabled={busy}>Save JSON to field</Button>
        <Button onClick={addRow} disabled={busy} variant="quiet">+ Row</Button>
        <Button onClick={addColumn} disabled={busy} variant="quiet">+ Column</Button>
        {sheetNames.length > 1 && (
          <Dropdown
            id="sheet"
            value={sheet || ''}
            options={sheetNames.map((n) => ({ label: n, value: n }))}
            onChange={(v) => setSheet(v as string)}
          />
        )}
      </div>

      <div className="ag-theme-alpine" style={{ height: 400, width: '100%' }}>
        <AgGridReact
          rowData={rows}
          columnDefs={columnDefs}
          onCellValueChanged={(e) => {
            const { rowIndex, colDef, newValue } = e as any;
            const key = colDef.field;
            setRows((prev) => {
              const copy = [...prev];
              copy[rowIndex] = { ...copy[rowIndex], [key]: newValue };
              return copy;
            });
          }}
        />
      </div>
    </Canvas>
  );
}

connect({
  // Config screen: set the CMA token once per plugin
  renderConfigScreen(ctx) {
    function Config() {
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
    ReactDOM.createRoot(document.getElementById('root')!).render(<Config />);
  },

  // Make this plugin available as a manual field editor for JSON fields
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
