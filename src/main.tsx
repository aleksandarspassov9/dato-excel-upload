import ReactDOM from 'react-dom/client';
import { connect, type RenderFieldExtensionCtx } from 'datocms-plugin-sdk';
import { Canvas } from 'datocms-react-ui';
import 'datocms-react-ui/styles.css';

function Editor({ ctx }: { ctx: RenderFieldExtensionCtx }) {
  return (
    <Canvas ctx={ctx}>
      <div style={{ padding: 12 }}>
        <h3>Test Editor</h3>
        <p><strong>ctx.parameters:</strong> {JSON.stringify(ctx.parameters)}</p>
        <p><strong>Appearance parameters fallback:</strong> {JSON.stringify(
          (ctx.field as any)?.attributes?.appearance?.parameters || (ctx as any)?.fieldAppearance?.parameters || {}
        )}</p>
      </div>
    </Canvas>
  );
}

connect({
  // No config screen for this test
  renderConfigScreen() {
    ReactDOM.createRoot(document.getElementById('root')!).render(
      <div style={{ padding: 12 }}>Config screen (not used in test)</div>
    );
  },

  manualFieldExtensions() {
    return [
      {
        id: 'testEditor',
        name: 'TEST — Params Check',
        type: 'editor',
        fieldTypes: ['json'],
        parameters: [
          { id: 'sourceFileApiKey', name: 'Source File API key', type: 'string', required: true },
        ],
      },
    ];
  },

  renderFieldExtension(id, ctx) {
    if (id === 'testEditor') {
      ReactDOM.createRoot(document.getElementById('root')!).render(<Editor ctx={ctx} />);
    }
  },
});

// Dev harness
if (window.self === window.top) {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <div style={{ padding: 12 }}>Open me inside Dato as a field editor.</div>
  );
}
