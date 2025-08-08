import { connect } from 'datocms-plugin-sdk';
import * as XLSX from 'xlsx';

let ctx; // plugin context

connect({
  onBoot(pluginCtx) {
    ctx = pluginCtx;
    log('Plugin ready. Choose a file to import.');
  }
});

function log(msg) {
  document.getElementById('log').textContent += msg + '\n';
}

document.getElementById('importBtn').addEventListener('click', async () => {
  const fileInput = document.getElementById('fileInput');
  if (!fileInput.files.length) {
    log('No file selected');
    return;
  }

  const file = fileInput.files[0];
  const data = await parseExcel(file);
  log(`Parsed ${data.length} rows from Excel.`);

  const itemTypeId = 'YOUR_ITEM_TYPE_ID'; // find in DatoCMS model settings

  for (const row of data) {
    try {
      const created = await ctx.loaders.items.create({
        itemType: itemTypeId,
        field1: row['Column A'],
        field2: row['Column B'],
      });
      log(`Created item ID: ${created.id}`);
    } catch (err) {
      log(`Error: ${err.message}`);
    }
  }
});

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      resolve(XLSX.utils.sheet_to_json(ws));
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}
