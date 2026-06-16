const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');

// Attempt to use OfficeCLI if available for Office formats; otherwise
// fallback to basic text extraction for txt/csv and a simple placeholder.

function runOfficeCliView(filePath) {
  return new Promise((resolve, reject) => {
    const args = ['view', '--json', '--file', filePath];
    const child = execFile('officecli', args, { timeout: 30_000 }, (err, stdout, stderr) => {
      if (err) return reject(err);
      try {
        const j = JSON.parse(stdout);
        resolve(j);
      } catch (e) {
        reject(e);
      }
    });
  });
}

async function parseTextFile(filePath) {
  const buf = await fs.promises.readFile(filePath, 'utf8');
  return { text: buf, canonical: { text: buf } };
}

async function runParser(filePath, originalName) {
  const ext = path.extname(originalName || filePath).toLowerCase();
  try {
    if (ext === '.docx' || ext === '.pptx' || ext === '.xlsx') {
      // Try OfficeCLI
      try {
        const res = await runOfficeCliView(filePath);
        // Normalize into canonical shape for prototype
        const text = JSON.stringify(res).slice(0, 10000);
        return { text, canonical: { source: 'officecli', raw: res } };
      } catch (e) {
        // fallback
        return { text: '', canonical: { error: 'officecli-unavailable', message: String(e) } };
      }
    }
    if (ext === '.txt' || ext === '.csv') return parseTextFile(filePath);

    // Fallback: return file size and name to show it was received
    const stat = await fs.promises.stat(filePath);
    return { text: `Uploaded ${path.basename(filePath)} (${stat.size} bytes)`, canonical: { filename: path.basename(filePath), size: stat.size } };
  } catch (err) {
    return { text: '', canonical: { error: String(err) } };
  }
}

module.exports = { runParser };
