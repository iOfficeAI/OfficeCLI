const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const net = require('net');

const UPLOAD_DIR = path.join(__dirname, 'uploads');

function checkUploadsDir() {
  const info = { exists: false, writable: false };
  try {
    info.exists = fs.existsSync(UPLOAD_DIR);
    if (!info.exists) {
      fs.mkdirSync(UPLOAD_DIR, { recursive: true });
      info.exists = true;
    }
    const testFile = path.join(UPLOAD_DIR, `.startup_test_${Date.now()}`);
    fs.writeFileSync(testFile, 'ok');
    fs.unlinkSync(testFile);
    info.writable = true;
  } catch (err) {
    info.error = String(err);
  }
  return info;
}

function checkBinary(name) {
  try {
    const cmd = spawnSync('bash', ['-lc', `command -v ${name} || which ${name}`], { encoding: 'utf8' });
    const out = (cmd.stdout || '').trim();
    if (out) return { found: true, path: out };
    return { found: false };
  } catch (err) {
    return { found: false, error: String(err) };
  }
}

function checkPortFree(port, timeout = 1000) {
  return new Promise((resolve) => {
    const socket = new net.Socket();
    let called = false;
    socket.setTimeout(timeout);
    socket.on('connect', () => {
      called = true;
      socket.destroy();
      resolve({ port, available: false, note: 'port in use (connection succeeded)' });
    });
    socket.on('timeout', () => {
      if (!called) { called = true; socket.destroy(); resolve({ port, available: true, note: 'timeout - assuming free' }); }
    });
    socket.on('error', (err) => {
      if (!called) { called = true; resolve({ port, available: true, note: 'connection refused - port likely free' }); }
    });
    socket.connect(port, '127.0.0.1');
  });
}

async function getStartupStatus(opts = {}) {
  const port = Number(process.env.PORT || opts.port || 4000);
  const uploads = checkUploadsDir();
  const officecli = checkBinary('officecli');
  const openai = { present: !!process.env.OPENAI_API_KEY };
  const node = { version: process.version };
  const portCheck = await checkPortFree(port);

  return {
    timestamp: Date.now(),
    node,
    env: { PORT: port, OPENAI_API_KEY: openai.present ? 'set' : 'not-set' },
    uploads,
    binaries: { officecli },
    port: portCheck
  };
}

if (require.main === module) {
  getStartupStatus().then((s) => {
    console.log(JSON.stringify(s, null, 2));
    // non-zero exit if critical failures
    const ok = s.uploads.writable && s.port.available;
    process.exit(ok ? 0 : 2);
  }).catch((err) => {
    console.error(err);
    process.exit(3);
  });
}

module.exports = { getStartupStatus };
