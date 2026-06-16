const { execFile } = require('child_process');

function callOfficeCli(args, timeout = 30000) {
  return new Promise((resolve, reject) => {
    execFile('officecli', args, { timeout }, (err, stdout, stderr) => {
      if (err) return reject({ err, stderr });
      resolve({ stdout, stderr });
    });
  });
}

module.exports = { callOfficeCli };
