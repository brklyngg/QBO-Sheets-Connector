import { readFileSync, writeFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawn } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const CODE_FILE = path.resolve(__dirname, '..', 'Code.js');

function incrementVersion(currentVersion) {
  const parts = currentVersion.split('.').map(Number);

  if (parts.length !== 3 || parts.some(Number.isNaN)) {
    throw new Error(`SCRIPT_VERSION is not in semver format: ${currentVersion}`);
  }

  const [major, minor, patch] = parts;
  const nextPatch = patch + 1;

  return [major, minor, nextPatch].join('.');
}

function updateScriptVersion() {
  const source = readFileSync(CODE_FILE, 'utf8');
  const versionRegex = /const SCRIPT_VERSION = '([^']+)';/;
  const match = source.match(versionRegex);

  if (!match) {
    throw new Error('Could not locate SCRIPT_VERSION constant in Code.js.');
  }

  const currentVersion = match[1];
  const nextVersion = incrementVersion(currentVersion);
  const updatedSource = source.replace(versionRegex, `const SCRIPT_VERSION = '${nextVersion}';`);

  writeFileSync(CODE_FILE, updatedSource, 'utf8');

  return { currentVersion, nextVersion };
}

function runClaspPush() {
  return new Promise((resolve, reject) => {
    const child = spawn('clasp', ['push'], {
      stdio: 'inherit',
      cwd: path.resolve(__dirname, '..')
    });

    child.on('exit', code => {
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`clasp push exited with code ${code}`));
      }
    });

    child.on('error', reject);
  });
}

async function main() {
  const { currentVersion, nextVersion } = updateScriptVersion();
  console.log(`SCRIPT_VERSION bumped: ${currentVersion} -> ${nextVersion}`);

  await runClaspPush();
  console.log('clasp push completed successfully.');
}

main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
