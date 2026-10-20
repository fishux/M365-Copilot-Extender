const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const root = path.join(__dirname, '..', 'extension');
execFileSync(process.execPath, [path.join(__dirname, 'build.cjs'), '--check'], { stdio: 'inherit' });
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
if (manifest.manifest_version !== 3 || manifest.permissions.join() !== 'storage') throw new Error('Unexpected manifest permissions');
for (const content of manifest.content_scripts) {
  if (content.run_at !== 'document_start' || content.matches.join() !== 'https://m365.cloud.microsoft/*') throw new Error('Unexpected injection scope');
  if (content.js.length !== 1) throw new Error('Each execution world must have one self-contained script');
  for (const file of content.js) if (!fs.existsSync(path.join(root, file))) throw new Error('Missing ' + file);
}
for (const file of [...Object.values(manifest.icons), ...Object.values(manifest.action.default_icon), manifest.action.default_popup]) {
  if (!fs.existsSync(path.join(root, file))) throw new Error('Missing ' + file);
}
const popup = fs.readFileSync(path.join(root, manifest.action.default_popup), 'utf8');
if ([...popup.matchAll(/<script\b/g)].length !== 1) throw new Error('Popup must have one self-contained script');
for (const match of popup.matchAll(/(?:src|href)="([^"]+)"/g)) {
  if (!fs.existsSync(path.join(root, match[1]))) throw new Error('Missing popup resource ' + match[1]);
}
for (const file of fs.readdirSync(root).filter(file => file.endsWith('.js'))) execFileSync(process.execPath, ['--check', path.join(root, file)]);
console.log('Manifest, file references, permission scope, and JavaScript syntax passed.');
