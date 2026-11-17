import path from 'node:path';
import fs from 'node:fs';
import { createRequire } from 'node:module';

const root = process.cwd();
const req = createRequire(path.join(root, 'package.json'));
const esbuild = req('esbuild');
const entry = path.join(root, 'api/index.ts');
const tmp = path.join(root, 'api/index.bundle.js');

if (fs.existsSync(entry)) {
  try {
    esbuild.buildSync({
      entryPoints: [entry],
      bundle: true,
      platform: 'node',
      format: 'esm',
      target: 'node22',
      outfile: tmp,
      external: ['@vercel/blob', '@netlify/blobs', '@netlify/functions']
    });
    fs.copyFileSync(tmp, entry);
    fs.rmSync(tmp, { force: true });
    console.log('Successfully bundled api/index.ts in-place! Size:', fs.statSync(entry).size);
  } catch (err) {
    console.error('Error bundling api/index.ts:', err);
    process.exit(1);
  }
}
