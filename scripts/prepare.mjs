import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
await fs.mkdir(path.join(root,'public/vendor'),{recursive:true});
await fs.copyFile(path.join(root,'node_modules/jszip/dist/jszip.min.js'),path.join(root,'public/vendor/jszip.min.js'));
await fs.copyFile(path.join(root,'node_modules/jszip/LICENSE.markdown'),path.join(root,'public/vendor/JSZip-LICENSE.md'));
console.log('Local workbook exporter prepared.');
