import { mkdir, copyFile, rm } from 'node:fs/promises';

const output = new URL('./dist/', import.meta.url);
await rm(output, { recursive: true, force: true });
await mkdir(new URL('vendor/', output), { recursive: true });
for (const file of ['index.html', 'styles.css', 'dark.css', 'app.js', 'model.js', 'excel.js']) {
  await copyFile(new URL(file, import.meta.url), new URL(file, output));
}
await copyFile(new URL('./node_modules/exceljs/dist/exceljs.min.js', import.meta.url), new URL('vendor/exceljs.js', output));
console.log('Static site built in dist/');
