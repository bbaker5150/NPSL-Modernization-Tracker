import path from 'node:path';
import { defineConfig } from 'vite';
import { inlineSingleFile } from './scripts/inlineSingleFile.mjs';
import { hardenInlineHtml } from './scripts/hardenInlineHtml.mjs';
import { forgeRuntime } from './scripts/forgeRuntime.mjs';
import { uploadGuide } from './tools/reference-export/guide.js';
export default defineConfig({
  root: 'tools/reference-export', base: './', publicDir: false,
  plugins: [inlineSingleFile(), hardenInlineHtml(), {
    name: 'reference-export-output', enforce: 'post',
    generateBundle(_options, bundle) {
      const entry = bundle['index.html']; delete bundle['index.html'];
      entry.fileName = 'reference-document-export.html'; bundle[entry.fileName] = entry;
    },
  }, forgeRuntime({ project: 'Modernization Tracker Reference Export', build: process.env.BUILD_STAMP || 'reference-export' }), {
    name: 'upload-guide', enforce: 'post',
    generateBundle() { this.emitFile({ type: 'asset', fileName: 'Reference-Document-Upload-Guide.html', source: uploadGuide() }); },
  }],
  build: { outDir: path.resolve('build-reference-export'), emptyOutDir: true, assetsInlineLimit: () => true, modulePreload: false, cssCodeSplit: false, rollupOptions: { output: { inlineDynamicImports: true } } },
});
