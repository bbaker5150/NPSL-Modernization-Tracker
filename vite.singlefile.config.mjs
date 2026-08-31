import { execSync } from 'node:child_process';
import path from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { hardenInlineHtml } from './scripts/hardenInlineHtml.mjs';
import { forgeRuntime } from './scripts/forgeRuntime.mjs';

function buildStamp() {
  if (process.env.BUILD_STAMP) return process.env.BUILD_STAMP;
  try {
    const sha = execSync('git rev-parse --short HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
    return `${sha} ${new Date().toISOString().slice(0, 16)}Z`;
  } catch {
    return `local ${new Date().toISOString().slice(0, 16)}Z`;
  }
}

export default defineConfig({
  base: './',
  plugins: [
    react(),
    viteSingleFile({ removeViteModuleLoader: true }),
    hardenInlineHtml(),
    {
      name: 'name-single-file-output',
      enforce: 'post',
      generateBundle(_options, bundle) {
        const entry = bundle['index.html'];
        if (!entry) return;
        delete bundle['index.html'];
        entry.fileName = 'modernization-project-tracker.html';
        bundle['modernization-project-tracker.html'] = entry;
      },
    },
    forgeRuntime({ project: 'Modernization Project Tracker', build: buildStamp() }),
  ],
  publicDir: false,
  build: {
    outDir: 'build-singlefile',
    emptyOutDir: true,
    assetsInlineLimit: Number.MAX_SAFE_INTEGER,
    cssCodeSplit: false,
    rollupOptions: {
      input: path.resolve(process.cwd(), 'index.html'),
      output: { inlineDynamicImports: true },
    },
    chunkSizeWarningLimit: 8192,
  },
  test: { environment: 'jsdom', globals: true },
});
