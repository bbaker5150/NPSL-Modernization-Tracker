// This app has one HTML entry and one JS/CSS bundle. Inline those exact output
// names without a glob dependency, and fail if any asset would be left behind.
const escapeRegex = value => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function inlineSingleFile() {
  return {
    name: 'tracker:inline-single-file',
    enforce: 'post',
    generateBundle(_options, bundle) {
      const entry = bundle['index.html'];
      if (!entry || entry.type !== 'asset') throw new Error('Expected index.html for the single-file build.');
      let html = String(entry.source);
      const inlined = [];
      for (const [name, asset] of Object.entries(bundle)) {
        if (name === 'index.html') continue;
        const filename = escapeRegex(name);
        if (asset.type === 'chunk' && asset.isEntry && !asset.imports.length && !asset.dynamicImports.length) {
          const tag = new RegExp(`<script\\b([^>]*?)\\s+src="(?:\\./)?${filename}"([^>]*)><\\/script>`, 'g');
          let count = 0;
          // Protect the HTML parser before the subsequent AST hardening pass.
          const code = asset.code.replace(/<(?=\/script\b|!--)/gi, '\\x3c');
          html = html.replace(tag, (_match, before, after) => { count++; return `<script${before}${after}>${code}</script>`; });
          if (count !== 1) throw new Error(`Expected exactly one script reference to ${name}; found ${count}.`);
        } else if (asset.type === 'asset' && name.endsWith('.css')) {
          const tag = new RegExp(`<link\\b[^>]*?\\s+href="(?:\\./)?${filename}"[^>]*>`, 'g');
          const css = String(asset.source).replace(/@charset\s+"UTF-8";/gi, '');
          if (/<\/style\b/i.test(css)) throw new Error('Unsafe closing style tag in emitted CSS.');
          let count = 0;
          html = html.replace(tag, () => { count++; return `<style>${css}</style>`; });
          if (count !== 1) throw new Error(`Expected exactly one stylesheet reference to ${name}; found ${count}.`);
        } else throw new Error(`Unexpected external asset in single-file build: ${name}`);
        inlined.push(name);
      }
      entry.source = html;
      for (const name of inlined) delete bundle[name];
    },
  };
}
