// @vitest-environment node
import { expect, it } from 'vitest';
import { inlineSingleFile } from './inlineSingleFile.mjs';

function fixture() {
  return {
    'index.html': { type: 'asset', source: '<script type="module" crossorigin src="./app.js"></script><link rel="stylesheet" crossorigin href="./app.css">' },
    'app.js': { type: 'chunk', isEntry: true, imports: [], dynamicImports: [], code: 'const value = "</script><!-- $&"; globalThis.value = value;' },
    'app.css': { type: 'asset', source: '@charset "UTF-8";body{color:red}' },
  };
}
it('inlines the entry script and styles and preserves literal replacement characters', () => {
  const bundle = fixture();
  inlineSingleFile().generateBundle({}, bundle);
  expect(Object.keys(bundle)).toEqual(['index.html']);
  const html = bundle['index.html'].source;
  expect(html).toContain('<style>body{color:red}</style>');
  expect(html).not.toMatch(/\s(?:src|href)=/);
  expect(html).toContain('\\x3c/script>\\x3c!-- $&');
  const code = html.match(/<script[^>]*>([\s\S]*?)<\/script>/)[1];
  const context = {};
  new Function('globalThis', code)(context);
  expect(context.value).toBe('</script><!-- $&');
});
it('fails closed for unmatched references, split chunks and external assets', () => {
  for (const mutate of [
    bundle => { bundle['index.html'].source = '<p>Missing assets</p>'; },
    bundle => { bundle['app.js'].imports = ['other.js']; },
    bundle => { bundle['logo.png'] = { type: 'asset', source: 'bytes' }; },
  ]) {
    const bundle = fixture(); mutate(bundle);
    expect(() => inlineSingleFile().generateBundle({}, bundle)).toThrow();
    expect(bundle['app.js']).toBeDefined();
  }
});
