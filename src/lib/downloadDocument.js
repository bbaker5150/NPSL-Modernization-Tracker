// Firepit mediates ordinary SharePoint file links in the trusted parent. This
// preserves its navigation review and supports stored files (including XLSM)
// without reading their bytes into the app or disguising the file extension.
export function downloadDocument(result, name) {
  const native = result?.type === 'sharepoint-download';
  const url = native ? result.url : URL.createObjectURL(result);
  const link = document.createElement('a');
  link.href = url;
  link.download = native ? result.fileName : name;
  if (native) { link.target = '_blank'; link.rel = 'noopener noreferrer'; }
  document.body.appendChild(link);
  link.click(); link.remove();
  if (!native) setTimeout(() => URL.revokeObjectURL(url), 30000);
}
