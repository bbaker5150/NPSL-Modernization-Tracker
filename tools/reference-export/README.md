# One-time reference document exporter

This branch builds a separate, read-only utility for the remaining ISEA METENG reference attachments. It is not a replacement for the tracker application and should not be merged just to deploy the utility.

The `One-time reference document exporter` workflow runs when this branch is pushed and publishes a separate prerelease, leaving the normal app's latest release unchanged:

- `reference-document-export.html`: self-contained HTML with the existing Forge runtime.
- `Reference-Document-Upload-Guide.html`: printable setup and upload instructions.
- `SHA256SUMS.txt`: release checksums.

Host the exporter temporarily on the **old ISEA METENG site**, using the same Firepit/ASPX-to-HTML process as the tracker. Scan, then download the ZIP. The completed ZIP includes the actual document/folder map in another upload guide and JSON manifest with SHA-256 hashes. Source requests are GET-only; permissions, membership, source records, and destination records are untouched.

The tool reads all list pages, includes archived attachment bytes separately, checks folder identities, preserves empty folders, and refuses an incomplete export. Binary reads use the list's direct attachment file paths instead of the older REST `$value` transport. Requests have a 120-second deadline covering the response body; catalog reads have 30 seconds. Completed downloads stay in memory for retry while the page is open. A manual original-file selection fallback is available when the hosting proxy cannot transport binary responses. Manually supplied files are identified as such in the manifest, with name and known source-size checks.

There is no previous 150 MB JSON-package cap or base64 expansion. ZIP uses STORE compression. Browser memory and ZIP32 limits still apply. Destination uploads use the tracker Reference Documents view and retain its existing 50 MB per-file limit.

The repository contains the Forge console/recorder integration, but no documented host permission auto-confirm API. This utility does not override `confirm`, intercept parent dialogs, substitute a privileged transport, or disable Firepit controls. A host-owned policy change would require a supported Firepit administrator setting.

## Verification

```
npm ci
npx vitest run tools/reference-export/exporter.test.js
npm run build:reference-export
npm run verify:reference-export
node scripts/smoke-reference-export.mjs
```

The browser smoke requires Playwright/Chromium and uses a mocked SharePoint source inside a real iframe srcdoc. It checks GET-only access, retry without repeating completed downloads, manual fallback, ZIP contents/CRC, manifest, destination guide and responsive rendering. Live tenant transport must still be tested after upload.
