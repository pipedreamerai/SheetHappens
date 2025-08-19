/*
 * Archive helpers: capability detection and filename utilities
 */

/* global Office, Excel */

// Detect whether compressed export is available on this host.
// This is the preferred, lossless "Save a Copy" path.
export function isCompressedExportSupported() {
  try {
    return !!(
      Office &&
      Office.context &&
      Office.context.document &&
      typeof Office.context.document.getFileAsync === 'function' &&
      Office.FileType &&
      Office.FileType.Compressed != null
    );
  } catch (_) {
    return false;
  }
}

// Build a timestamped filename like `<WorkbookName>_YYYYMMDD_HHMMSS.xlsx`.
export function buildArchiveFilename(workbookName) {
  try {
    const base = sanitizeBaseName(String(workbookName || 'Workbook')) || 'Workbook';
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const ts = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    return `${base}_${ts}.xlsx`;
  } catch (_) {
    return 'Workbook_00000000_000000.xlsx';
  }
}

// Try to get a friendly workbook name. Prefer the document URL filename; fall back to active worksheet name or 'Workbook'.
export async function getWorkbookNameSafe() {
  // 1) Try document URL filename (works for saved files)
  try {
    const url = Office && Office.context && Office.context.document && Office.context.document.url;
    const nameFromUrl = url ? extractNameFromUrl(url) : null;
    if (nameFromUrl) return nameFromUrl;
  } catch (_) { /* ignore */ }
  // 2) Try active worksheet name as a hint (not ideal, but better than generic)
  try {
    return await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.load(['name']);
      await context.sync();
      const n = ws && ws.name ? `${ws.name}` : 'Workbook';
      return sanitizeBaseName(n) || 'Workbook';
    });
  } catch (_) {
    // 3) Fallback
    return 'Workbook';
  }
}

// ----- internal helpers -----

function sanitizeBaseName(input) {
  // Remove path separators and characters illegal in filenames on common OSes
  const stripped = String(input || '')
    .replace(/^[\s.]+|[\s.]+$/g, '') // trim spaces/dots at ends
    .replace(/[\\/]/g, '-') // slashes
    .replace(/[\u0000-\u001F\u007F]/g, '-') // control chars
    .replace(/[:*?"<>|]/g, '-') // reserved
    .slice(0, 80); // keep it short
  // Drop extension if present (e.g., .xlsx, .xlsm, .xls)
  const noExt = stripped.replace(/\.(xlsx|xlsm|xls|xltx|xltm|xml)$/i, '');
  return noExt || 'Workbook';
}

function extractNameFromUrl(url) {
  try {
    let u = String(url);
    if (!u) return null;
    // Remove query/hash
    const q = u.indexOf('?');
    if (q >= 0) u = u.slice(0, q);
    const h = u.indexOf('#');
    if (h >= 0) u = u.slice(0, h);
    // Take last path segment
    const parts = u.split(/[\\/]/);
    const last = parts[parts.length - 1] || '';
    const base = last.replace(/\.(xlsx|xlsm|xls|xltx|xltm|xml)$/i, '');
    return sanitizeBaseName(base);
  } catch (_) {
    return null;
  }
}


