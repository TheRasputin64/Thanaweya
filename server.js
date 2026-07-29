const path = require('path');
const express = require('express');
const XLSX = require('xlsx');
const EXCEL_FILE = process.env.EXCEL_FILE || path.join(__dirname, 'data.xlsx');
const PORT = process.env.PORT || 3000;
const SHEET_NAME = process.env.SHEET_NAME || null;
const COLUMNS = {
  seatingNo: 'seating_no',
  arabicName: 'arabic_name',
  totalDegree: 'total_degree',
  caseDesc: 'student_case_desc',
};
const SEARCHABLE_FIELDS = new Set([
  COLUMNS.seatingNo,
  COLUMNS.arabicName,
  COLUMNS.caseDesc,
]);
let rows = [];
let stats = null;
let loadedAt = null;
let loadError = null;
function normalizeForSearch(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toLowerCase();
}
function loadWorkbook() {
  console.log(`[load] Reading workbook: ${EXCEL_FILE}`);
  const startedAt = Date.now();
  const workbook = XLSX.readFile(EXCEL_FILE, {
    cellDates: false,
    cellNF: false,
    cellText: false,
  });
  const sheetName = SHEET_NAME || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${workbook.SheetNames.join(', ')}`);
  }
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  const requiredCols = Object.values(COLUMNS);
  if (raw.length > 0) {
    const firstRowCols = Object.keys(raw[0]);
    const missing = requiredCols.filter((c) => !firstRowCols.includes(c));
    if (missing.length > 0) {
      throw new Error(
        `Missing expected column(s): ${missing.join(', ')}. ` +
        `Found columns: ${firstRowCols.join(', ')}. ` +
        `Check that row 1 of the sheet has exactly these headers: ${requiredCols.join(', ')}`
      );
    }
  }
  let sumDegree = 0;
  let countWithDegree = 0;
  let minDegree = Infinity;
  let maxDegree = -Infinity;
  let missingDegreeCount = 0;
  const caseCounts = Object.create(null);
  const missingNameCount0 = { count: 0 };
  const builtRows = new Array(raw.length);
  for (let i = 0; i < raw.length; i++) {
    const r = raw[i];
    const seatingNo = r[COLUMNS.seatingNo];
    const arabicName = r[COLUMNS.arabicName];
    const totalDegreeRaw = r[COLUMNS.totalDegree];
    const caseDesc = r[COLUMNS.caseDesc];
    let totalDegree = null;
    if (totalDegreeRaw !== '' && totalDegreeRaw !== null && totalDegreeRaw !== undefined) {
      const n = Number(totalDegreeRaw);
      if (!Number.isNaN(n)) {
        totalDegree = n;
      }
    }
    if (totalDegree !== null) {
      sumDegree += totalDegree;
      countWithDegree += 1;
      if (totalDegree < minDegree) minDegree = totalDegree;
      if (totalDegree > maxDegree) maxDegree = totalDegree;
    } else {
      missingDegreeCount += 1;
    }
    const caseKey = caseDesc === '' ? '(فارغ / بدون بيان)' : String(caseDesc);
    caseCounts[caseKey] = (caseCounts[caseKey] || 0) + 1;
    if (arabicName === '') missingNameCount0.count += 1;
    builtRows[i] = {
      __idx: i,
      [COLUMNS.seatingNo]: seatingNo,
      [COLUMNS.arabicName]: arabicName,
      [COLUMNS.totalDegree]: totalDegree,
      [COLUMNS.caseDesc]: caseDesc,
      __search_seatingNo: normalizeForSearch(seatingNo),
      __search_arabicName: normalizeForSearch(arabicName),
      __search_caseDesc: normalizeForSearch(caseDesc),
    };
  }
  rows = builtRows;
  const caseDistribution = Object.entries(caseCounts)
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count);
  stats = {
    totalRows: rows.length,
    totalDegree: {
      count: countWithDegree,
      missing: missingDegreeCount,
      average: countWithDegree > 0 ? sumDegree / countWithDegree : null,
      min: countWithDegree > 0 ? minDegree : null,
      max: countWithDegree > 0 ? maxDegree : null,
    },
    caseDistribution,
    missingArabicName: missingNameCount0.count,
    sheetName,
    sourceFile: path.basename(EXCEL_FILE),
  };
  loadedAt = new Date();
  loadError = null;
  const ms = Date.now() - startedAt;
  console.log(`[load] Done: ${rows.length} rows indexed in ${ms}ms`);
}
function fieldToSearchKey(field) {
  switch (field) {
    case COLUMNS.seatingNo: return '__search_seatingNo';
    case COLUMNS.arabicName: return '__search_arabicName';
    case COLUMNS.caseDesc: return '__search_caseDesc';
    default: return null;
  }
}
function stripInternalFields(row) {
  const { __idx, __search_seatingNo, __search_arabicName, __search_caseDesc, ...clean } = row;
  return clean;
}
function paginate(list, page, pageSize) {
  const total = list.length;
  const safePageSize = Math.min(Math.max(pageSize, 1), 500);
  const totalPages = Math.max(Math.ceil(total / safePageSize), 1);
  const safePage = Math.min(Math.max(page, 1), totalPages);
  const start = (safePage - 1) * safePageSize;
  const slice = list.slice(start, start + safePageSize).map(stripInternalFields);
  return {
    items: slice,
    page: safePage,
    pageSize: safePageSize,
    total,
    totalPages,
  };
}
const app = express();
app.use(express.static(path.join(__dirname, 'public')));
app.get('/api/health', (req, res) => {
  res.json({
    ok: !loadError,
    loadedAt,
    rowCount: rows.length,
    error: loadError,
  });
});
app.get('/api/stats', (req, res) => {
  if (loadError) return res.status(500).json({ error: loadError });
  res.json(stats);
});
app.get('/api/columns', (req, res) => {
  res.json({
    columns: COLUMNS,
    searchable: Array.from(SEARCHABLE_FIELDS),
  });
});
app.get('/api/rows', (req, res) => {
  if (loadError) return res.status(500).json({ error: loadError });
  const page = parseInt(req.query.page, 10) || 1;
  const pageSize = parseInt(req.query.pageSize, 10) || 50;
  res.json(paginate(rows, page, pageSize));
});
app.get('/api/search', (req, res) => {
  if (loadError) return res.status(500).json({ error: loadError });
  const field = req.query.field;
  const qRaw = req.query.q || '';
  const page = parseInt(req.query.page, 10) || 1;
  const pageSize = parseInt(req.query.pageSize, 10) || 50;
  if (!field || !SEARCHABLE_FIELDS.has(field)) {
    return res.status(400).json({
      error: `Invalid or missing "field". Must be one of: ${Array.from(SEARCHABLE_FIELDS).join(', ')}`,
    });
  }
  const searchKey = fieldToSearchKey(field);
  const q = normalizeForSearch(qRaw);
  if (q === '') {
    return res.json(paginate(rows, page, pageSize));
  }
  const matches = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][searchKey].startsWith(q)) {
      matches.push(rows[i]);
    }
  }
  res.json(paginate(matches, page, pageSize));
});
app.get('/api/reload', (req, res) => {
  try {
    loadWorkbook();
    res.json({ ok: true, rowCount: rows.length, loadedAt });
  } catch (err) {
    loadError = err.message;
    res.status(500).json({ ok: false, error: err.message });
  }
});
try {
  loadWorkbook();
} catch (err) {
  loadError = err.message;
  console.error(`[load] FAILED: ${err.message}`);
  console.error('[load] Server will still start, but API calls will return this error until fixed.');
}
app.listen(PORT, () => {
  console.log(`\nExcel Results Viewer running at http://localhost:${PORT}`);
  console.log(`Reading from: ${EXCEL_FILE}`);
  if (loadError) {
    console.log(`WARNING: load failed — ${loadError}`);
  }
});