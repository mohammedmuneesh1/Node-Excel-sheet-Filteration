#!/usr/bin/env node
/**
 * Mailing List Merge/Purge
 * Usage: node merge.js merge_test_input.xlsx output.xlsx
 *
 * Structure:
 *   1. readWorkbook()   - load the 4 input sheets
 *   2. normalize*()     - rules 2.1
 *   3. validate         - rules 2.2
 *   4. isDuplicate()    - rules 2.3 (pairwise test)
 *   5. UnionFind groups - transitive matching
 *   6. decide groups    - rules 2.4
 *   7. writeWorkbook()  - section 3 output
 */
const ExcelJS = require('exceljs');

// ---------------------------------------------------------------- reading --

/** Convert any ExcelJS cell value to a plain trimmed-safe string ('' if empty). */
function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v instanceof Date) return v.toISOString();
    if (v.richText) return v.richText.map((r) => r.text).join('');
    if (v.result !== undefined) return cellText(v.result); // formula cell
    if (v.text !== undefined) return cellText(v.text); // hyperlink cell
    return String(v);
  }
  return String(v);
}

/** Read a sheet into an array of {HEADER: rawValueAsText} objects. */
function readSheet(workbook, name) {
  const ws = workbook.getWorksheet(name);
  if (!ws) throw new Error(`Missing sheet: ${name}`);
  const headers = [];
  ws.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
    headers[col] = cellText(cell.value).trim();
  });
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const obj = {};
    let any = false;
    headers.forEach((h, col) => {
      if (!h) return;
      const t = cellText(row.getCell(col).value);
      obj[h] = t;
      if (t.trim() !== '') any = true;
    });
    if (any) rows.push(obj);
  }
  return rows;
}

// ---------------------------------------------------------- normalization --

/** Names/suffix: upper case, keep only A-Z. */
function normalizeName(s) {
  return s.toUpperCase().replace(/[^A-Z]/g, '');
}

/** ZIP: part before "-", trim; 3-5 digits -> left-pad to 5; else null (invalid). */
function normalizeZip(s) {
  const part = s.split('-')[0].trim();
  if (!/^\d{3,5}$/.test(part)) return null;
  return part.padStart(5, '0');
}

/** ADDRESS_1 -> {house, street}; house is null when it is not all digits. */
function normalizeAddress1(s, suffixMap) {
  const tokens = s.toUpperCase().replace(/[.,]/g, ' ').split(/\s+/).filter(Boolean);
  if (tokens.length === 0 || !/^\d+$/.test(tokens[0])) return { house: null, street: '' };
  const street = tokens.slice(1);
  if (street.length > 0) {
    const last = street[street.length - 1];
    if (suffixMap.has(last)) street[street.length - 1] = suffixMap.get(last);
  }
  return { house: tokens[0], street: street.join(' ') };
}

const UNIT_WORDS = new Set(['APT', 'APARTMENT', 'UNIT', 'STE', 'SUITE']);

/** ADDRESS_2: remove . , # ; drop unit words; join the rest without spaces. */
function normalizeUnit(s) {
  return s
    .toUpperCase()
    .replace(/[.,#]/g, '')
    .split(/\s+/)
    .filter((w) => w && !UNIT_WORDS.has(w))
    .join('');
}

// --------------------------------------------------------------- matching --

/** Canonical first name: FULL_NAME if listed as a nickname, else the name itself. */
function canonicalFirst(name, nicknameMap) {
  return nicknameMap.get(name) || name;
}

/**
 * Rule 2.3 conditions 2-4 (condition 1, same address, is guaranteed by the
 * blocking key the caller uses).
 */
function isDuplicate(a, b, nicknameMap) {
  if (a.last !== b.last) return false;
  if (a.suffix && b.suffix && a.suffix !== b.suffix) return false;
  if (!a.first || !b.first) return false;
  const initialMatch =
    (a.first.length === 1 && a.first === b.first[0]) ||
    (b.first.length === 1 && b.first === a.first[0]);
  if (initialMatch) return true;
  return canonicalFirst(a.first, nicknameMap) === canonicalFirst(b.first, nicknameMap);
}

class UnionFind {
  constructor(n) {
    this.parent = Array.from({ length: n }, (_, i) => i);
  }
  find(x) {
    while (this.parent[x] !== x) {
      this.parent[x] = this.parent[this.parent[x]];
      x = this.parent[x];
    }
    return x;
  }
  union(a, b) {
    const ra = this.find(a);
    const rb = this.find(b);
    if (ra !== rb) this.parent[rb] = ra;
  }
}

// ------------------------------------------------------------------- main --

async function main() {
  const [inPath, outPath] = process.argv.slice(2);
  if (!inPath || !outPath) {
    console.error('Usage: node merge.js <input.xlsx> <output.xlsx>');
    process.exit(1);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inPath);
  const plan = readSheet(wb, 'Plan');
  const rawRecords = readSheet(wb, 'Records');
  const nicknameMap = new Map(
    readSheet(wb, 'Nicknames').map((r) => [normalizeName(r.NICKNAME), normalizeName(r.FULL_NAME)])
  );
  const suffixMap = new Map(
    readSheet(wb, 'StreetSuffixes').map((r) => [
      r.WORD.trim().toUpperCase(),
      r.STANDARD.trim().toUpperCase(),
    ])
  );

  // Plan lookup (KEY_CODE -> {type, priority, list})
  const planByKey = new Map();
  for (const p of plan) {
    planByKey.set(p.KEY_CODE.trim(), {
      key: p.KEY_CODE.trim(),
      type: p.TYPE.trim(),
      priority: Number(p.PRIORITY),
      list: p.LIST.trim(),
    });
  }

  // Normalize + validate every record
  const records = [];
  for (const raw of rawRecords) {
    const key = raw.KEY_CODE.trim();
    const file = planByKey.get(key);
    if (!file) {
      console.warn(`Warning: REC_ID ${raw.REC_ID} has unknown KEY_CODE "${key}" - ignored`);
      continue;
    }
    const addr = normalizeAddress1(raw.ADDRESS_1, suffixMap);
    const zip = normalizeZip(raw.ZIP);
    const rec = {
      raw,
      recId: Number(raw.REC_ID),
      key,
      file,
      isSuppress: file.type.toLowerCase() === 'suppress',
      first: normalizeName(raw.FIRST_NAME),
      last: normalizeName(raw.LAST_NAME),
      suffix: normalizeName(raw.SUFFIX),
      house: addr.house,
      street: addr.street,
      unit: normalizeUnit(raw.ADDRESS_2),
      zip,
    };
    rec.valid = rec.last !== '' && zip !== null && addr.house !== null;
    records.push(rec);
  }

  // Matching: only valid records take part. Block by identical address.
  const active = records.filter((r) => r.valid);
  const indexOf = new Map(active.map((r, i) => [r, i]));
  const uf = new UnionFind(active.length);
  const blocks = new Map();
  for (const r of active) {
    const k = [r.zip, r.house, r.street, r.unit].join('|');
    if (!blocks.has(k)) blocks.set(k, []);
    blocks.get(k).push(r);
  }
  for (const block of blocks.values()) {
    for (let i = 0; i < block.length; i++) {
      for (let j = i + 1; j < block.length; j++) {
        if (isDuplicate(block[i], block[j], nicknameMap)) {
          uf.union(indexOf.get(block[i]), indexOf.get(block[j]));
        }
      }
    }
  }

  // Collect groups
  const groups = new Map();
  for (const r of active) {
    const root = uf.find(indexOf.get(r));
    if (!groups.has(root)) groups.set(root, []);
    groups.get(root).push(r);
  }

  // Decide each group (rules 2.4)
  const survivors = []; // {rec, multiCount}
  const drops = []; // {rec, reason, survivorId}
  for (const rec of records) {
    if (!rec.valid && !rec.isSuppress) drops.push({ rec, reason: 'REJECT', survivorId: '' });
  }
  for (const members of groups.values()) {
    const mail = members.filter((r) => !r.isSuppress);
    if (mail.length === 0) continue;
    if (members.some((r) => r.isSuppress)) {
      for (const r of mail) drops.push({ rec: r, reason: 'SUPPRESS', survivorId: '' });
      continue;
    }
    mail.sort((a, b) => a.file.priority - b.file.priority || a.recId - b.recId);
    const winner = mail[0];
    const outsideLists = new Set(
      mail.filter((r) => r.file.type.toLowerCase() === 'outside').map((r) => r.file.list)
    );
    const multiCount = winner.file.type.toLowerCase() === 'outside' ? outsideLists.size : 0;
    survivors.push({ rec: winner, multiCount });
    for (const r of mail.slice(1)) {
      drops.push({
        rec: r,
        reason: r.key === winner.key ? 'DUP_INTRA' : 'DUP_INTER',
        survivorId: winner.recId,
      });
    }
  }
  survivors.sort((a, b) => a.rec.recId - b.rec.recId);
  drops.sort((a, b) => a.rec.recId - b.rec.recId);

  // Summary
  const mailFiles = plan
    .map((p) => planByKey.get(p.KEY_CODE.trim()))
    .filter((f) => f.type.toLowerCase() !== 'suppress');
  const summary = mailFiles.map((f) => ({
    KEY_CODE: f.key, INPUT: 0, REJECT: 0, SUPPRESS: 0, DUP_INTER: 0, DUP_INTRA: 0, NET: 0, MULTIS: 0,
  }));
  const sumByKey = new Map(summary.map((s) => [s.KEY_CODE, s]));
  for (const r of records) if (!r.isSuppress) sumByKey.get(r.key).INPUT++;
  for (const d of drops) sumByKey.get(d.rec.key)[d.reason]++;
  for (const s of survivors) {
    const row = sumByKey.get(s.rec.key);
    row.NET++;
    if (s.multiCount >= 2) row.MULTIS++;
  }
  const total = { KEY_CODE: 'TOTAL' };
  for (const c of ['INPUT', 'REJECT', 'SUPPRESS', 'DUP_INTER', 'DUP_INTRA', 'NET', 'MULTIS']) {
    total[c] = summary.reduce((a, s) => a + s[c], 0);
  }
  summary.push(total);
  for (const s of summary) {
    if (s.INPUT !== s.REJECT + s.SUPPRESS + s.DUP_INTER + s.DUP_INTRA + s.NET) {
      throw new Error(`Summary row ${s.KEY_CODE} does not balance`);
    }
  }

  // Write output
  const out = new ExcelJS.Workbook();
  const addSheet = (name, columns, rows) => {
    const ws = out.addWorksheet(name);
    ws.addRow(columns);
    rows.forEach((r) => ws.addRow(r));
    ws.getRow(1).font = { bold: true };
    ws.columns.forEach((c, i) => (c.width = Math.max(12, columns[i].length + 2)));
  };
  addSheet(
    'Mail',
    ['REC_ID', 'KEY_CODE', 'FIRST_NAME', 'LAST_NAME', 'SUFFIX', 'ADDRESS_1', 'ADDRESS_2', 'CITY', 'STATE', 'ZIP5', 'MULTI_COUNT'],
    survivors.map(({ rec, multiCount }) => [
      rec.recId, rec.key, rec.raw.FIRST_NAME, rec.raw.LAST_NAME, rec.raw.SUFFIX,
      rec.raw.ADDRESS_1, rec.raw.ADDRESS_2, rec.raw.CITY, rec.raw.STATE, rec.zip, multiCount,
    ])
  );
  addSheet(
    'Drops',
    ['REC_ID', 'KEY_CODE', 'REASON', 'SURVIVOR_REC_ID'],
    drops.map((d) => [d.rec.recId, d.rec.key, d.reason, d.survivorId])
  );
  addSheet(
    'Summary',
    ['KEY_CODE', 'INPUT', 'REJECT', 'SUPPRESS', 'DUP_INTER', 'DUP_INTRA', 'NET', 'MULTIS'],
    summary.map((s) => [s.KEY_CODE, s.INPUT, s.REJECT, s.SUPPRESS, s.DUP_INTER, s.DUP_INTRA, s.NET, s.MULTIS])
  );
  // ZIP5 must stay text so leading zeros are kept
  out.getWorksheet('Mail').getColumn(10).numFmt = '@';
  await out.xlsx.writeFile(outPath);

  console.log(`Done. Survivors: ${survivors.length}, drops: ${drops.length}`);
  console.table(summary);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});