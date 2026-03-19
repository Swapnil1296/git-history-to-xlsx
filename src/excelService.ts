import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';
import { CommitRecord } from './gitHistory';

const SHEET_NAME = 'Git Commits';
const COLS = {
  date: 1,
  time: 2,
  commitId: 3,
  shortCommitId: 4,
  message: 5,
  author: 6,
} as const;

export async function getOrCreateWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(filePath)) {
    try {
      await workbook.xlsx.readFile(filePath);
      return workbook;
    } catch {
      // Existing file is not a valid Excel workbook (e.g. empty or corrupted).
      // Rename it as a backup and create a fresh workbook instead.
      const backupPath = `${filePath}.bak`;
      try {
        fs.renameSync(filePath, backupPath);
      } catch {
        // ignore rename failures; we'll just overwrite the file.
      }
    }
  }
  return workbook;
}

function getOrCreateSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet {
  let sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) {
    sheet = workbook.addWorksheet(SHEET_NAME, {
      views: [{ state: 'frozen', ySplit: 1 }],
    });
    sheet.columns = [
      { header: 'Date', key: 'date', width: 12 },
      { header: 'Time', key: 'time', width: 10 },
      { header: 'Commit ID', key: 'commitId', width: 42 },
      { header: 'Short ID', key: 'shortId', width: 10 },
      { header: 'Commit Message', key: 'message', width: 60 },
      { header: 'Author', key: 'author', width: 20 },
    ];
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' },
    };
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  }
  return sheet;
}

function existingCommitIds(sheet: ExcelJS.Worksheet): Set<string> {
  const ids = new Set<string>();
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const id = row.getCell(COLS.commitId).value?.toString?.();
    if (id) ids.add(id);
  });
  return ids;
}

export async function appendCommitsToExcel(
  filePath: string,
  commits: CommitRecord[],
  deduplicateByCommitId: boolean = true
): Promise<{ appended: number; totalRows: number }> {
  const workbook = await getOrCreateWorkbook(filePath);
  const sheet = getOrCreateSheet(workbook);
  const existing = deduplicateByCommitId ? existingCommitIds(sheet) : new Set<string>();
  let appended = 0;

  for (const c of commits) {
    if (existing.has(c.hash)) continue;
    const row = [
      c.dateStr,
      c.timeStr,
      c.hash,
      c.shortHash,
      c.message,
      c.author,
    ];
    sheet.addRow(row);
    existing.add(c.hash);
    appended++;
  }

  if (sheet.autoFilter) {
    try {
      sheet.autoFilter = undefined;
    } catch (_) {}
  }
  const lastRow = sheet.rowCount;
  if (lastRow >= 2) {
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: lastRow, column: 6 },
    };
  }

  await workbook.xlsx.writeFile(filePath);
  return { appended, totalRows: sheet.rowCount };
}

export function getExcelFilePath(workspaceRoot: string, outputFolder: string, fileName: string): string {
  const base = outputFolder && path.isAbsolute(outputFolder)
    ? outputFolder
    : path.join(workspaceRoot, outputFolder || '');
  return path.join(base, fileName);
}
