"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getOrCreateWorkbook = getOrCreateWorkbook;
exports.appendCommitsToExcel = appendCommitsToExcel;
exports.getExcelFilePath = getExcelFilePath;
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const exceljs_1 = __importDefault(require("exceljs"));
const SHEET_NAME = 'Git Commits';
const COLS = {
    date: 1,
    time: 2,
    commitId: 3,
    shortCommitId: 4,
    message: 5,
    author: 6,
};
async function getOrCreateWorkbook(filePath) {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
    const workbook = new exceljs_1.default.Workbook();
    if (fs.existsSync(filePath)) {
        try {
            await workbook.xlsx.readFile(filePath);
            return workbook;
        }
        catch {
            // Existing file is not a valid Excel workbook (e.g. empty or corrupted).
            // Rename it as a backup and create a fresh workbook instead.
            const backupPath = `${filePath}.bak`;
            try {
                fs.renameSync(filePath, backupPath);
            }
            catch {
                // ignore rename failures; we'll just overwrite the file.
            }
        }
    }
    return workbook;
}
function getOrCreateSheet(workbook) {
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
function existingCommitIds(sheet) {
    const ids = new Set();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1)
            return;
        const id = row.getCell(COLS.commitId).value?.toString?.();
        if (id)
            ids.add(id);
    });
    return ids;
}
async function appendCommitsToExcel(filePath, commits, deduplicateByCommitId = true) {
    const workbook = await getOrCreateWorkbook(filePath);
    const sheet = getOrCreateSheet(workbook);
    const existing = deduplicateByCommitId ? existingCommitIds(sheet) : new Set();
    let appended = 0;
    for (const c of commits) {
        if (existing.has(c.hash))
            continue;
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
        }
        catch (_) { }
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
function getExcelFilePath(workspaceRoot, outputFolder, fileName) {
    const base = outputFolder && path.isAbsolute(outputFolder)
        ? outputFolder
        : path.join(workspaceRoot, outputFolder || '');
    return path.join(base, fileName);
}
//# sourceMappingURL=excelService.js.map