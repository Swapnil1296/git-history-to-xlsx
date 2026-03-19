"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getGitHistory = getGitHistory;
exports.getRepoRoot = getRepoRoot;
const simple_git_1 = __importDefault(require("simple-git"));
async function getGitHistory(repoRoot, maxCount = 5000) {
    const git = (0, simple_git_1.default)(repoRoot);
    const isRepo = await git.checkIsRepo();
    if (!isRepo) {
        return [];
    }
    const log = await git.log({ maxCount });
    const records = log.all.map((c) => {
        const d = new Date(c.date);
        const dateStr = d.toISOString().slice(0, 10);
        const timeStr = d.toTimeString().slice(0, 8);
        return {
            hash: c.hash,
            shortHash: c.hash.slice(0, 7),
            message: (c.message || '').replace(/\r?\n/g, ' ').trim(),
            date: d,
            dateStr,
            timeStr,
            author: c.author_name || '',
            authorEmail: c.author_email || '',
        };
    });
    return records;
}
function getRepoRoot(workspaceFolder) {
    return workspaceFolder;
}
//# sourceMappingURL=gitHistory.js.map