import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { promisify } from 'util';
import { execFile } from 'child_process';
import { getGitHistory } from './gitHistory';
import { appendCommitsToExcel, getExcelFilePath } from './excelService';

let watcher: vscode.FileSystemWatcher | undefined;
let lastKnownHead: string = '';
const execFileAsync = promisify(execFile);

function getWorkspaceRoot(): string | undefined {
  const folder = vscode.workspace.workspaceFolders?.[0];
  return folder?.uri.fsPath;
}

function getConfig() {
  const folder = vscode.workspace.workspaceFolders?.[0];
  const config = vscode.workspace.getConfiguration('gitHistoryToExcel', folder?.uri);
  return {
    outputFolder: config.get<string>('outputFolder') ?? '',
    fileName: config.get<string>('fileName') ?? 'git-commit-history.xlsx',
    autoUpdateOnCommit: config.get<boolean>('autoUpdateOnCommit') ?? true,
  };
}

function getExcelPath(): string | null {
  const root = getWorkspaceRoot();
  if (!root) return null;
  const { outputFolder, fileName } = getConfig();
  return getExcelFilePath(root, outputFolder, fileName);
}

async function getCurrentGitUser(root: string): Promise<{ name?: string; email?: string }> {
  try {
    const opts = { cwd: root };
    const [{ stdout: nameOut }, { stdout: emailOut }] = await Promise.all([
      execFileAsync('git', ['config', 'user.name'], opts),
      execFileAsync('git', ['config', 'user.email'], opts),
    ]);
    const name = nameOut.trim();
    const email = emailOut.trim();
    return {
      name: name || undefined,
      email: email || undefined,
    };
  } catch {
    return {};
  }
}

async function updateExcel(options?: { onlyCurrentUser?: boolean }): Promise<{ appended: number; path: string } | null> {
  const root = getWorkspaceRoot();
  if (!root) {
    vscode.window.showWarningMessage('Git History to Excel: No workspace folder open.');
    return null;
  }
  const excelPath = getExcelPath();
  if (!excelPath) return null;

  try {
    let commits = await getGitHistory(root);
    if (options?.onlyCurrentUser) {
      const { name, email } = await getCurrentGitUser(root);
      if (!name && !email) {
        vscode.window.showWarningMessage(
          'Git History to Excel: git user.name/user.email are not configured for this repository.'
        );
        return { appended: 0, path: excelPath };
      }
      const nameLc = name?.toLowerCase();
      const emailLc = email?.toLowerCase();
      commits = commits.filter((c) => {
        const authorName = c.author.toLowerCase();
        const authorEmail = c.authorEmail.toLowerCase();
        return (
          (nameLc && authorName.includes(nameLc)) ||
          (emailLc && authorEmail.includes(emailLc))
        );
      });
    }
    if (commits.length === 0) {
      return { appended: 0, path: excelPath };
    }
    const { appended, totalRows } = await appendCommitsToExcel(excelPath, commits);
    return { appended, path: excelPath };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    vscode.window.showErrorMessage(`Git History to Excel: ${msg}`);
    return null;
  }
}

async function updateAllCommits(): Promise<void> {
  const result = await updateExcel();
  if (result === null) return;
  if (result.appended > 0) {
    vscode.window.showInformationMessage(`Git History to Excel: Added ${result.appended} commit(s) (all users).`);
  } else {
    vscode.window.showInformationMessage('Git History to Excel: Already up to date.');
  }
}

async function updateCurrentUser(): Promise<void> {
  const result = await updateExcel({ onlyCurrentUser: true });
  if (result === null) return;
  if (result.appended > 0) {
    vscode.window.showInformationMessage('Git History to Excel: Added commits for current git user.');
  } else {
    vscode.window.showInformationMessage(
      'Git History to Excel: No new commits found for the current git user.'
    );
  }
}

async function openExcel(): Promise<void> {
  const excelPath = getExcelPath();
  if (!excelPath) {
    vscode.window.showWarningMessage('Git History to Excel: No workspace open.');
    return;
  }
  if (!fs.existsSync(excelPath)) {
    const create = await vscode.window.showWarningMessage(
      'Excel file not found. Create it now?',
      'Create'
    );
    if (create !== 'Create') return;
    await updateExcel();
  }
  if (fs.existsSync(excelPath)) {
    const uri = vscode.Uri.file(excelPath);
    await vscode.env.openExternal(uri);
    vscode.window.showInformationMessage('Use the Date column dropdown in Excel to filter by date.');
  }
}

async function setOutputPath(): Promise<void> {
  const root = getWorkspaceRoot();
  const current = getConfig().outputFolder;
  const placeHolder = root ? path.join(root, 'git-history') : '';
  const value = await vscode.window.showInputBox({
    prompt: 'Folder path for Excel file (absolute or relative to workspace). Leave empty for workspace root.',
    value: current || placeHolder,
    placeHolder: 'e.g. C:\\Reports or git-history',
  });
  if (value === undefined) return;
  const config = vscode.workspace.getConfiguration('gitHistoryToExcel');
  await config.update('outputFolder', value.trim(), vscode.ConfigurationTarget.Workspace);
  vscode.window.showInformationMessage(`Git History to Excel: Output folder set to "${value.trim() || 'workspace root'}".`);
}

async function filterByDate(): Promise<void> {
  const excelPath = getExcelPath();
  if (!excelPath || !fs.existsSync(excelPath)) {
    vscode.window.showWarningMessage('Git History to Excel: Open or create the Excel sheet first.');
    return;
  }
  await vscode.env.openExternal(vscode.Uri.file(excelPath));
  vscode.window.showInformationMessage(
    'Filter by date: In Excel, click the Date column filter (dropdown) and choose dates or date range.'
  );
}

async function tryAutoUpdate(): Promise<void> {
  const root = getWorkspaceRoot();
  if (!root || !getConfig().autoUpdateOnCommit) return;
  const headPath = path.join(root, '.git', 'HEAD');
  try {
    if (!fs.existsSync(headPath)) return;
    const content = fs.readFileSync(headPath, 'utf8').trim();
    if (content && content !== lastKnownHead) {
      lastKnownHead = content;
      // In auto mode, only update for the current git user to avoid noise.
      await updateExcel({ onlyCurrentUser: true });
    }
  } catch {
    // ignore
  }
}

function setupWatcher(): void {
  const root = getWorkspaceRoot();
  if (!root || !getConfig().autoUpdateOnCommit) return;

  const gitDir = path.join(root, '.git');
  if (!fs.existsSync(gitDir)) return;

  watcher?.dispose();
  const pattern = new vscode.RelativePattern(vscode.Uri.file(gitDir), '**');
  watcher = vscode.workspace.createFileSystemWatcher(pattern);
  watcher.onDidChange(async () => {
    await tryAutoUpdate();
  });
  watcher.onDidCreate(async () => {
    await tryAutoUpdate();
  });
}

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    vscode.commands.registerCommand('gitHistoryToExcel.updateAllCommits', updateAllCommits),
    vscode.commands.registerCommand('gitHistoryToExcel.updateCurrentUserCommits', updateCurrentUser),
    vscode.commands.registerCommand('gitHistoryToExcel.openExcel', openExcel),
    vscode.commands.registerCommand('gitHistoryToExcel.setOutputPath', setOutputPath),
    vscode.commands.registerCommand('gitHistoryToExcel.filterByDate', filterByDate)
  );

  setupWatcher();
  vscode.workspace.onDidChangeConfiguration((e) => {
    if (e.affectsConfiguration('gitHistoryToExcel')) {
      setupWatcher();
    }
  });

  tryAutoUpdate();
}

export function deactivate(): void {
  watcher?.dispose();
}
