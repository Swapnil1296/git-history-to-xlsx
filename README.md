## Git History → Excel

**Git History → Excel** is a VS Code extension that exports your git commit history into an Excel workbook, so you can slice, filter, and report on activity by date, time, and author.

### Features

- **One Excel file per workspace** – commits are always appended, never overwritten.
- **Two reporting modes**
  - **All commits**: export every commit in the repository.
  - **Current git user only**: export only commits authored by the active `git config user.name` / `user.email`.
- **Auto‑sync on commit**
  - Watches `.git/HEAD`.
  - When a new commit is detected, the Excel report is updated **for the current git user**.
- **Clean Excel layout**
  - Worksheet: `Git Commits`
  - Columns: `Date`, `Time`, `Commit ID`, `Short ID`, `Commit Message`, `Author`
  - Header row is frozen and has an auto‑filter so you can filter by **date**, **author**, etc.
- **Idempotent / append‑only**
  - Each commit hash is added **once**; re‑running sync does not duplicate rows.
  - If the existing `.xlsx` is corrupted, it is backed up to `.bak` and rebuilt automatically.

### Requirements

- VS Code `1.74.0` or newer.
- A git repository opened as your workspace.
- Node.js installed to build the extension (for local development).

### Commands

Open the **Command Palette** (`Ctrl+Shift+P`) and search for:

- **Git History: Update Excel (All Commits)**  
  Exports all commits in the current repository into the Excel file, appending only new commit hashes.

- **Git History: Update Excel (Current Git User)**  
  Exports only commits whose author matches the current repository’s:
  - `git config user.name`
  - `git config user.email`

- **Git History: Open Excel Sheet**  
  Opens the Excel report in your default spreadsheet application. If the file does not exist, you’ll be prompted to create it.

- **Git History: Set Excel Output Folder**  
  Set where the Excel file is stored (absolute path, or relative to the workspace root).

- **Git History: Filter by Date**  
  Opens the Excel file and reminds you to use the **Date** column filter to narrow by day or range.

### Settings

All settings live under the `gitHistoryToExcel` section:

- **`gitHistoryToExcel.outputFolder`** (string, default `""`)  
  - `""` → Excel goes in the workspace root.  
  - Relative path → resolved under the workspace root, e.g. `reports/git`.  
  - Absolute path → used as‑is, e.g. `D:\reports\git-history`.

- **`gitHistoryToExcel.fileName`** (string, default `"git-commit-history.xlsx"`)  
  Name of the Excel file to create.

- **`gitHistoryToExcel.autoUpdateOnCommit`** (boolean, default `true`)  
  When enabled, the extension watches `.git/HEAD` and automatically:
  - Detects new commits.
  - Appends those commits **only for the current git user** (same behavior as the “Current Git User” command).

### How author filtering works

For the **Current Git User** mode and auto‑sync:

1. The extension runs in the workspace folder:
   - `git config user.name`
   - `git config user.email`
2. It reads commit history using `simple-git`.
3. It keeps only commits where:
   - `author` contains `user.name` (case‑insensitive), **or**
   - `authorEmail` contains `user.email` (case‑insensitive).

If `user.name` and `user.email` are not configured for the repo, the extension shows a warning and skips adding rows for that run.

### Excel file details

- **File location**:  
  `getExcelFilePath(workspaceRoot, outputFolder, fileName)`  
  where `workspaceRoot` is the first workspace folder.

- **Sheet**: `Git Commits`
- **Columns**:
  1. `Date` (`YYYY-MM-DD`)
  2. `Time` (`HH:MM:SS`)
  3. `Commit ID` (full SHA)
  4. `Short ID` (first 7 characters of SHA)
  5. `Commit Message` (single line; newlines stripped)
  6. `Author` (git author name)

The header row is frozen with an auto‑filter range automatically applied to all populated rows.

### Production readiness notes

- **Safe file handling**
  - Directories are created as needed.
  - If an existing Excel file cannot be read (`corrupted zip`, etc.), it is renamed to `*.bak` and a fresh workbook is created.
- **Git repository detection**
  - If the workspace is not a git repo, commands are no‑ops with a friendly warning.
- **Idempotent syncing**
  - Commit hashes already present in the sheet are never re‑added.
- **Minimal surface area**
  - Extension activates only when a workspace contains a `.git` folder or after startup.

### Extension UI in VS Code

- **Commands** are exposed via the Command Palette (no extra views or panels).
- **Notifications** appear in the bottom‑right corner to indicate:
  - How many commits were appended.
  - When no new commits are found.
  - When git user configuration is missing.
  - When the workspace is not a git repo.
- **Source Control integration** is indirect: the extension reads from the repository on disk and writes to Excel; it does not modify git state.

### Icon and branding

Use a simple, legible icon that works on both light and dark themes:

- **File**: `images/icon.png` (256×256 or 128×128 PNG, transparent background).
- **Design suggestion**:
  - Dark navy square background.
  - White Git branching symbol on the left.
  - Green Excel‑style table/grid on the right.
  - No text; flat, minimal, and high‑contrast so it’s readable at small sizes.

Place your final icon at `images/icon.png` – `package.json` already points to this path.

### Local development

1. Install dependencies:

   ```bash
   npm install
   ```

2. Build:

   ```bash
   npm run compile
   ```

3. Run the extension:

   - Open this folder in VS Code.
   - Go to **Run and Debug** (`Ctrl+Shift+D`).
   - Choose **Run Extension**.
   - Press the green **Run** button (or `Fn+F5` / `Ctrl+F5`).

4. In the Extension Development Host window that opens:

   - Open a git repository.
   - Run the commands described above to generate and open the Excel report.

### Packaging for the Marketplace

- Install `vsce` globally:

  ```bash
  npm install -g vsce
  ```

- From the extension root:

  ```bash
  vsce package
  ```

This will produce a `.vsix` file that you can:

- Install locally via **Extensions → ... → Install from VSIX…**
- Or publish to the VS Code Marketplace under your publisher account.

