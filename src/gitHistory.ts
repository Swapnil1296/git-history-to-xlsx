import simpleGit, { SimpleGit } from 'simple-git';
import * as path from 'path';

export interface CommitRecord {
  hash: string;
  shortHash: string;
  message: string;
  date: Date;
  dateStr: string;
  timeStr: string;
  author: string;
  authorEmail: string;
}

export async function getGitHistory(repoRoot: string, maxCount: number = 5000): Promise<CommitRecord[]> {
  const git: SimpleGit = simpleGit(repoRoot);
  const isRepo = await git.checkIsRepo();
  if (!isRepo) {
    return [];
  }

  const log = await git.log({ maxCount });
  const records: CommitRecord[] = log.all.map((c) => {
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

export function getRepoRoot(workspaceFolder: string): string | null {
  return workspaceFolder;
}
