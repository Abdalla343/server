import fs from "node:fs";
import path from "node:path";
import type { Config } from "./config.js";
import type { Db } from "./db.js";
import type { Log } from "./log.js";
import type { QueueProcessor } from "./queue/processor.js";
import { localParts } from "./time.js";

export type BackupKind = "hourly" | "daily" | "manual";
const KEEP: Record<BackupKind, number> = { hourly: 48, daily: 30, manual: 20 };

/** Consistent snapshots: runs while the writer is idle. SQLite via the online backup API; Excel files by copy. */
export class BackupService {
  readonly root: string;
  constructor(private cfg: Config, private db: Db, private proc: QueueProcessor, private log: Log, private clock: () => Date = () => new Date()) {
    this.root = path.join(cfg.dataDir, "backups", "snapshots");
  }

  run(kind: BackupKind): Promise<{ name: string; dir: string }> {
    return this.proc.exclusive(async () => {
      const stamp = this.clock().toISOString().replace(/[:.]/g, "-").slice(0, 19);
      const name = `${kind}-${stamp}`, dir = path.join(this.root, name);
      fs.mkdirSync(dir, { recursive: true });
      await this.db.backup(path.join(dir, "app.db"));
      const wf = path.join(this.cfg.dataDir, "workers", "Workers.xlsx");
      if (fs.existsSync(wf)) fs.copyFileSync(wf, path.join(dir, "Workers.xlsx"));
      // Attendance: today and the previous days (history is also fully contained in app.db's operations table).
      const att = path.join(this.cfg.dataDir, "attendance");
      if (fs.existsSync(att)) {
        const keep = new Set<string>();
        for (let i = 0; i < this.cfg.backupAttendanceDays; i++) keep.add(localParts(new Date(this.clock().getTime() - i * 86_400_000), this.cfg.timezone).date);
        for (const d of fs.readdirSync(att)) if (d === "rejected" || keep.has(d)) fs.cpSync(path.join(att, d), path.join(dir, "attendance", d), { recursive: true });
      }
      this.prune(kind);
      if (this.cfg.backupExtraDir) {
        try { fs.cpSync(dir, path.join(this.cfg.backupExtraDir, name), { recursive: true }); }
        catch (e) { this.log.error({ err: String(e) }, "backup mirror to BACKUP_EXTRA_DIR failed"); }
      }
      this.log.info({ name }, "backup created");
      return { name, dir };
    });
  }

  list() {
    if (!fs.existsSync(this.root)) return [];
    return fs.readdirSync(this.root).sort().reverse().map((name) => ({ name, createdAt: fs.statSync(path.join(this.root, name)).mtime.toISOString() }));
  }

  private prune(kind: BackupKind) {
    const mine = fs.readdirSync(this.root).filter((n) => n.startsWith(kind + "-")).sort();
    for (const n of mine.slice(0, Math.max(0, mine.length - KEEP[kind]))) fs.rmSync(path.join(this.root, n), { recursive: true, force: true });
  }
}
