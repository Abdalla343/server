import type { Config } from "./config.js";
import type { Db } from "./db.js";
import type { Log } from "./log.js";
import type { AttendanceWriter } from "./excel/attendance.js";
import type { QueueProcessor } from "./queue/processor.js";
import type { BackupService } from "./backup.js";
import { getSetting, setSetting } from "./settings.js";
import { localParts } from "./time.js";

/** Local replacement for cron: creates each day's files when the date changes, and runs scheduled backups. */
export class Scheduler {
  private timer: NodeJS.Timeout | undefined; private lastDate = ""; private busy = false;
  constructor(private cfg: Config, private db: Db, private writer: AttendanceWriter, private proc: QueueProcessor,
    private backup: BackupService, private log: Log, private clock: () => Date = () => new Date()) {}

  start(intervalMs = 30_000) { void this.tick(); this.timer = setInterval(() => void this.tick(), intervalMs); this.timer.unref(); }
  stop() { clearInterval(this.timer); }

  async tick() {
    if (this.busy) return; this.busy = true;
    try {
      const now = this.clock(), { date, time } = localParts(now, this.cfg.timezone);
      if (date !== this.lastDate) { // first tick of a new local day (and at startup)
        await this.proc.exclusive(() => this.writer.ensureDay(date)); this.lastDate = date;
        this.log.info({ date }, "day files ensured");
      }
      const last = Number(getSetting(this.db, "lastHourlyBackup", "0"));
      if (now.getTime() - last >= this.cfg.backupIntervalMin * 60_000) {
        setSetting(this.db, "lastHourlyBackup", String(now.getTime())); await this.backup.run("hourly");
      }
      if (time >= this.cfg.backupDailyTime && getSetting(this.db, "lastDailyBackup", "") !== date) {
        setSetting(this.db, "lastDailyBackup", date); await this.backup.run("daily");
      }
    } catch (e) { this.log.error({ err: String(e) }, "scheduler tick failed"); }
    finally { this.busy = false; }
  }
}
