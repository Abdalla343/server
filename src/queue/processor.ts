import type { Db } from "../db.js";
import type { AttendanceRecord, AttendanceWriter } from "../excel/attendance.js";

/** Test seam: aborts the loop WITHOUT touching the DB, like a process death at that exact point. */
export class SimulatedCrash extends Error {}

export interface ProcessorOptions {
  batchSize: number; maxAttempts: number; baseBackoffMs: number; maxBackoffMs: number;
  beforeMarkCompleted?: () => void;
  log?: { info: (o: any, m?: string) => void; error: (o: any, m?: string) => void };
}
export const DEFAULT_PROCESSOR: ProcessorOptions = { batchSize: 200, maxAttempts: 8, baseBackoffMs: 1000, maxBackoffMs: 30_000 };

interface OpRow {
  id: number; operation_id: string; direction: "IN" | "OUT"; national_id: string; worker_id: string; name: string;
  job_title: string; company: string; worker_status: string; time_text: string; ts_iso: string; date_key: string; attempts: number; id_input: string; outcome: string; termination_reason: string; note: string;
  scanned_by: string; scanned_username: string; scan_method: string;
}
const FINAL = ["COMPLETED", "FAILED"];

/**
 * The single writer. Picks PENDING operations for one (day, direction) file, writes them in ONE save
 * (group commit), then marks them COMPLETED in one SQLite transaction.
 */
export class QueueProcessor {
  private tail: Promise<unknown> = Promise.resolve();
  private waiters = new Map<string, Array<() => void>>();
  private timer: NodeJS.Timeout | undefined;
  private draining = false;
  private stopped = true;
  dead = false;
  lastError: { at: string; message: string } | null = null;
  readonly o: ProcessorOptions;

  constructor(private db: Db, private writer: AttendanceWriter, o: Partial<ProcessorOptions> = {}) {
    this.o = { ...DEFAULT_PROCESSOR, ...o };
  }
  get running() { return !this.stopped && !this.dead; }

  /** Crash recovery, then start. PROCESSING rows can only exist after an unclean stop. */
  start(): { recovered: number; tempFilesRemoved: number } {
    const tempFilesRemoved = this.writer.cleanupTemps();
    const recovered = this.db.prepare("UPDATE operations SET state='PENDING', next_attempt_at=0 WHERE state='PROCESSING' AND deleted_at IS NULL").run().changes;
    this.stopped = false; this.dead = false;
    this.o.log?.info({ recovered, tempFilesRemoved }, "queue started");
    this.kick();
    return { recovered, tempFilesRemoved };
  }
  async stop() { this.stopped = true; clearTimeout(this.timer); await this.tail; }

  /** Runs fn when no batch write is in progress (used by backups and day-file creation). */
  exclusive<T>(fn: () => Promise<T> | T): Promise<T> {
    const run = this.tail.then(fn);
    this.tail = run.catch(() => undefined);
    return run;
  }

  kick() { clearTimeout(this.timer); if (!this.stopped && !this.draining) { this.draining = true; void this.drain(); } }

  /** Resolves when the operation reaches COMPLETED or FAILED, or after ms. */
  waitFor(operationId: string, ms: number): Promise<void> {
    const st = (this.db.prepare("SELECT state FROM active_operations WHERE operation_id=?").get(operationId) as { state: string } | undefined)?.state;
    if (!st || FINAL.includes(st)) return Promise.resolve();
    return new Promise((resolve) => {
      const t = setTimeout(done, ms); const list = this.waiters.get(operationId) ?? [];
      function done() { clearTimeout(t); resolve(); }
      list.push(done); this.waiters.set(operationId, list);
    });
  }
  async waitIdle(timeoutMs = 15_000) {
    const end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      const n = (this.db.prepare("SELECT COUNT(*) c FROM active_operations WHERE state IN ('PENDING','PROCESSING')").get() as { c: number }).c;
      if (!n && !this.deletionStatus().pending) return true; await new Promise((r) => setTimeout(r, 15));
    }
    return false;
  }

  private notify(ids: string[]) { for (const id of ids) { this.waiters.get(id)?.forEach((f) => f()); this.waiters.delete(id); } }

  private async drain() {
    try { for (;;) { const did = await this.exclusive(() => this.processOne()); if (!did || this.stopped) break; } }
    catch (e) {
      if (e instanceof SimulatedCrash) { this.dead = true; this.stopped = true; this.draining = false; return; }
      this.lastError = { at: new Date().toISOString(), message: String((e as Error).message) };
      this.o.log?.error({ err: String(e) }, "queue loop error");
    }
    this.draining = false;
    if (!this.stopped) this.armTimer();
  }

  private armTimer() {
    const m = (this.db.prepare("SELECT MIN(next_attempt_at) m FROM (SELECT next_attempt_at FROM active_operations WHERE state='PENDING' UNION ALL SELECT next_attempt_at FROM excel_deletions WHERE state='PENDING')").get() as { m: number | null }).m;
    const delay = m === null ? 5000 : Math.min(Math.max(m - Date.now(), 10), 5000); // 5 s safety poll
    this.timer = setTimeout(() => this.kick(), delay); this.timer.unref();
  }

  private async processOne(): Promise<boolean> {
    const now = Date.now();
    if (await this.processDeletion()) return true;
    const head = this.db.prepare("SELECT date_key, direction, (outcome='APPROVED') AS successful FROM active_operations WHERE state='PENDING' AND next_attempt_at<=? ORDER BY id LIMIT 1")
      .get(now) as { date_key: string; direction: "IN" | "OUT"; successful: number } | undefined;
    if (!head) return false;
    const ops = this.db.prepare("SELECT * FROM active_operations WHERE state='PENDING' AND next_attempt_at<=? AND date_key=? AND direction=? AND (outcome='APPROVED')=? ORDER BY id LIMIT ?")
      .all(now, head.date_key, head.direction, head.successful, this.o.batchSize) as OpRow[];
    const ids = ops.map((x) => x.operation_id);
    const mark = (sql: string, f: (o: OpRow) => unknown[]) => { const st = this.db.prepare(sql); this.db.transaction(() => ops.forEach((o) => st.run(...f(o))))(); };
    mark("UPDATE operations SET state='PROCESSING' WHERE id=?", (o) => [o.id]);

    try {
      const records: AttendanceRecord[] = ops.map((o) => ({ operationId: o.operation_id, nationalId: o.national_id ?? "", workerId: o.worker_id || (o.outcome === "APPROVED" ? "" : o.id_input),
        name: o.name, jobTitle: o.job_title ?? "", company: o.company, status: o.outcome === "APPROVED" ? "Valid" : o.note === "NO_CHECK_IN" ? "No completed active check-in" : o.termination_reason || o.note || "",
        timeText: o.time_text, timestampIso: o.ts_iso, scanMethod: o.scan_method === "SCANNER" ? "Scanner" : o.scan_method === "MANUAL" ? "Manual ID Entry" : "Unknown", scannedBy: o.scanned_by, scannedUsername: o.scanned_username,
        dateKey: o.date_key, direction: o.direction,
        rejectionReason: o.note === "NO_CHECK_IN" ? "No Check-In" : o.outcome === "NOT_APPROVED" ? "Terminated" : o.outcome === "NOT_FOUND" ? "Not Found" : "Already Registered" }));
      const res = head.successful ? await this.writer.appendBatch(head.direction, head.date_key, records) : await this.writer.appendRejected(records);
      this.o.beforeMarkCompleted?.(); // test seam: "crash" after the Excel save, before SQLite is updated
      mark("UPDATE operations SET state='COMPLETED', last_error=NULL, completed_at=strftime('%Y-%m-%dT%H:%M:%fZ','now') WHERE id=?", (o) => [o.id]);
      this.o.log?.info({ file: `${head.date_key}/${head.direction}`, appended: res.appended.length, duplicates: res.duplicates.length }, "batch saved");
      this.notify(ids);
    } catch (e: any) {
      if (e instanceof SimulatedCrash) throw e;
      const msg = String(e.message ?? e).slice(0, 500);
      this.lastError = { at: new Date().toISOString(), message: msg };
      this.o.log?.error({ code: e.code, msg, ops: ids.length }, "batch failed");
      mark("UPDATE operations SET attempts=attempts+1, last_error=?, state=CASE WHEN ?=1 OR attempts+1>=? THEN 'FAILED' ELSE 'PENDING' END, next_attempt_at=? WHERE id=?",
        (o) => [msg, e.code === "FILE_INVALID" ? 1 : 0, this.o.maxAttempts, Date.now() + Math.min(this.o.baseBackoffMs * 2 ** o.attempts, this.o.maxBackoffMs), o.id]);
      this.notify(ids); // waiters re-read state (PENDING => they get 202 and poll)
    }
    return true;
  }

  notifyDeleted() {
    const deleted = this.db.prepare("SELECT deleted_at FROM operations WHERE operation_id=?");
    for (const id of this.waiters.keys()) {
      if ((deleted.get(id) as { deleted_at?: string } | undefined)?.deleted_at) this.notify([id]);
    }
  }

  deletionStatus() {
    const rows = this.db.prepare("SELECT state, COUNT(*) n FROM excel_deletions GROUP BY state").all() as Array<{ state: string; n: number }>;
    return { pending: rows.find((row) => row.state === "PENDING")?.n ?? 0, failed: rows.find((row) => row.state === "FAILED")?.n ?? 0 };
  }

  async flushDeletions() {
    while (await this.processDeletion()) { /* serialize each affected workbook */ }
  }

  private async processDeletion(): Promise<boolean> {
    const job = this.db.prepare("SELECT * FROM excel_deletions WHERE state='PENDING' AND next_attempt_at<=? ORDER BY file_key LIMIT 1")
      .get(Date.now()) as { file_key: string; direction: "IN" | "OUT" | "REJECTED"; date_key: string; attempts: number } | undefined;
    if (!job) return false;
    const ids = this.db.prepare(job.direction === "REJECTED"
      ? "SELECT operation_id FROM operations WHERE deleted_at IS NOT NULL AND outcome!='APPROVED'"
      : "SELECT operation_id FROM operations WHERE deleted_at IS NOT NULL AND outcome='APPROVED' AND direction=? AND date_key=?")
      .all(...(job.direction === "REJECTED" ? [] : [job.direction, job.date_key])) as Array<{ operation_id: string }>;
    try {
      await this.writer.removeOperations(job.direction, job.date_key, ids.map((row) => row.operation_id));
      this.db.prepare("UPDATE excel_deletions SET state='COMPLETED',last_error=NULL WHERE file_key=?").run(job.file_key);
    } catch (e: any) {
      if (e instanceof SimulatedCrash) throw e;
      const message = String(e.message ?? e).slice(0, 500);
      this.lastError = { at: new Date().toISOString(), message };
      this.db.prepare(`UPDATE excel_deletions SET attempts=attempts+1,last_error=?,
        state=CASE WHEN ? OR attempts+1>=? THEN 'FAILED' ELSE 'PENDING' END,next_attempt_at=? WHERE file_key=?`)
        .run(message, e.code === "FILE_INVALID" ? 1 : 0, this.o.maxAttempts,
          Date.now() + Math.min(this.o.baseBackoffMs * 2 ** job.attempts, this.o.maxBackoffMs), job.file_key);
    }
    return true;
  }

  /** FAILED -> PENDING (admin retry or client resubmission). */
  requeue(operationIds?: string[]): number {
    const st = this.db.prepare("UPDATE operations SET state='PENDING', attempts=0, next_attempt_at=0 WHERE state='FAILED' AND deleted_at IS NULL AND operation_id=?");
    const all = this.db.prepare("UPDATE operations SET state='PENDING', attempts=0, next_attempt_at=0 WHERE state='FAILED' AND deleted_at IS NULL");
    const n = operationIds ? this.db.transaction(() => operationIds.reduce((a, id) => a + st.run(id).changes, 0))() : all.run().changes;
    const deletions = operationIds ? 0 : this.db.prepare("UPDATE excel_deletions SET state='PENDING',attempts=0,next_attempt_at=0 WHERE state='FAILED'").run().changes;
    if (n || deletions) this.kick();
    return n + deletions;
  }
}
