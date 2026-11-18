import type { Config } from "./config.js";
import type { Db } from "./db.js";
import { AttendanceWriter, type WriterOptions } from "./excel/attendance.js";
import { createLogger, type Log } from "./log.js";
import { QueueProcessor, type ProcessorOptions } from "./queue/processor.js";
import { ScanService } from "./queue/scans.js";
import { BackupService } from "./backup.js";
import type { WorkerStore } from "./workers/store.js";

export interface Context {
  cfg: Config; db: Db; store: WorkerStore; writer: AttendanceWriter; processor: QueueProcessor;
  scans: ScanService; backup: BackupService; log: Log; clock: () => Date;
}
export interface ContextOptions { clock?: () => Date; log?: Log; processor?: Partial<ProcessorOptions>; writer?: WriterOptions }

/** Wires the services. Does not start the queue loop (index.ts / tests call processor.start()). */
export function createContext(cfg: Config, db: Db, store: WorkerStore, opt: ContextOptions = {}): Context {
  const clock = opt.clock ?? (() => new Date()), log = opt.log ?? createLogger(null);
  const operator = db.prepare("SELECT scanned_by AS scannedBy, scanned_username AS scannedUsername, CASE scan_method WHEN 'SCANNER' THEN 'Scanner' WHEN 'MANUAL' THEN 'Manual ID Entry' ELSE 'Unknown' END AS scanMethod FROM operations WHERE operation_id=?");
  const writer = new AttendanceWriter(cfg.dataDir, {
    operatorFor: (id) => operator.get(id) as { scannedBy: string; scannedUsername: string; scanMethod: string } | undefined,
    ...opt.writer,
  });
  const processor = new QueueProcessor(db, writer, { log, batchSize: cfg.batchSize, maxAttempts: cfg.maxAttempts, ...opt.processor });
  const scans = new ScanService({ db, store, processor, clock, timezone: cfg.timezone, scanWaitMs: cfg.scanWaitMs });
  return { cfg, db, store, writer, processor, scans, backup: new BackupService(cfg, db, processor, log, clock), log, clock };
}
