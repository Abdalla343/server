import type { AuthUser } from "../auth.js";
import type { Db } from "../db.js";
import { isFormat14, normalizeCode, normalizeNid } from "../nationalId.js";
import { duplicateMode } from "../settings.js";
import { localParts } from "../time.js";
import type { WorkerStore } from "../workers/store.js";
import type { QueueProcessor } from "./processor.js";

export interface ScanInput { operationId: string; idNumber: string; direction: "IN" | "OUT"; scanMethod?: "SCANNER" | "MANUAL" | "UNKNOWN"; idType?: "auto" | "employee"; gateId?: string | undefined }
export interface ScanReply { status: number; body: Record<string, unknown> }
export interface ScanDeps { db: Db; store: WorkerStore; processor: QueueProcessor; clock: () => Date; timezone: string; scanWaitMs: number }

const keyOf = (raw: string) => { const n = normalizeNid(raw); return isFormat14(n) ? n : normalizeCode(raw); };
const isoDay = (dateKey: string) => `${dateKey.slice(6)}-${dateKey.slice(3, 5)}-${dateKey.slice(0, 2)}`;

export class ScanService {
  constructor(private d: ScanDeps) {}
  private row(opId: string) { return this.d.db.prepare("SELECT * FROM operations WHERE operation_id=?").get(opId) as any; }

  async submit(user: AuthUser, input: ScanInput): Promise<ScanReply> {
    const { db, store, processor } = this.d;
    if (!store.ready) return { status: 503, body: { error: "Worker database is not loaded. Ask an ADMIN to upload Workers.xlsx." } };
    const key = keyOf(input.idNumber);

    const prior = this.row(input.operationId);
    if (prior) {
      if (prior.deleted_at) return { status: 410, body: { error: "This operation was deleted by an admin" } };
      if (prior.user_id !== user.id && user.role !== "ADMIN")
        return { status: 403, body: { error: "Operation belongs to another user" } };
      if (prior.direction !== input.direction || prior.id_input !== key || prior.id_type !== (input.idType ?? "auto") || (input.scanMethod !== undefined && prior.scan_method !== input.scanMethod))
        return { status: 409, body: { error: "operationId already used for a different scan" } };
      if (prior.state === "FAILED") processor.requeue([prior.operation_id]); // the officer pressed "retry"
      return this.respond(input.operationId);
    }

    const worker = store.lookup(input.idNumber, input.idType);
    const now = this.d.clock(), { date, time } = localParts(now, this.d.timezone), day = isoDay(date);
    const mode = duplicateMode(db);

    // One synchronous SQLite transaction: the already-registered check and the insert cannot interleave with another scan.
    db.transaction(() => {
      let outcome = "NOT_FOUND", warning: string | null = null, note: string | null = null, state = "PENDING";
      if (worker?.status === "TERMINATED") outcome = "NOT_APPROVED";
      else if (worker) {
        outcome = "APPROVED"; state = "PENDING";
        const latest = db.prepare("SELECT direction, state FROM active_operations WHERE national_id=? AND outcome='APPROVED' ORDER BY id DESC LIMIT 1")
          .get(worker.nationalId) as { direction: string; state: string } | undefined;
        const last = db.prepare("SELECT direction, time_text FROM active_operations WHERE national_id=? AND day=? AND outcome='APPROVED' ORDER BY id DESC LIMIT 1")
          .get(worker.nationalId, day) as { direction: string; time_text: string } | undefined;
        if (input.direction === "OUT" && (latest?.direction !== "IN" || latest.state !== "COMPLETED")) {
          outcome = "NOT_APPROVED"; note = "NO_CHECK_IN";
        } else if (last && last.direction === input.direction && mode !== "allow") {
          note = `${last.direction} ${last.time_text}`;
          if (mode === "block") { outcome = "ALREADY_REGISTERED"; } else warning = "ALREADY_REGISTERED";
        }
      }
      db.prepare(`INSERT INTO operations (operation_id, direction, id_input, id_type, scan_method, national_id, worker_id, name, job_title, company, worker_status, termination_reason,
          outcome, warning, note, state, gate_id, user_id, scanned_by, scanned_username, ts_iso, date_key, day, time_text, completed_at)
        VALUES (@operation_id,@direction,@id_input,@id_type,@scan_method,@national_id,@worker_id,@name,@job_title,@company,@worker_status,@reason,@outcome,@warning,@note,@state,@gate_id,@user_id,@scanned_by,@scanned_username,@ts_iso,@date_key,@day,@time_text,
          CASE WHEN @state='COMPLETED' THEN strftime('%Y-%m-%dT%H:%M:%fZ','now') END)`)
        .run({ operation_id: input.operationId, direction: input.direction, id_input: key, id_type: input.idType ?? "auto", scan_method: input.scanMethod ?? "UNKNOWN", national_id: worker?.nationalId ?? (input.idType !== "employee" && isFormat14(key) ? key : null),
          worker_id: worker?.workerId ?? null, name: worker?.name ?? null, job_title: worker?.jobTitle ?? null, company: worker?.company ?? null,
          worker_status: worker?.status ?? null, reason: worker?.terminationReason || null, outcome, warning, note, state,
          gate_id: input.gateId ?? user.username, user_id: user.id, scanned_by: user.displayName || user.username, scanned_username: user.username, ts_iso: now.toISOString(), date_key: date, day, time_text: time });
    }).immediate();
    return this.respond(input.operationId);
  }

  /** Waits (bounded) for a queued write, then reports the true state. Success is only reported once COMPLETED. */
  async respond(operationId: string): Promise<ScanReply> {
    let r = this.row(operationId);
    if (r.deleted_at) return { status: 410, body: { error: "This operation was deleted by an admin" } };
    if (r.state === "PENDING" || r.state === "PROCESSING") {
      this.d.processor.kick();
      await this.d.processor.waitFor(operationId, this.d.scanWaitMs);
      r = this.row(operationId);
      if (r.deleted_at) return { status: 410, body: { error: "This operation was deleted by an admin" } };
    }
    return { status: r.state === "COMPLETED" ? 200 : r.state === "FAILED" ? 500 : 202, body: this.view(r) };
  }

  view(r: any) {
    return {
      operationId: r.operation_id, state: r.state, outcome: r.outcome, warning: r.warning ?? undefined,
      scanMethod: r.scan_method, rejectionReason: r.note === "NO_CHECK_IN" ? "NO_CHECK_IN" : undefined, scannedBy: r.scanned_by, scannedUsername: r.scanned_username, previous: r.note ?? undefined, direction: r.direction, date: r.date_key, time: r.time_text,
      worker: r.name ? { nationalId: r.national_id, workerId: r.worker_id, name: r.name, jobTitle: r.job_title, company: r.company,
        status: r.worker_status, terminationReason: r.termination_reason ?? undefined } : undefined,
      error: r.state === "FAILED" ? "Saving failed. Retry required (resend the same operationId)." : undefined,
    };
  }
  get(operationId: string) { const r = this.row(operationId); return r && !r.deleted_at ? this.view(r) : null; }
  owner(operationId: string): number | null { const row = this.row(operationId); return row && !row.deleted_at ? row.user_id : null; }
}
