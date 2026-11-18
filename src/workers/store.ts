import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import type { Config } from "../config.js";
import { isFormat14, normalizeCode, normalizeNid } from "../nationalId.js";
import { parseWorkersBuffer, type Issue, type Worker } from "./parse.js";

export interface LoadMeta { loadedAt: string; source: string; sha256: string; total: number; valid: number; terminated: number; warningCount: number }
export type LoadResult =
  | { ok: true; meta: LoadMeta; warnings: Issue[] }
  | { ok: false; errors: Issue[]; errorCount: number };
const cap = (a: Issue[], n = 100) => a.slice(0, n);

/** In-memory worker index. The active index is only ever replaced by a fully validated one. */
export class WorkerStore {
  private byNid = new Map<string, Worker>();
  private byWid = new Map<string, Worker>();
  meta: LoadMeta | null = null;
  lastError: { at: string; errors: Issue[] } | null = null;
  private chain: Promise<unknown> = Promise.resolve();
  readonly file: string; readonly backupDir: string;

  constructor(cfg: Pick<Config, "dataDir">) {
    this.file = path.join(cfg.dataDir, "workers", "Workers.xlsx");
    this.backupDir = path.join(cfg.dataDir, "backups");
    fs.mkdirSync(path.dirname(this.file), { recursive: true });
    fs.mkdirSync(this.backupDir, { recursive: true });
  }

  get ready() { return this.meta !== null; }

  /** Loads Workers.xlsx from disk (startup / admin reload). Keeps the previous index on failure. */
  loadFromDisk(): Promise<LoadResult> {
    return this.serial(async () => {
      if (!fs.existsSync(this.file)) return this.fail([{ row: 0, field: "file", message: "Workers.xlsx not found" }]);
      return this.apply(fs.readFileSync(this.file), "disk");
    });
  }

  /** Validates an uploaded file; only if valid: back up the old file, atomically replace it, then swap the index. */
  replaceFromBuffer(buf: Buffer): Promise<LoadResult> {
    return this.serial(async () => {
      const parsed = await parseWorkersBuffer(buf);
      if (parsed.errors.length) return this.fail(parsed.errors);
      const tmp = this.file + ".tmp";
      const fd = fs.openSync(tmp, "w");
      try { fs.writeSync(fd, buf); fs.fsyncSync(fd); } finally { fs.closeSync(fd); }
      try {
        if (fs.existsSync(this.file))
          fs.copyFileSync(this.file, path.join(this.backupDir, `Workers_${new Date().toISOString().replace(/[:.]/g, "-")}.xlsx`));
        fs.renameSync(tmp, this.file);
      } catch (e: any) {
        fs.rmSync(tmp, { force: true });
        return this.fail([{ row: 0, field: "file", message: `Could not replace Workers.xlsx (${e.code ?? e.message}). Previous data kept.` }]);
      }
      return this.install(parsed.workers, parsed.warnings, buf, "upload");
    });
  }

  private async apply(buf: Buffer, source: string): Promise<LoadResult> {
    const parsed = await parseWorkersBuffer(buf);
    if (parsed.errors.length) return this.fail(parsed.errors);
    return this.install(parsed.workers, parsed.warnings, buf, source);
  }

  private install(workers: Worker[], warnings: Issue[], buf: Buffer, source: string): LoadResult {
    const byNid = new Map<string, Worker>(), byWid = new Map<string, Worker>();
    for (const w of workers) { byNid.set(w.nationalId, w); const k = normalizeCode(w.workerId); if (k) byWid.set(k, w); }
    this.byNid = byNid; this.byWid = byWid; // atomic swap of references
    this.lastError = null;
    this.meta = { loadedAt: new Date().toISOString(), source, sha256: crypto.createHash("sha256").update(buf).digest("hex"),
      total: workers.length, valid: workers.filter((w) => w.status === "VALID").length,
      terminated: workers.filter((w) => w.status === "TERMINATED").length, warningCount: warnings.length };
    return { ok: true, meta: this.meta, warnings: cap(warnings) };
  }

  private fail(errors: Issue[]): LoadResult {
    this.lastError = { at: new Date().toISOString(), errors: cap(errors) };
    return { ok: false, errors: cap(errors), errorCount: errors.length };
  }

  /** Synchronous in-memory lookup by National ID (14 digits) or ID Card code (letters/digits, case-insensitive). Never touches disk. */
  lookup(raw: string, idType: "auto" | "employee" = "auto"): Worker | undefined {
    if (idType === "employee") return this.byWid.get(normalizeCode(raw));
    const nid = normalizeNid(raw);
    if (isFormat14(nid)) return this.byNid.get(nid);
    const code = normalizeCode(raw);
    return code ? this.byWid.get(code) : undefined;
  }

  /**
   * Workers whose ID is ONE character away from what was read (OCR misreads are usually one character off).
   * National IDs: one substituted digit. Codes: one substitution, insertion or deletion. Exact matches are excluded.
   */
  similar(raw: string, limit = 5): Worker[] {
    const nid = normalizeNid(raw), out: Worker[] = [];
    if (isFormat14(nid)) {
      for (const w of this.byNid.values()) { let d = 0; for (let i = 0; i < 14 && d < 2; i++) if (w.nationalId[i] !== nid[i]) d++; if (d === 1) out.push(w); if (out.length >= limit) break; }
      return out;
    }
    const code = normalizeCode(raw); if (code.length < 4) return out;
    for (const [k, w] of this.byWid) { if (k !== code && withinOneEdit(k, code)) out.push(w); if (out.length >= limit) break; }
    return out;
  }

  /** Per-company counts for dashboards (in memory, no Excel access). */
  companySummary(): Record<string, { valid: number; terminated: number }> {
    const out: Record<string, { valid: number; terminated: number }> = {};
    for (const w of this.byNid.values()) { const c = (out[w.company] ??= { valid: 0, terminated: 0 }); w.status === "VALID" ? c.valid++ : c.terminated++; }
    return out;
  }

  private serial<T>(fn: () => Promise<T>): Promise<T> {
    const p = this.chain.then(fn, fn);
    this.chain = p.catch(() => undefined);
    return p;
  }
}

/** True when a and b differ by exactly one substitution, insertion or deletion. */
export function withinOneEdit(a: string, b: string): boolean {
  if (a === b || Math.abs(a.length - b.length) > 1) return false;
  let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++;
  if (a.length === b.length) return a.slice(i + 1) === b.slice(i + 1);
  return a.length > b.length ? a.slice(i + 1) === b.slice(i) : b.slice(i + 1) === a.slice(i);
}
