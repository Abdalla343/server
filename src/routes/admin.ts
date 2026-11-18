import { Router } from "express";
import fs from "node:fs";
import path from "node:path";
import { z } from "zod";
import { authenticate, requireRole } from "../auth.js";
import type { Context } from "../context.js";
import { audit } from "../db.js";
import { buildExport } from "../export.js";
import { attendancePath, rejectedPath } from "../excel/attendance.js";
import { isFormat14, normalizeCode, normalizeNid } from "../nationalId.js";
import { duplicateMode, setSetting } from "../settings.js";
import { deleteOperations } from "../operationDeletion.js";
import { onSiteCount, onSitePage } from "../presence.js";
import { localParts } from "../time.js";

const DAY = /^\d{4}-\d{2}-\d{2}$/, DATE_KEY = /^\d{2}-\d{2}-\d{4}$/;
const isoDay = (k: string) => `${k.slice(6)}-${k.slice(3, 5)}-${k.slice(0, 2)}`;
const operationView = (r: any) => ({ operationId: r.operation_id, time: r.time_text, date: r.date_key, direction: r.direction, outcome: r.outcome, state: r.state,
  nationalId: r.national_id || String(r.id_input), scanMethod: r.scan_method, rejectionReason: r.note === "NO_CHECK_IN" ? "NO_CHECK_IN" : undefined, scannedBy: r.scanned_by, scannedUsername: r.scanned_username, name: r.name, company: r.company, gate: r.gate_id, warning: r.warning ?? undefined });

export function adminRouter(c: Context) {
  const r = Router(), { db } = c;
  r.use("/admin", authenticate(c.cfg, db), requireRole("ADMIN"));
  r.use(["/attendance", "/dashboard", "/queue"], authenticate(c.cfg, db), requireRole("ADMIN"));
  const today = () => { const d = localParts(c.clock(), c.cfg.timezone).date; return isoDay(d); };
  const one = <T = any>(sql: string, ...a: unknown[]) => db.prepare(sql).get(...a) as T;
  const all = <T = any>(sql: string, ...a: unknown[]) => db.prepare(sql).all(...a) as T[];

  r.get("/admin/onsite", (req, res) => {
    const offset = Number(req.query.offset ?? 0), limit = Number(req.query.limit ?? 100);
    if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(limit) || limit < 1 || limit > 200)
      return void res.status(400).json({ error: "Invalid pagination" });
    res.setHeader("Cache-Control", "no-store");
    res.json(onSitePage(db, typeof req.query.search === "string" ? req.query.search : "", offset, limit, c.clock()));
  });

  r.get("/admin/operations", (req, res) => {
    const offset = Number(req.query.offset ?? 0), limit = Number(req.query.limit ?? 100);
    if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(limit) || limit < 1 || limit > 200)
      return void res.status(400).json({ error: "Invalid pagination" });
    res.json({
      total: one("SELECT COUNT(*) n FROM active_operations").n,
      operations: all("SELECT * FROM active_operations ORDER BY id DESC LIMIT ? OFFSET ?", limit, offset).map(operationView),
      excelPending: c.processor.deletionStatus(),
    });
  });

  r.post("/admin/operations/delete", async (req, res) => {
    const p = z.union([
      z.object({ all: z.literal(true) }).strict(),
      z.object({ operationIds: z.array(z.string().min(1).max(64)).min(1).max(500) }).strict(),
    ]).safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "Select operationIds or explicitly request all=true" });
    res.json(await deleteOperations(c, req.user!.id, "all" in p.data ? "all" : p.data.operationIds));
  });

  // ---- queue ----
  r.get("/queue/status", (_q, res) => {
    const counts = Object.fromEntries(all("SELECT state, COUNT(*) n FROM active_operations GROUP BY state").map((x) => [x.state, x.n]));
    const oldest = one("SELECT MIN(created_at) t FROM active_operations WHERE state IN ('PENDING','PROCESSING')").t as string | null;
    res.json({ counts, deletions: c.processor.deletionStatus(), oldestPendingSeconds: oldest ? Math.round((Date.now() - Date.parse(oldest)) / 1000) : 0,
      processor: { running: c.processor.running, dead: c.processor.dead, lastError: c.processor.lastError },
      failed: all("SELECT * FROM active_operations WHERE state='FAILED' ORDER BY id LIMIT 50").map((x) => ({ ...operationView(x), error: x.last_error })) });
  });
  r.post("/queue/retry-failed", (req, res) => { const n = c.processor.requeue(); audit(db, req.user!.id, "QUEUE_RETRY_ALL", String(n)); res.json({ requeued: n }); });
  r.post("/queue/:operationId/retry", (req, res) => {
    const n = c.processor.requeue([String(req.params.operationId)]);
    n ? res.json({ requeued: n }) : res.status(404).json({ error: "No FAILED operation with that id" });
  });

  // ---- settings ----
  r.get("/admin/settings", (_q, res) => void res.json({ duplicateMode: duplicateMode(db), timezone: c.cfg.timezone }));
  r.put("/admin/settings", (req, res) => {
    const p = z.object({ duplicateMode: z.enum(["block", "warn", "allow"]) }).safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "duplicateMode must be block, warn or allow" });
    setSetting(db, "duplicateMode", p.data.duplicateMode); audit(db, req.user!.id, "SETTING_CHANGED", `duplicateMode=${p.data.duplicateMode}`);
    res.json({ duplicateMode: p.data.duplicateMode });
  });

  const byCompany = (day: string) => {
    const occupants = day === today()
      ? all("SELECT company, COUNT(*) n FROM current_presence GROUP BY company")
      : all(`SELECT company, COUNT(*) n FROM active_operations o WHERE day=? AND outcome='APPROVED' AND state='COMPLETED' AND direction='IN'
          AND id=(SELECT MAX(id) FROM active_operations WHERE national_id=o.national_id AND day=o.day AND outcome='APPROVED' AND state='COMPLETED')
          GROUP BY company`, day);
    const on = new Map(occupants.map((row) => [row.company, row.n]));
    const counts = all("SELECT company, SUM(direction='IN') checkIns, SUM(direction='OUT') checkOuts FROM active_operations WHERE day=? AND outcome='APPROVED' AND state='COMPLETED' GROUP BY company", day);
    const companies = new Map(counts.map((row) => [row.company, row]));
    for (const company of on.keys()) {
      if (!companies.has(company)) companies.set(company, { company, checkIns: 0, checkOuts: 0 });
    }
    return [...companies.values()].sort((a, b) => String(a.company).localeCompare(String(b.company)))
      .map((row) => ({ ...row, onSite: on.get(row.company) ?? 0 }));
  };

  // ---- dashboard (aggregates from the operations table; full IDs visible to admins) ----
  r.get("/dashboard/summary", (req, res) => {
    const day = typeof req.query.date === "string" ? req.query.date : today();
    if (!DAY.test(day)) return void res.status(400).json({ error: "date must be YYYY-MM-DD" });
    const n = (dir: string, outcome: string, state = "COMPLETED") => one("SELECT COUNT(*) n FROM active_operations WHERE day=? AND direction=? AND outcome=? AND state=?", day, dir, outcome, state).n;
    const onSite = day === today() ? onSiteCount(db) : one(`SELECT COUNT(*) n FROM active_operations o WHERE day=? AND outcome='APPROVED' AND state='COMPLETED' AND direction='IN'
      AND id=(SELECT MAX(id) FROM active_operations WHERE national_id=o.national_id AND day=o.day AND outcome='APPROVED' AND state='COMPLETED')`, day).n;
    const fileInfo = (d: "IN" | "OUT") => { const f = attendancePath(c.cfg.dataDir, `${day.slice(8)}-${day.slice(5, 7)}-${day.slice(0, 4)}`, d);
      try { const s = fs.statSync(f); return { exists: true, bytes: s.size, modified: s.mtime.toISOString() }; } catch { return { exists: false }; } };
    res.json({
      day, workers: c.store.meta, workersByCompany: c.store.companySummary(),
      checkIns: n("IN", "APPROVED"), checkOuts: n("OUT", "APPROVED"), onSite,
      rejected: { notApproved: one("SELECT COUNT(*) n FROM active_operations WHERE day=? AND outcome='NOT_APPROVED'", day).n,
        notFound: one("SELECT COUNT(*) n FROM active_operations WHERE day=? AND outcome='NOT_FOUND'", day).n,
        alreadyRegistered: one("SELECT COUNT(*) n FROM active_operations WHERE day=? AND outcome='ALREADY_REGISTERED'", day).n },
      queue: { pending: one("SELECT COUNT(*) n FROM active_operations WHERE state IN ('PENDING','PROCESSING')").n + c.processor.deletionStatus().pending, failed: one("SELECT COUNT(*) n FROM active_operations WHERE state='FAILED'").n + c.processor.deletionStatus().failed },
      byCompany: byCompany(day),
      hourly: all("SELECT substr(time_text,1,2) hour, SUM(direction='IN') checkIns, SUM(direction='OUT') checkOuts FROM active_operations WHERE day=? AND outcome='APPROVED' AND state='COMPLETED' GROUP BY hour ORDER BY hour", day),
      files: { checkIn: fileInfo("IN"), checkOut: fileInfo("OUT") },
      recent: all("SELECT * FROM active_operations WHERE day=? ORDER BY id DESC LIMIT 50", day).map(operationView),
    });
  });

  r.get("/attendance/search", (req, res) => {
    const raw = typeof req.query.id === "string" ? req.query.id : "";
    const nid = normalizeNid(raw), key = isFormat14(nid) ? nid : normalizeCode(raw);
    if (!key) return void res.status(400).json({ error: "id is required" });
    const w = c.store.lookup(raw); // resolves a company code to the worker's National ID too
    const rows = all("SELECT * FROM active_operations WHERE national_id=? OR id_input=? ORDER BY id DESC LIMIT 200", w?.nationalId ?? key, key);
    res.json({ results: rows.map(operationView) });
  });

  // ---- history / export ----
  r.get("/attendance/days", (_q, res) => {
    const root = path.join(c.cfg.dataDir, "attendance");
    const days = fs.existsSync(root) ? fs.readdirSync(root).filter((d) => DATE_KEY.test(d)).sort((a, b) => isoDay(b).localeCompare(isoDay(a))) : [];
    res.json({ days });
  });
  r.get("/attendance/rejected", (req, res) => {
    const file = rejectedPath(c.cfg.dataDir);
    if (!fs.existsSync(file)) return void res.status(404).json({ error: "No rejected attempts have been saved" });
    audit(db, req.user!.id, "FILE_DOWNLOAD", "Rejected.xlsx");
    if (one("SELECT COUNT(*) n FROM excel_deletions WHERE direction='REJECTED' AND state!='COMPLETED'").n) return void res.status(409).json({ error: "Excel deletion cleanup is pending" });
    res.download(file, "Rejected.xlsx");
  });
  r.get("/attendance/file", (req, res) => {
    const date = String(req.query.date ?? ""), dir = String(req.query.dir ?? "");
    if (!DATE_KEY.test(date) || (dir !== "IN" && dir !== "OUT")) return void res.status(400).json({ error: "date=DD-MM-YYYY and dir=IN|OUT required" });
    const f = attendancePath(c.cfg.dataDir, date, dir);
    if (!fs.existsSync(f)) return void res.status(404).json({ error: "No such file" });
    audit(db, req.user!.id, "FILE_DOWNLOAD", `${date} ${dir}`);
    if (one("SELECT COUNT(*) n FROM excel_deletions WHERE direction=? AND date_key=? AND state!='COMPLETED'", dir, date).n) return void res.status(409).json({ error: "Excel deletion cleanup is pending" });
    res.download(f, path.basename(f)); // a regex-validated name: no path traversal possible
  });
  r.get("/attendance/export", async (req, res) => {
    const from = String(req.query.from ?? ""), to = String(req.query.to ?? "");
    if (!DAY.test(from) || !DAY.test(to) || from > to) return void res.status(400).json({ error: "from/to must be YYYY-MM-DD, from <= to" });
    if ((Date.parse(to) - Date.parse(from)) / 86_400_000 > 31) return void res.status(400).json({ error: "Range is limited to 32 days; export in several parts" });
    audit(db, req.user!.id, "EXPORT", `${from}..${to}`);
    res.setHeader("Content-Disposition", `attachment; filename="Attendance_${from}_to_${to}.xlsx"`)
      .type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").send(await buildExport(db, from, to));
  });

  // Regenerate a lost/corrupted daily file from the completed operations in SQLite. The damaged file is kept aside, never deleted.
  r.post("/admin/attendance/rebuild", async (req, res) => {
    const p = z.object({ date: z.string().regex(DATE_KEY), direction: z.enum(["IN", "OUT"]) }).safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "date=DD-MM-YYYY and direction=IN|OUT required" });
    const { date, direction } = p.data, file = attendancePath(c.cfg.dataDir, date, direction);
    const out = await c.processor.exclusive(async () => {
      let movedTo: string | null = null;
      if (fs.existsSync(file)) { movedTo = `${file}.replaced-${Date.now()}`; fs.renameSync(file, movedTo); }
      const ops = all("SELECT * FROM active_operations WHERE date_key=? AND direction=? AND outcome='APPROVED' AND state='COMPLETED' ORDER BY id", date, direction);
      const result = await c.writer.appendBatch(direction, date, ops.map((o) => ({ operationId: o.operation_id, nationalId: o.national_id, workerId: o.worker_id ?? "",
        name: o.name, jobTitle: o.job_title ?? "", company: o.company, status: "Valid", timeText: o.time_text, timestampIso: o.ts_iso, scanMethod: o.scan_method === "SCANNER" ? "Scanner" : o.scan_method === "MANUAL" ? "Manual ID Entry" : "Unknown", scannedBy: o.scanned_by, scannedUsername: o.scanned_username })));
      return { rebuiltRows: result.appended.length, previousFileKeptAs: movedTo ? path.basename(movedTo) : null };
    });
    audit(db, req.user!.id, "ATTENDANCE_REBUILT", `${date} ${direction} rows=${out.rebuiltRows}`);
    res.json(out);
  });

  r.post("/admin/attendance/rebuild-rejected", async (req, res) => {
    const result = await c.processor.exclusive(async () => {
      const file = rejectedPath(c.cfg.dataDir);
      let previousFileKeptAs: string | null = null;
      if (fs.existsSync(file)) {
        previousFileKeptAs = path.basename(file) + ".replaced-" + Date.now();
        fs.renameSync(file, path.join(path.dirname(file), previousFileKeptAs));
      }
      const rows = all("SELECT * FROM active_operations WHERE outcome!='APPROVED' ORDER BY id");
      const saved = await c.writer.appendRejected(rows.map((o) => ({
        operationId: o.operation_id, nationalId: o.national_id ?? "", workerId: o.worker_id || o.id_input,
        name: o.name ?? "", jobTitle: o.job_title ?? "", company: o.company ?? "", status: o.note === "NO_CHECK_IN" ? "No completed active check-in" : o.termination_reason || o.note || "",
        timeText: o.time_text, timestampIso: o.ts_iso, dateKey: o.date_key, direction: o.direction,
        scanMethod: o.scan_method === "SCANNER" ? "Scanner" : o.scan_method === "MANUAL" ? "Manual ID Entry" : "Unknown", scannedBy: o.scanned_by, scannedUsername: o.scanned_username,
        rejectionReason: o.note === "NO_CHECK_IN" ? "No Check-In" : o.outcome === "NOT_APPROVED" ? "Terminated" : o.outcome === "NOT_FOUND" ? "Not Found" : "Already Registered",
      })));
      return { rebuiltRows: saved.appended.length, previousFileKeptAs };
    });
    audit(db, req.user!.id, "REJECTED_REBUILT", String(result.rebuiltRows));
    res.json(result);
  });

  // ---- backups ----
  r.get("/admin/backups", (_q, res) => void res.json({ backups: c.backup.list() }));
  r.post("/admin/backup", async (req, res) => { const b = await c.backup.run("manual"); audit(db, req.user!.id, "BACKUP", b.name); res.status(201).json({ name: b.name }); });
  return r;
}
