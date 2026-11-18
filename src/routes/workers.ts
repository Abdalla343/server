import express, { Router } from "express";
import { authenticate, requireRole } from "../auth.js";
import { audit, type Db } from "../db.js";
import type { Config } from "../config.js";
import type { WorkerStore, LoadResult } from "../workers/store.js";

export function workersRouter(cfg: Config, db: Db, store: WorkerStore) {
  const r = Router();
  const auth = authenticate(cfg, db);
  const admin = [auth, requireRole("ADMIN")];

  // Gate lookup: full details so the officer can see who is in front of them.
  r.get("/workers/lookup", auth, (req, res) => {
    const id = typeof req.query.id === "string" ? req.query.id.slice(0, 40) : "";
    if (!id) return void res.status(400).json({ error: "Query parameter 'id' is required" });
    if (!store.ready) return void res.status(503).json({ error: "Worker database is not loaded" });
    const w = store.lookup(id, req.query.type === "employee" ? "employee" : "auto");
    if (!w) return void res.json({ found: false });
    res.json({ found: true, approved: w.status === "VALID",
      worker: { nationalId: w.nationalId, workerId: w.workerId, name: w.name, jobTitle: w.jobTitle, company: w.company,
        status: w.status, terminationReason: w.status === "TERMINATED" ? w.terminationReason : undefined } });
  });

  // "Did you mean...?" for the manual-correction screen: near misses of an unknown reading. The officer decides.
  r.get("/workers/similar", auth, (req, res) => {
    const id = typeof req.query.id === "string" ? req.query.id.slice(0, 40) : "";
    if (!id) return void res.status(400).json({ error: "Query parameter 'id' is required" });
    if (!store.ready) return void res.status(503).json({ error: "Worker database is not loaded" });
    res.json({ matches: store.similar(id).map((w) => ({ nationalId: w.nationalId, workerId: w.workerId, name: w.name, company: w.company, status: w.status })) });
  });

  r.get("/admin/workers/status", ...admin, (_q, res) =>
    void res.json({ ready: store.ready, meta: store.meta, lastError: store.lastError }));

  const reply = (res: express.Response, userId: number, action: string, result: LoadResult) => {
    if (result.ok) {
      audit(db, userId, action, `total=${result.meta.total} valid=${result.meta.valid} terminated=${result.meta.terminated}`);
      return void res.json(result);
    }
    audit(db, userId, `${action}_REJECTED`, `${result.errorCount} error(s)`);
    res.status(422).json({ ...result, message: "Validation failed. The previous worker data is still active." });
  };

  // Raw .xlsx body (Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
  r.post("/admin/workers/upload", ...admin, express.raw({ type: () => true, limit: "20mb" }), async (req, res) => {
    const buf = req.body as Buffer;
    if (!Buffer.isBuffer(buf) || buf.length < 4 || buf.subarray(0, 2).toString("latin1") !== "PK")
      return void res.status(400).json({ error: "Send the raw .xlsx file as the request body" });
    reply(res, req.user!.id, "WORKERS_UPLOADED", await store.replaceFromBuffer(buf));
  });

  r.post("/admin/workers/reload", ...admin, async (req, res) =>
    reply(res, req.user!.id, "WORKERS_RELOADED", await store.loadFromDisk()));

  // Aggregate-only summary for dashboards: no individual National IDs are returned.
  r.get("/admin/workers/summary", ...admin, (_q, res) => void res.json({ meta: store.meta, byCompany: store.companySummary() }));
  return r;
}
