import { describe, it, expect, afterEach, vi } from "vitest";
import request from "supertest";
import ExcelJS from "exceljs";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import type { Server } from "node:http";
import { createApp } from "../src/app.js";
import { loadConfig } from "../src/config.js";
import { openDb } from "../src/db.js";
import { hashPassword } from "../src/auth.js";
import { WorkerStore } from "../src/workers/store.js";
import { createContext, type Context } from "../src/context.js";
import { attendancePath } from "../src/excel/attendance.js";
import { SimulatedCrash } from "../src/queue/processor.js";
import { Scheduler } from "../src/scheduler.js";
import { createLogger } from "../src/log.js";

const nid = (i: number) => "2950101" + "01" + String(i).padStart(5, "0");
const HDR = ["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status", "reason of Terminated"];
async function workersFile(n: number): Promise<Buffer> {
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("Workers"); ws.addRow(HDR);
  for (let i = 1; i <= n; i++) ws.addRow([nid(i), String(1000 + i), `Worker ${i}`, "Welder", i % 2 ? "Alpha Co" : "Beta Co", "Valid", ""]);
  ws.addRow([nid(900), "1900", "Fired Guy", "Driver", "Alpha Co", "terminated", "No show"]);
  return Buffer.from(await wb.xlsx.writeBuffer());
}
const NOON = new Date("2026-11-14T10:05:00Z"); // 12:05 Cairo, 14-11-2026
const D = "14-11-2026";
const closers: Array<() => void> = [];
afterEach(() => { closers.splice(0).forEach((f) => f()); });

async function env(opts: { workers?: number; loadWorkers?: boolean; start?: boolean; clock?: () => Date; scanWaitMs?: number; ctxOpt?: any } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "q-"));
  const cfg = loadConfig({ dataDir: dir, scanWaitMs: opts.scanWaitMs ?? 20_000 });
  const db = openDb(dir), store = new WorkerStore(cfg);
  if (opts.loadWorkers !== false) await store.replaceFromBuffer(await workersFile(opts.workers ?? 300));
  const pw = await hashPassword("correct-horse-1");
  db.prepare("INSERT INTO users (username,password_hash,role) VALUES ('boss',?, 'ADMIN'),('gate1',?, 'USER'),('gate2',?, 'USER')").run(pw, pw, pw);
  const ctx = createContext(cfg, db, store, { clock: opts.clock ?? (() => NOON), processor: { baseBackoffMs: 20, maxBackoffMs: 50 }, ...opts.ctxOpt });
  const app = createApp(cfg, db, store, ctx);
  const server: Server = app.listen(0); closers.push(() => { server.close(); ctx.processor.stop(); });
  if (opts.start !== false) ctx.processor.start();
  const tok = async (u: string) => (await request(server).post("/api/auth/login").send({ username: u, password: "correct-horse-1" })).body.token as string;
  const t = { admin: await tok("boss"), g1: await tok("gate1"), g2: await tok("gate2") };
  const scan = (token: string, operationId: string, idNumber: string, direction: "IN" | "OUT" = "IN") =>
    request(server).post("/api/scans").set("Authorization", `Bearer ${token}`).send({ operationId, idNumber, direction });
  const get = (token: string, url: string) => request(server).get(url).set("Authorization", `Bearer ${token}`);
  return { dir, cfg, db, store, ctx, app, server, t, scan, get };
}
async function rows(file: string) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file);
  return (wb.worksheets[0]!.getSheetValues().slice(2) as any[][]).map((r) => ({ nid: r[1], name: r[3], time: r[7], op: r[8] }));
}
const op = (i: number | string) => `op-${String(i).padStart(6, "0")}`;

describe("scan flow", () => {
  it("approved scan returns 200 only after the Excel row exists", async () => {
    const e = await env();
    const r = await e.scan(e.t.g1, op(1), nid(1));
    expect(r.status).toBe(200);
    expect(r.body).toMatchObject({ state: "COMPLETED", outcome: "APPROVED", date: D, time: "12:05", worker: { name: "Worker 1", company: "Alpha Co" } });
    expect(await rows(attendancePath(e.dir, D, "IN"))).toEqual([{ nid: nid(1), name: "Worker 1", time: "12:05", op: op(1) }]);
    await e.scan(e.t.g1, op(20), "1002", "IN");
    expect((e.scan, (await e.scan(e.t.g1, op(2), "1002", "OUT")).body.worker.nationalId)).toBe(nid(2)); // by ID Card
    expect((await rows(attendancePath(e.dir, D, "OUT"))).length).toBe(1);
  });

  it("terminated / unknown / bad input: audited in SQLite, kept out of successful attendance", async () => {
    const e = await env();
    const t = await e.scan(e.t.g1, op(1), nid(900));
    expect(t.status).toBe(200); expect(t.body).toMatchObject({ outcome: "NOT_APPROVED", worker: { status: "TERMINATED", terminationReason: "No show" } });
    expect((await e.scan(e.t.g1, op(2), "29501010100999")).body).toMatchObject({ outcome: "NOT_FOUND", state: "COMPLETED" });
    expect(fs.existsSync(attendancePath(e.dir, D, "IN"))).toBe(false);
    expect(e.db.prepare("SELECT outcome FROM operations ORDER BY id").all()).toEqual([{ outcome: "NOT_APPROVED" }, { outcome: "NOT_FOUND" }]);
    expect((await e.scan(e.t.g1, "x", nid(1))).status).toBe(400);
    expect((await request(e.server).post("/api/scans").send({})).status).toBe(401);
  });

  it("returns 503 while the worker database is not loaded", async () => {
    const e = await env({ loadWorkers: false });
    expect((await e.scan(e.t.g1, op(1), nid(1))).status).toBe(503);
  });

  it("resending an operationId (sequentially or 10x at once) never creates a second row; a mismatched payload is 409", async () => {
    const e = await env();
    const rs = await Promise.all(Array.from({ length: 10 }, () => e.scan(e.t.g1, op(1), nid(1))));
    expect(rs.every((r) => r.status === 200 && r.body.outcome === "APPROVED")).toBe(true);
    expect((await e.scan(e.t.g1, op(1), nid(1))).status).toBe(200);
    expect((await e.scan(e.t.g1, op(1), nid(2))).status).toBe(409);
    expect((await e.scan(e.t.g1, op(1), nid(1), "OUT")).status).toBe(409);
    expect((await rows(attendancePath(e.dir, D, "IN"))).length).toBe(1);
  });

  it("a scan queued at 23:59 Cairo is written to THAT day's file even when processed after midnight", async () => {
    const e = await env({ start: false, scanWaitMs: 100, clock: () => new Date("2026-11-14T21:59:30Z") });
    const r = await e.scan(e.t.g1, op(1), nid(1));
    expect(r.status).toBe(202); expect(r.body.state).toBe("PENDING");
    e.ctx.processor.start(); await e.ctx.processor.waitIdle();
    expect(await rows(attendancePath(e.dir, "14-11-2026", "IN"))).toMatchObject([{ time: "23:59", op: op(1) }]);
    expect((await e.get(e.t.g1, `/api/scans/${op(1)}`)).body.state).toBe("COMPLETED");
  });
});

describe("already-registered rule", () => {
  it("block (default): IN,IN rejected; OUT,OUT rejected; IN,OUT,IN allowed; unaffected workers fine", async () => {
    const e = await env();
    const s = async (i: number, id: number, d: "IN" | "OUT") => (await e.scan(e.t.g1, op(i), nid(id), d)).body;
    expect((await s(1, 1, "IN")).outcome).toBe("APPROVED");
    const dup = await s(2, 1, "IN"); expect(dup).toMatchObject({ outcome: "ALREADY_REGISTERED", previous: "IN 12:05", state: "COMPLETED" });
    expect((await s(3, 2, "IN")).outcome).toBe("APPROVED");
    expect((await s(4, 1, "OUT")).outcome).toBe("APPROVED");
    expect(await s(5, 1, "OUT")).toMatchObject({ outcome: "NOT_APPROVED", rejectionReason: "NO_CHECK_IN" });
    expect((await s(6, 1, "IN")).outcome).toBe("APPROVED");
    expect((await rows(attendancePath(e.dir, D, "IN"))).map((r) => r.op)).toEqual([op(1), op(3), op(6)]);
  });
  it("two gates scanning the same worker at the same instant: exactly one is approved", async () => {
    const e = await env();
    const rs = await Promise.all([e.scan(e.t.g1, op(1), nid(5)), e.scan(e.t.g2, op(2), nid(5)), e.scan(e.t.g1, op(3), nid(5))]);
    expect(rs.map((r) => r.body.outcome).sort()).toEqual(["ALREADY_REGISTERED", "ALREADY_REGISTERED", "APPROVED"]);
    expect((await rows(attendancePath(e.dir, D, "IN"))).length).toBe(1);
  });
  it("warn records it with a warning; allow records silently; setting is ADMIN-only and validated", async () => {
    const e = await env();
    const put = (tk: string, v: string) => request(e.server).put("/api/admin/settings").set("Authorization", `Bearer ${tk}`).send({ duplicateMode: v });
    expect((await put(e.t.g1, "warn")).status).toBe(403); expect((await put(e.t.admin, "nope")).status).toBe(400);
    await put(e.t.admin, "warn");
    await e.scan(e.t.g1, op(1), nid(1)); const w = await e.scan(e.t.g1, op(2), nid(1));
    expect(w.body).toMatchObject({ outcome: "APPROVED", warning: "ALREADY_REGISTERED" });
    await put(e.t.admin, "allow"); expect((await e.scan(e.t.g1, op(3), nid(1))).body.warning).toBeUndefined();
    expect((await rows(attendancePath(e.dir, D, "IN"))).length).toBe(3);
  });
});

describe("concurrency", () => {
  it("300 simultaneous scans from 3 'gates' (200 IN + 100 OUT, different workers): none lost, none duplicated, few saves", async () => {
    const e = await env({ workers: 300 });
    await Promise.all(Array.from({ length: 100 }, (_, i) => e.scan(e.t.g1, op(5000 + i), nid(201 + i), "IN")));
    const spy = vi.spyOn(e.ctx.writer, "appendBatch");
    const tokens = [e.t.g1, e.t.g2, e.t.admin];
    const reqs = [
      ...Array.from({ length: 200 }, (_, i) => e.scan(tokens[i % 3]!, op(i + 1), nid(i + 1), "IN")),
      ...Array.from({ length: 100 }, (_, i) => e.scan(tokens[i % 3]!, op(1000 + i), nid(201 + i), "OUT")),
    ];
    const t0 = performance.now(); const rs = await Promise.all(reqs); const ms = performance.now() - t0;
    expect(rs.every((r) => r.status === 200 && r.body.state === "COMPLETED" && r.body.outcome === "APPROVED")).toBe(true);
    const inn = await rows(attendancePath(e.dir, D, "IN")), out = await rows(attendancePath(e.dir, D, "OUT"));
    expect(inn.length).toBe(300); expect(new Set(inn.map((r) => r.op)).size).toBe(300);
    expect(out.length).toBe(100); expect(new Set(out.map((r) => r.nid)).size).toBe(100);
    const writes = spy.mock.calls.filter((c) => c[2].length).length;
    console.log(`300 concurrent scans: ${ms.toFixed(0)} ms total, ${writes} Excel saves`);
    expect(writes).toBeLessThan(30);
    expect(e.db.prepare("SELECT COUNT(*) c FROM operations WHERE state!='COMPLETED'").get()).toEqual({ c: 0 });
  }, 60_000);
});

describe("crash recovery (process death simulated at each step)", () => {
  async function crashAndRestart(e: Awaited<ReturnType<typeof env>>, ctxOpt: any, scanIds: number[]) {
    const crashed = createContext({ ...e.cfg, scanWaitMs: 300 }, e.db, e.store, { clock: () => NOON, ...ctxOpt }); // replaces the running one's behaviour
    e.ctx.processor.stop();
    const app = createApp({ ...e.cfg, scanWaitMs: 300 }, e.db, e.store, crashed); const srv = app.listen(0); closers.push(() => srv.close());
    crashed.processor.start();
    const res = await Promise.all(scanIds.map((i) => request(srv).post("/api/scans").set("Authorization", `Bearer ${e.t.g1}`).send({ operationId: op(i), idNumber: nid(i), direction: "IN" })));
    expect(res.every((r) => r.status === 202)).toBe(true); // never reported as saved
    expect(crashed.processor.dead).toBe(true);
    // --- "restart": new DB connection, new context, same files ---
    e.db.close(); const db2 = openDb(e.dir);
    const ctx2 = createContext(e.cfg, db2, e.store, { clock: () => NOON, processor: { baseBackoffMs: 20 } });
    closers.push(() => ctx2.processor.stop());
    return { db2, ctx2 };
  }

  it("crash AFTER Excel saved, BEFORE SQLite updated: recovered, no duplicate rows", async () => {
    const e = await env({ start: false });
    const { db2, ctx2 } = await crashAndRestart(e, { processor: { beforeMarkCompleted: () => { throw new SimulatedCrash(); } } }, [1, 2, 3]);
    const stuck = (db2.prepare("SELECT COUNT(*) c FROM operations WHERE state='PROCESSING'").get() as any).c as number;
    expect(stuck).toBeGreaterThanOrEqual(1);                              // saved to Excel but never marked done
    expect(db2.prepare("SELECT COUNT(*) c FROM operations WHERE state='COMPLETED'").get()).toEqual({ c: 0 });
    expect((await rows(attendancePath(e.dir, D, "IN"))).length).toBe(stuck); // already in the file
    expect(ctx2.processor.start().recovered).toBe(stuck);
    expect(await ctx2.processor.waitIdle()).toBe(true);
    expect((await rows(attendancePath(e.dir, D, "IN"))).map((r) => r.op)).toEqual([op(1), op(2), op(3)]); // still exactly once
    expect(db2.prepare("SELECT COUNT(*) c FROM operations WHERE state='COMPLETED'").get()).toEqual({ c: 3 });
  });

  it("crash BEFORE the Excel file was replaced: original intact, recovered and written once", async () => {
    const e = await env({ start: false });
    const { db2, ctx2 } = await crashAndRestart(e, { writer: { hooks: { afterTempWrite: () => { throw new SimulatedCrash(); } } } }, [1, 2]);
    expect(fs.existsSync(attendancePath(e.dir, D, "IN"))).toBe(false); // nothing was ever saved
    expect(ctx2.processor.start().recovered).toBeGreaterThanOrEqual(1); await ctx2.processor.waitIdle();
    expect((await rows(attendancePath(e.dir, D, "IN"))).map((r) => r.op)).toEqual([op(1), op(2)]);
  });

  it("client resubmission after recovery returns the stored result and adds nothing", async () => {
    const e = await env({ start: false });
    const { db2, ctx2 } = await crashAndRestart(e, { processor: { beforeMarkCompleted: () => { throw new SimulatedCrash(); } } }, [1]);
    ctx2.processor.start(); await ctx2.processor.waitIdle();
    const app2 = createApp(e.cfg, db2, e.store, ctx2); const srv = app2.listen(0); closers.push(() => srv.close());
    const tok = (await request(srv).post("/api/auth/login").send({ username: "gate1", password: "correct-horse-1" })).body.token;
    const r = await request(srv).post("/api/scans").set("Authorization", `Bearer ${tok}`).send({ operationId: op(1), idNumber: nid(1), direction: "IN" });
    expect(r.status).toBe(200); expect(r.body.state).toBe("COMPLETED");
    expect((await rows(attendancePath(e.dir, D, "IN"))).length).toBe(1);
  });
});

describe("Excel failures", () => {
  it("locked file: scan answers 202 (queued, not lost); it completes once the lock clears", async () => {
    let locked = true; const busy = Object.assign(new Error("busy"), { code: "EBUSY" });
    const e = await env({ scanWaitMs: 300, ctxOpt: { writer: { retries: 1, rename: (a: string, b: string) => { if (locked) throw busy; fs.renameSync(a, b); } } } });
    const r = await e.scan(e.t.g1, op(1), nid(1));
    expect(r.status).toBe(202); expect(r.body.state).toMatch(/PENDING|PROCESSING/);
    locked = false;
    expect(await e.ctx.processor.waitIdle()).toBe(true);
    const g = await e.get(e.t.g1, `/api/scans/${op(1)}`); expect(g.status).toBe(200);
    expect((await rows(attendancePath(e.dir, D, "IN"))).length).toBe(1);
  });

  it("corrupted file: operation FAILED (500, visible to admin), nothing overwritten; fixing it + client resend completes it", async () => {
    const e = await env();
    await e.scan(e.t.g1, op(1), nid(1));
    const file = attendancePath(e.dir, D, "IN"), good = fs.readFileSync(file);
    fs.writeFileSync(file, "corrupted by someone");
    const r = await e.scan(e.t.g1, op(2), nid(2));
    expect(r.status).toBe(500); expect(r.body.state).toBe("FAILED");
    expect(fs.readFileSync(file).toString()).toBe("corrupted by someone");
    const st = (await e.get(e.t.admin, "/api/queue/status")).body;
    expect(st.counts.FAILED).toBe(1); expect(st.failed[0].error).toMatch(/cannot be read/);
    expect(st.failed[0].nationalId).toBe(nid(2)); // full National ID
    fs.writeFileSync(file, good);                                  // operator restores the file
    const again = await e.scan(e.t.g1, op(2), nid(2));             // officer presses "retry"
    expect(again.status).toBe(200); expect(again.body.state).toBe("COMPLETED");
    expect((await rows(file)).map((x) => x.op)).toEqual([op(1), op(2)]);
  });

  it("admin retry-failed re-queues all FAILED operations", async () => {
    const e = await env(); await e.scan(e.t.g1, op(1), nid(1));
    const file = attendancePath(e.dir, D, "IN"), good = fs.readFileSync(file); fs.writeFileSync(file, "x");
    await e.scan(e.t.g1, op(2), nid(2)); await e.scan(e.t.g1, op(3), nid(3));
    fs.writeFileSync(file, good);
    expect((await request(e.server).post("/api/queue/retry-failed").set("Authorization", `Bearer ${e.t.g1}`)).status).toBe(403);
    expect((await request(e.server).post("/api/queue/retry-failed").set("Authorization", `Bearer ${e.t.admin}`)).body).toEqual({ requeued: 2 });
    await e.ctx.processor.waitIdle();
    expect((await rows(file)).length).toBe(3);
  });
});

describe("admin: dashboard, export, files, backups, scheduler", () => {
  it("authorization: USER cannot reach admin endpoints or other users' operations", async () => {
    const e = await env(); await e.scan(e.t.g1, op(1), nid(1));
    for (const url of ["/api/dashboard/summary", "/api/queue/status", "/api/attendance/days", "/api/attendance/export?from=2026-11-14&to=2026-11-14", "/api/admin/backups", "/api/attendance/search?id=1"])
      expect((await e.get(e.t.g1, url)).status, url).toBe(403);
    expect((await request(e.server).get("/api/dashboard/summary")).status).toBe(401);
    expect((await e.get(e.t.g2, `/api/scans/${op(1)}`)).status).toBe(404);
    expect((await e.get(e.t.admin, `/api/scans/${op(1)}`)).status).toBe(200);
  });

  it("dashboard aggregates, displays full IDs, and reports file status", async () => {
    const e = await env();
    await e.scan(e.t.g1, op(1), nid(1)); await e.scan(e.t.g1, op(2), nid(2)); await e.scan(e.t.g1, op(3), nid(1), "OUT");
    await e.scan(e.t.g1, op(4), nid(900)); await e.scan(e.t.g1, op(5), "29501010100999"); await e.scan(e.t.g1, op(6), nid(2));
    const d = (await e.get(e.t.admin, "/api/dashboard/summary?date=2026-11-14")).body;
    expect(d).toMatchObject({ checkIns: 2, checkOuts: 1, onSite: 1, rejected: { notApproved: 1, notFound: 1, alreadyRegistered: 1 },
      queue: { pending: 0, failed: 0 }, files: { checkIn: { exists: true }, checkOut: { exists: true } }, workers: { total: 301 } });
    expect(d.byCompany).toEqual([{ company: "Alpha Co", checkIns: 1, checkOuts: 1, onSite: 0 }, { company: "Beta Co", checkIns: 1, checkOuts: 0, onSite: 1 }]);
    expect(d.hourly).toEqual([{ hour: "12", checkIns: 2, checkOuts: 1 }]);
    expect(JSON.stringify(d)).toContain(nid(1)); expect(d.recent[0].nationalId).toBe(nid(2));
    expect((await e.get(e.t.admin, "/api/dashboard/summary?date=bad")).status).toBe(400);
    const s = (await e.get(e.t.admin, `/api/attendance/search?id=${nid(1)}`)).body.results; expect(s.length).toBe(2);
  });

  it("export builds a consolidated workbook; file download is traversal-safe", async () => {
    const e = await env(); await e.scan(e.t.g1, op(1), nid(1)); await e.scan(e.t.g1, op(2), nid(900));
    const x = await e.get(e.t.admin, "/api/attendance/export?from=2026-11-14&to=2026-11-14").buffer(true).parse((res, cb) => { const c: Buffer[] = []; res.on("data", (d) => c.push(d)); res.on("end", () => cb(null, Buffer.concat(c))); });
    expect(x.status).toBe(200);
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(x.body);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["CheckIn", "CheckOut", "Rejected"]);
    expect(wb.getWorksheet("CheckIn")!.getRow(2).getCell(2).value).toBe(nid(1));
    expect(wb.getWorksheet("Rejected")!.getRow(2).getCell(10).value).toBe("NOT_APPROVED");
    expect(wb.getWorksheet("Rejected")!.getRow(1).getCell(12).value).toBe("Scanned By");
    expect(wb.getWorksheet("Rejected")!.getRow(2).getCell(12).value).toBe("gate1");
    expect(wb.getWorksheet("Rejected")!.getRow(2).getCell(14).value).toBe("Terminated");
    expect(wb.getWorksheet("CheckIn")!.getRow(2).getCell(10).value).toBe("gate1");
    expect((await e.get(e.t.admin, "/api/attendance/export?from=2026-01-01&to=2026-11-14")).status).toBe(400);
    expect((await e.get(e.t.admin, `/api/attendance/file?date=${D}&dir=IN`)).status).toBe(200);
    expect((await e.get(e.t.admin, "/api/attendance/file?date=../../workers&dir=IN")).status).toBe(400);
    expect((await e.get(e.t.admin, "/api/attendance/days")).body).toEqual({ days: [D] });
  });

  it("backups contain app.db, Workers.xlsx and recent attendance; old ones are pruned", async () => {
    const e = await env(); await e.scan(e.t.g1, op(1), nid(1));
    const b = (await request(e.server).post("/api/admin/backup").set("Authorization", `Bearer ${e.t.admin}`)).body.name as string;
    const dir = path.join(e.dir, "backups", "snapshots", b);
    expect(fs.readdirSync(dir).sort()).toEqual(["Workers.xlsx", "app.db", "attendance"]);
    expect(fs.existsSync(path.join(dir, "attendance", D, `${D}CheckIn.xlsx`))).toBe(true);
    for (let i = 0; i < 22; i++) fs.mkdirSync(path.join(e.dir, "backups", "snapshots", `manual-2020-01-01T00-00-${String(i).padStart(2, "0")}`));
    await e.ctx.backup.run("manual");
    expect(e.ctx.backup.list().filter((x) => x.name.startsWith("manual-")).length).toBe(20);
  });

  it("scheduler creates both day files on date change and runs the hourly + daily backups once", async () => {
    let now = new Date("2026-11-14T21:59:00Z"); // 23:59 Cairo
    const e = await env({ clock: () => now });
    const sch = new Scheduler(e.cfg, e.db, e.ctx.writer, e.ctx.processor, e.ctx.backup, createLogger(null), () => now);
    await sch.tick();
    expect(fs.existsSync(attendancePath(e.dir, "14-11-2026", "OUT"))).toBe(true);
    expect(e.ctx.backup.list().map((b) => b.name.split("-")[0]).sort()).toEqual(["daily", "hourly"]); // 23:59 >= 23:30
    now = new Date("2026-11-14T22:01:00Z"); await sch.tick(); // 00:01 on 15-11
    expect(fs.existsSync(attendancePath(e.dir, "15-11-2026", "IN"))).toBe(true);
    expect(fs.existsSync(attendancePath(e.dir, "15-11-2026", "OUT"))).toBe(true);
    expect(e.ctx.backup.list().length).toBe(2); // hourly interval not elapsed, daily not repeated
  });
});

describe("rebuild from SQLite", () => {
  it("regenerates a destroyed daily file from completed operations; the damaged file is kept aside", async () => {
    const e = await env();
    for (let i = 1; i <= 5; i++) await e.scan(e.t.g1, op(i), nid(i));
    const file = attendancePath(e.dir, D, "IN"); fs.writeFileSync(file, "destroyed");
    const bad = await request(e.server).post("/api/admin/attendance/rebuild").set("Authorization", `Bearer ${e.t.g1}`).send({ date: D, direction: "IN" });
    expect(bad.status).toBe(403);
    const r = await request(e.server).post("/api/admin/attendance/rebuild").set("Authorization", `Bearer ${e.t.admin}`).send({ date: D, direction: "IN" });
    expect(r.body.rebuiltRows).toBe(5); expect(r.body.previousFileKeptAs).toMatch(/replaced-/);
    expect((await rows(file)).map((x) => x.op)).toEqual([1, 2, 3, 4, 5].map(op));
    expect((await e.scan(e.t.g1, op(6), nid(6))).status).toBe(200); // normal scanning continues
    expect((await rows(file)).length).toBe(6);
  });
});

describe("scanning company codes", () => {
  it("a scanned code (any case / Arabic digits) is recorded against the worker, and findable by code or National ID", async () => {
    const e = await env();
    const rowsX: any[][] = [["IR1234", "Ahmed Code"], ["ZTC535", "Omar Code"]];
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("Workers"); ws.addRow(HDR);
    rowsX.forEach(([code, name], i) => ws.addRow([nid(i + 1), code, name, "Welder", "Alpha Co", "Valid", ""]));
    expect((await e.store.replaceFromBuffer(Buffer.from(await wb.xlsx.writeBuffer()))).ok).toBe(true);
    const a = await e.scan(e.t.g1, op(1), "ir١٢٣٤");
    expect(a.status).toBe(200); expect(a.body).toMatchObject({ outcome: "APPROVED", worker: { name: "Ahmed Code", workerId: "IR1234", nationalId: nid(1) } });
    expect((await e.scan(e.t.g1, op(2), "Ir-1234")).body.outcome).toBe("ALREADY_REGISTERED");           // same worker, other spelling
    expect((await e.scan(e.t.g1, op(3), nid(1))).body.outcome).toBe("ALREADY_REGISTERED");              // same worker by National ID
    expect((await e.scan(e.t.g1, op(4), "NOPE99")).body).toMatchObject({ outcome: "NOT_FOUND" });
    const f = await rows(attendancePath(e.dir, D, "IN")); expect(f).toEqual([{ nid: nid(1), name: "Ahmed Code", time: "12:05", op: op(1) }]);
    const L = await e.get(e.t.g1, "/api/workers/lookup?id=zTc%2D535"); expect(L.body).toMatchObject({ found: true, approved: true, worker: { name: "Omar Code" } });
    const s1 = (await e.get(e.t.admin, "/api/attendance/search?id=ir1234")).body.results, s2 = (await e.get(e.t.admin, `/api/attendance/search?id=${nid(1)}`)).body.results;
    expect(s1.length).toBe(4 - 1); expect(s2.length).toBe(s1.length); // 3 scans of that worker: 1 approved + 2 repeats
  });
});

describe("operator attribution, rejection files and live presence", () => {
  it("captures the authenticated display name, ignores forged identity and keeps snapshots after user deletion", async () => {
    const e = await env({ workers: 3 });
    e.db.prepare("UPDATE users SET display_name='Ahmed Gate One' WHERE username='gate1'").run();
    const r = await request(e.server).post("/api/scans").set("Authorization", `Bearer ${e.t.g1}`)
      .send({ operationId: op(101), idNumber: "1001", direction: "IN", scannedBy: "Forged", username: "boss", displayName: "Forged" });
    expect(r.status).toBe(200);
    expect(r.body).toMatchObject({ scannedBy: "Ahmed Gate One", scannedUsername: "gate1" });
    const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(attendancePath(e.dir, D, "IN"));
    const ws = wb.worksheets[0]!;
    expect(ws.getRow(2).getCell(10).value).toBe("Ahmed Gate One");
    expect(ws.getRow(2).getCell(11).value).toBe("gate1");
    await e.scan(e.t.g2, op(103), "1002", "IN");
    await e.scan(e.t.g2, op(102), "1002", "OUT");
    const out = new ExcelJS.Workbook(); await out.xlsx.readFile(attendancePath(e.dir, D, "OUT"));
    expect(out.worksheets[0]!.getRow(2).getCell(10).value).toBe("gate2");
    expect(out.worksheets[0]!.getRow(2).getCell(11).value).toBe("gate2");
    const userId = (e.db.prepare("SELECT id FROM users WHERE username='gate1'").get() as any).id;
    expect((await request(e.server).delete(`/api/admin/users/${userId}`).set("Authorization", `Bearer ${e.t.admin}`)).status).toBe(200);
    expect((await e.get(e.t.g1, "/api/auth/me")).status).toBe(401);
    const presence = (await e.get(e.t.admin, "/api/admin/onsite")).body;
    expect(presence.employees[0]).toMatchObject({ employeeId: "1001", name: "Worker 1", scannedBy: "Ahmed Gate One", scannedUsername: "gate1", nationalId: nid(1) });
    expect((await e.get(e.t.admin, "/api/dashboard/summary")).body.recent.find((row: any) => row.operationId === op(101)).scannedBy).toBe("Ahmed Gate One");
  });

  it("writes both rejected directions to one separate workbook, preserves reasons, and retries idempotently", async () => {
    const e = await env({ workers: 2 });
    await e.scan(e.t.g1, op(1), "1900", "IN");
    await e.scan(e.t.g2, op(2), "UNKNOWN123", "OUT");
    await Promise.all(Array.from({ length: 4 }, () => e.scan(e.t.g2, op(2), "UNKNOWN123", "OUT")));
    const file = path.join(e.dir, "attendance", "rejected", "Rejected.xlsx");
    const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file);
    const ws = wb.worksheets[0]!;
    expect(wb.worksheets).toHaveLength(1); expect(ws.actualRowCount).toBe(3);
    expect(ws.getRow(2).values).toEqual(expect.arrayContaining(["1900", "Fired Guy", "Terminated", "Check-In", "No show", "gate1"]));
    expect(ws.getRow(3).values).toEqual(expect.arrayContaining(["UNKNOWN123", "Not Found", "Check-Out", "gate2"]));
    expect(fs.existsSync(attendancePath(e.dir, D, "IN"))).toBe(false);
    expect(fs.existsSync(attendancePath(e.dir, D, "OUT"))).toBe(false);
    expect((await e.get(e.t.admin, "/api/admin/onsite")).body.total).toBe(0);
    expect((await e.get(e.t.g1, "/api/attendance/rejected")).status).toBe(403);
    expect((await e.get(e.t.admin, "/api/attendance/rejected")).status).toBe(200);
    const snapshot = await e.ctx.backup.run("manual");
    expect(fs.existsSync(path.join(snapshot.dir, "attendance", "rejected", "Rejected.xlsx"))).toBe(true);
  });

  it("retries a rejected workbook failure without approving the employee or duplicating the row", async () => {
    const e = await env({ workers: 2 });
    await e.scan(e.t.g1, op(1), "UNKNOWN001");
    const file = path.join(e.dir, "attendance", "rejected", "Rejected.xlsx"), good = fs.readFileSync(file);
    fs.writeFileSync(file, "corrupt");
    const failed = await e.scan(e.t.g1, op(2), "1900", "OUT");
    expect(failed.status).toBe(500); expect(failed.body).toMatchObject({ state: "FAILED", outcome: "NOT_APPROVED" });
    expect(fs.readFileSync(file).toString()).toBe("corrupt");
    fs.writeFileSync(file, good);
    expect((await e.scan(e.t.g1, op(2), "1900", "OUT")).body.state).toBe("COMPLETED");
    expect((await rows(file)).length).toBe(2);
    expect((await e.get(e.t.admin, "/api/admin/onsite")).body.total).toBe(0);
  });

  it("shows completed entries only, retains overnight occupants on a new DB connection, and removes successful exits", async () => {
    let now = new Date("2026-11-14T21:50:00Z");
    const e = await env({ workers: 3, clock: () => now, start: false, scanWaitMs: 5 });
    expect((await e.scan(e.t.g1, op(1), "1001")).status).toBe(202);
    expect((await e.get(e.t.admin, "/api/admin/onsite")).body.total).toBe(0);
    e.ctx.processor.start(); expect(await e.ctx.processor.waitIdle()).toBe(true);
    now = new Date("2026-11-15T00:50:00Z");
    const on = (await e.get(e.t.admin, "/api/admin/onsite?search=Alpha&limit=1")).body;
    expect(on).toMatchObject({ total: 1, filtered: 1 });
    expect(on.employees[0]).toMatchObject({ employeeId: "1001", secondsInside: 10800, checkInTime: "23:50", scannedBy: "gate1" });
    const summary = (await e.get(e.t.admin, "/api/dashboard/summary")).body;
    expect(summary.onSite).toBe(1);
    expect(summary.byCompany).toContainEqual({ company: "Alpha Co", checkIns: 0, checkOuts: 0, onSite: 1 });
    const reopened = openDb(e.dir);
    expect((reopened.prepare("SELECT COUNT(*) n FROM current_presence").get() as any).n).toBe(1);
    reopened.close();
    expect((await e.get(e.t.admin, "/api/admin/onsite?search=nothing")).body).toMatchObject({ total: 1, filtered: 0, employees: [] });
    const out = await e.scan(e.t.g2, op(2), "1001", "OUT");
    if (out.status === 202) expect(await e.ctx.processor.waitIdle()).toBe(true);
    expect((await e.get(e.t.admin, "/api/admin/onsite")).body.total).toBe(0);
    expect((await e.get(e.t.g1, "/api/admin/onsite")).status).toBe(403);
  });

  it("handles simultaneous gates and keeps the original arrival when duplicate INs are allowed", async () => {
    let now = NOON;
    const e = await env({ workers: 3, clock: () => now });
    const results = await Promise.all([e.scan(e.t.g1, op(1), "1001"), e.scan(e.t.g2, op(2), "1001"), e.scan(e.t.g2, op(3), "1002")]);
    expect(results.filter((r) => r.body.outcome === "APPROVED")).toHaveLength(2);
    expect((await e.get(e.t.admin, "/api/admin/onsite")).body.total).toBe(2);
    await request(e.server).put("/api/admin/settings").set("Authorization", `Bearer ${e.t.admin}`).send({ duplicateMode: "allow" });
    now = new Date(NOON.getTime() + 3600000);
    await e.scan(e.t.g2, op(4), "1001");
    const presence = (await e.get(e.t.admin, "/api/admin/onsite?search=1001")).body;
    expect(presence).toMatchObject({ total: 2, filtered: 1 });
    expect(presence.employees[0]).toMatchObject({ secondsInside: 3600, scannedUsername: "gate1" });
    expect((await e.scan(e.t.g2, op(1), "1001")).status).toBe(403);
  });

  it("resolves a 14-digit employee card in its own namespace and rejects changing type on retry", async () => {
    const e = await env({ workers: 2 });
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("Workers");
    ws.addRow(HDR);
    ws.addRow([nid(1), "A123", "First", "Welder", "Alpha", "Valid"]);
    ws.addRow([nid(2), nid(1), "Second", "Welder", "Alpha", "Valid"]);
    expect((await e.store.replaceFromBuffer(Buffer.from(await wb.xlsx.writeBuffer()))).ok).toBe(true);
    expect((await e.get(e.t.g1, `/api/workers/lookup?id=${nid(1)}&type=employee`)).body.worker.name).toBe("Second");
    expect((await e.get(e.t.g1, `/api/workers/lookup?id=${nid(1)}`)).body.worker.name).toBe("First");
    const response = await request(e.server).post("/api/scans").set("Authorization", `Bearer ${e.t.g1}`)
      .send({ operationId: op(1), idNumber: nid(1), idType: "employee", direction: "IN" });
    expect(response.body.worker.name).toBe("Second");
    expect((await e.scan(e.t.g1, op(1), nid(1))).status).toBe(409);
  });
});

describe("check-out registration, scan method and operation deletion", () => {
  it("blocks first check-out, check-out after failed or pending entry, and concurrent repeated exits", async () => {
    const e = await env({ workers: 3 });
    const first = await e.scan(e.t.g1, op(701), "1001", "OUT");
    expect(first.body).toMatchObject({ outcome: "NOT_APPROVED", rejectionReason: "NO_CHECK_IN", state: "COMPLETED" });
    expect(fs.existsSync(attendancePath(e.dir, D, "OUT"))).toBe(false);
    expect((await e.scan(e.t.g1, op(702), "1001")).body.outcome).toBe("APPROVED");
    const exits = await Promise.all([e.scan(e.t.g1, op(703), "1001", "OUT"), e.scan(e.t.g2, op(704), "1001", "OUT")]);
    expect(exits.filter((r) => r.body.outcome === "APPROVED")).toHaveLength(1);
    expect(exits.find((r) => r.body.outcome !== "APPROVED")!.body.rejectionReason).toBe("NO_CHECK_IN");
    expect((await rows(attendancePath(e.dir, D, "OUT")))).toHaveLength(1);
    const pending = await env({ workers: 1, start: false, scanWaitMs: 5 });
    await pending.scan(pending.t.g1, op(705), "1001");
    const blocked = await pending.scan(pending.t.g2, op(706), "1001", "OUT");
    expect(blocked.body.outcome).toBe("NOT_APPROVED");
    pending.db.prepare("UPDATE operations SET state='FAILED' WHERE operation_id=?").run(op(705));
    expect((await pending.scan(pending.t.g2, op(707), "1001", "OUT")).body.rejectionReason).toBe("NO_CHECK_IN");
  });

  it("writes scanner and manual source to every workbook and the date-range export", async () => {
    const e = await env({ workers: 2 });
    const send = (id: string, value: string, direction: string, scanMethod: string) =>
      request(e.server).post("/api/scans").set("Authorization", `Bearer ${e.t.g1}`)
        .send({ operationId: id, idNumber: value, direction, scanMethod });
    await send(op(710), "1001", "IN", "SCANNER");
    await send(op(711), "1001", "OUT", "MANUAL");
    await send(op(712), "UNKNOWN7", "IN", "MANUAL");
    for (const [file, expected] of [[attendancePath(e.dir, D, "IN"), "Scanner"], [attendancePath(e.dir, D, "OUT"), "Manual ID Entry"],
      [path.join(e.dir, "attendance", "rejected", "Rejected.xlsx"), "Manual ID Entry"]]) {
      const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file!);
      const ws = wb.worksheets[0]!;
      const values = ws.getRow(1).values as string[];
      const column = values.indexOf("Operation Scan");
      expect(column).toBeGreaterThan(0);
      expect(ws.getRow(2).getCell(column).value).toBe(expected);
    }
    const list = (await e.get(e.t.admin, "/api/admin/operations")).body.operations;
    expect(list.find((row: any) => row.operationId === op(710)).scanMethod).toBe("SCANNER");
    expect((await send(op(710), "1001", "IN", "MANUAL")).status).toBe(409);
    const exported = await e.get(e.t.admin, "/api/attendance/export?from=2026-11-14&to=2026-11-14").buffer(true).parse((res, cb) => {
      const buffers: Buffer[] = []; res.on("data", (b) => buffers.push(b)); res.on("end", () => cb(null, Buffer.concat(buffers)));
    });
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(exported.body);
    for (const name of ["CheckIn", "CheckOut", "Rejected"])
      expect((wb.getWorksheet(name)!.getRow(1).values as string[])).toContain("Operation Scan");
  });

  it("deletes selected operations from Excel, counts, search and presence, while preventing retry resurrection", async () => {
    const e = await env({ workers: 2 });
    await e.scan(e.t.g1, op(720), "1001");
    await e.scan(e.t.g2, op(721), "1002");
    expect((await request(e.server).post("/api/admin/operations/delete").set("Authorization", `Bearer ${e.t.g1}`).send({ all: true })).status).toBe(403);
    const result = await request(e.server).post("/api/admin/operations/delete").set("Authorization", `Bearer ${e.t.admin}`)
      .send({ operationIds: [op(720)] });
    expect(result.body).toMatchObject({ deleted: 1, excelPending: { pending: 0, failed: 0 } });
    expect((await rows(attendancePath(e.dir, D, "IN"))).map((row) => row.op)).toEqual([op(721)]);
    const summary = (await e.get(e.t.admin, "/api/dashboard/summary")).body;
    expect(summary).toMatchObject({ checkIns: 1, onSite: 1 });
    expect(summary.recent.some((row: any) => row.operationId === op(720))).toBe(false);
    expect((await e.get(e.t.admin, "/api/attendance/search?id=1001")).body.results).toEqual([]);
    expect((await e.scan(e.t.g1, op(720), "1001")).status).toBe(410);
    expect((await e.scan(e.t.g1, op(722), "1001", "OUT")).body.rejectionReason).toBe("NO_CHECK_IN");
    expect((await e.scan(e.t.g1, op(723), "1001")).body.outcome).toBe("APPROVED");
  });

  it("deletes all dates and all three workbook categories, with a durable retry after a file failure", async () => {
    let now = NOON;
    const e = await env({ workers: 2, clock: () => now });
    await e.scan(e.t.g1, op(730), "1001");
    await e.scan(e.t.g1, op(731), "1001", "OUT");
    now = new Date(NOON.getTime() + 86400000);
    await e.scan(e.t.g2, op(732), "1002");
    await e.scan(e.t.g1, op(733), "UNKNOWN7");
    const file = attendancePath(e.dir, D, "OUT"), original = fs.readFileSync(file);
    fs.writeFileSync(file, "corrupt");
    const deleted = await request(e.server).post("/api/admin/operations/delete").set("Authorization", `Bearer ${e.t.admin}`).send({ all: true });
    expect(deleted.body.deleted).toBe(4); expect(deleted.body.excelPending.failed).toBe(1);
    expect((await e.get(e.t.admin, "/api/admin/operations")).body.total).toBe(0);
    expect((await e.get(e.t.admin, "/api/admin/onsite")).body.total).toBe(0);
    expect((await e.get(e.t.admin, `/api/attendance/file?date=${D}&dir=OUT`)).status).toBe(409);
    fs.writeFileSync(file, original);
    e.ctx.processor.requeue();
    expect(await e.ctx.processor.waitIdle()).toBe(true);
    for (const f of [attendancePath(e.dir, D, "IN"), file, attendancePath(e.dir, "15-11-2026", "IN"),
      path.join(e.dir, "attendance", "rejected", "Rejected.xlsx")]) expect(await rows(f)).toHaveLength(0);
    const reopened = openDb(e.dir);
    expect((reopened.prepare("SELECT COUNT(*) n FROM active_operations").get() as any).n).toBe(0);
    expect((reopened.prepare("SELECT COUNT(*) n FROM operations WHERE deleted_at IS NOT NULL").get() as any).n).toBe(4);
    reopened.close();
    expect((await request(e.server).post("/api/admin/operations/delete").set("Authorization", `Bearer ${e.t.admin}`).send({})).status).toBe(400);
  });
});

describe("deletion recovery and overnight exit", () => {
  it("allows yesterday's completed entry to be checked out today, regardless of duplicate settings", async () => {
    let now = NOON;
    const e = await env({ workers: 1, clock: () => now });
    await e.scan(e.t.g1, op(800), "1001");
    now = new Date(NOON.getTime() + 86400000);
    await request(e.server).put("/api/admin/settings").set("Authorization", `Bearer ${e.t.admin}`).send({ duplicateMode: "allow" });
    expect((await e.scan(e.t.g2, op(801), "1001", "OUT")).body.outcome).toBe("APPROVED");
    expect((await e.scan(e.t.g1, op(802), "1001", "OUT")).body).toMatchObject({ outcome: "NOT_APPROVED", rejectionReason: "NO_CHECK_IN" });
  });

  it("replays a durable workbook cleanup after restart, without rewriting surviving records", async () => {
    const e = await env({ workers: 2 });
    await e.scan(e.t.g1, op(810), "1001"); await e.scan(e.t.g2, op(811), "1002");
    await e.ctx.processor.stop();
    // Represents a crash immediately after the admin deletion transaction commits.
    e.db.transaction(() => {
      e.db.prepare("UPDATE operations SET deleted_at=? WHERE operation_id=?").run(NOON.toISOString(), op(810));
      e.db.prepare("INSERT INTO excel_deletions(file_key,direction,date_key) VALUES(?,?,?)").run("IN/" + D, "IN", D);
    })();
    expect((await rows(attendancePath(e.dir, D, "IN")))).toHaveLength(2);
    const reopened = openDb(e.dir);
    const restarted = createContext(e.cfg, reopened, e.store);
    restarted.processor.start(); expect(await restarted.processor.waitIdle()).toBe(true); await restarted.processor.stop();
    expect((await rows(attendancePath(e.dir, D, "IN"))).map((row) => row.op)).toEqual([op(811)]);
    expect((reopened.prepare("SELECT COUNT(*) n FROM current_presence").get() as any).n).toBe(1);
    expect(restarted.processor.deletionStatus()).toEqual({ pending: 0, failed: 0 });
    reopened.close();
  });

  it("cancels a pending scan on deletion and cannot later add it to Excel", async () => {
    const e = await env({ workers: 1, start: false, scanWaitMs: 5 });
    await e.scan(e.t.g1, op(820), "1001");
    const deleted = await request(e.server).post("/api/admin/operations/delete").set("Authorization", `Bearer ${e.t.admin}`)
      .send({ operationIds: [op(820)] });
    expect(deleted.body.deleted).toBe(1);
    e.ctx.processor.start(); expect(await e.ctx.processor.waitIdle()).toBe(true);
    expect(fs.existsSync(attendancePath(e.dir, D, "IN"))).toBe(false);
    expect((await e.scan(e.t.g1, op(820), "1001")).status).toBe(410);
  });
});
