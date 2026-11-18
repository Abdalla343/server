import { describe, it, expect, beforeEach } from "vitest";
import request from "supertest";
import ExcelJS from "exceljs";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createApp } from "../src/app.js";
import { loadConfig } from "../src/config.js";
import { openDb } from "../src/db.js";
import { hashPassword } from "../src/auth.js";
import { WorkerStore } from "../src/workers/store.js";
import { parseWorkersBuffer } from "../src/workers/parse.js";
import { maskNid, nidSemanticProblem, normalizeNid } from "../src/nationalId.js";

const HEADERS = ["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status", "reason of Terminated"];
const nid = (i: number) => "2950101" + "01" + String(i).padStart(5, "0");
type Row = (string | number)[];
const good = (i: number, status = "Valid"): Row => [nid(i), String(1000 + i), `Worker ${i}`, "Welder", i % 2 ? "Alpha Co" : "Beta Co", status, status === "Valid" ? "" : "Absent"];

async function xlsx(rows: Row[], headers: string[] = HEADERS, extraSheet = false): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  if (extraSheet) wb.addWorksheet("Notes").addRow(["nothing here"]);
  const ws = wb.addWorksheet("Workers");
  ws.addRow(headers); rows.forEach((r) => ws.addRow(r));
  return Buffer.from(await wb.xlsx.writeBuffer());
}
async function env() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "wk-"));
  const cfg = loadConfig({ dataDir: dir }); const db = openDb(dir); const store = new WorkerStore(cfg);
  const pw = await hashPassword("correct-horse-1");
  db.prepare("INSERT INTO users (username,password_hash,role) VALUES ('boss',?, 'ADMIN'),('gate',?, 'USER')").run(pw, pw);
  const app = createApp(cfg, db, store);
  const tok = async (u: string) => (await request(app).post("/api/auth/login").send({ username: u, password: "correct-horse-1" })).body.token as string;
  return { dir, db, store, app, admin: await tok("boss"), user: await tok("gate") };
}

describe("national id helpers", () => {
  it("normalizes Arabic-Indic digits and validates semantics", () => {
    expect(normalizeNid("٢٩٥٠١٠١٠١٠٠٠٠١ ")).toBe(nid(1));
    expect(nidSemanticProblem(nid(1))).toBeNull();
    expect(nidSemanticProblem("29513320100001")).toBe("invalid birth date");
    expect(nidSemanticProblem("19501010100001")).toMatch(/century/);
    expect(nidSemanticProblem("29501019900001")).toMatch(/governorate/);
    expect(maskNid(nid(1))).toBe("295********001");
  });
});

describe("parser", () => {
  it("reads the real headers, numeric/Arabic cells, and skips blank rows", async () => {
    const buf = await xlsx([good(1), [Number(nid(2)), 7, "Ali", "", "Alpha Co", "terminated", "Fired"], [], ["٢٩٥٠١٠١٠١٠٠٠٠٣", "", "Omar", "", "Alpha Co", "VALID", ""]], HEADERS, true);
    const r = await parseWorkersBuffer(buf);
    expect(r.errors).toEqual([]);
    expect(r.sheet).toBe("Workers");
    expect(r.workers.map((w) => [w.nationalId, w.status])).toEqual([[nid(1), "VALID"], [nid(2), "TERMINATED"], [nid(3), "VALID"]]);
    expect(r.workers[1]!.workerId).toBe("7");
    expect(r.workers[1]!.terminationReason).toBe("Fired");
  });
  it("reports row-level errors without stopping at the first", async () => {
    const r = await parseWorkersBuffer(await xlsx([good(1), good(1), ["123", "", "X", "", "Co", "Valid", ""], good(4, "Maybe"), [nid(5), "1001", "Dup card", "", "Co", "Valid", ""], ["", "", "", "", "", "", ""]]));
    const msgs = r.errors.map((e) => `${e.row}:${e.field}`);
    expect(msgs).toEqual(expect.arrayContaining(["3:nationalId", "4:nationalId", "5:status", "6:workerId"]));
  });
  it("rejects missing required columns and non-xlsx input", async () => {
    expect((await parseWorkersBuffer(await xlsx([good(1)], ["National ID Number", "Name"]))).errors[0]!.message).toMatch(/company, status/);
    expect((await parseWorkersBuffer(Buffer.from("hello"))).errors[0]!.message).toMatch(/Not a readable/);
  });
  it("loads semantically odd IDs with a warning (they may still be scanned)", async () => {
    const r = await parseWorkersBuffer(await xlsx([["29513320100001", "", "N", "", "Co", "Valid", ""]]));
    expect(r.errors).toEqual([]); expect(r.warnings[0]!.message).toMatch(/invalid birth date/);
  });
  it("handles 3,200 workers quickly, with sub-millisecond lookups", async () => {
    const buf = await xlsx(Array.from({ length: 3200 }, (_, i) => good(i + 1, i % 10 === 0 ? "terminated" : "Valid")));
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "perf-")); const store = new WorkerStore({ dataDir: dir });
    const t0 = performance.now(); const r = await store.replaceFromBuffer(buf); const load = performance.now() - t0;
    expect(r.ok && r.meta.total).toBe(3200);
    const t1 = performance.now(); for (let i = 0; i < 100_000; i++) store.lookup(nid((i % 3200) + 1)); const per = (performance.now() - t1) / 100_000;
    console.log(`3200 rows: load+validate+write ${load.toFixed(0)} ms, lookup ${(per * 1000).toFixed(2)} µs`);
    expect(per).toBeLessThan(0.05);
  });
});

describe("store safety", () => {
  it("keeps the last valid data when a bad file is uploaded or the disk file is corrupted", async () => {
    const { store } = await env();
    expect((await store.replaceFromBuffer(await xlsx([good(1), good(2)]))).ok).toBe(true);
    const before = fs.readFileSync(store.file);
    expect((await store.replaceFromBuffer(await xlsx([good(9), good(9)]))).ok).toBe(false);
    expect(fs.readFileSync(store.file).equals(before)).toBe(true);       // file untouched
    expect(store.lookup(nid(1))?.name).toBe("Worker 1");                  // index untouched
    expect(store.lookup(nid(9))).toBeUndefined();
    fs.writeFileSync(store.file, "corrupted");                            // someone broke the file on disk
    const r = await store.loadFromDisk();
    expect(r.ok).toBe(false); expect(store.lookup(nid(1))).toBeDefined(); expect(store.lastError).not.toBeNull();
  });
  it("backs up the old file on replace, and reload picks up external edits", async () => {
    const { store } = await env();
    await store.replaceFromBuffer(await xlsx([good(1)]));
    await store.replaceFromBuffer(await xlsx([good(2)]));
    expect(fs.readdirSync(store.backupDir).filter((f) => f.startsWith("Workers_")).length).toBe(1);
    expect(store.lookup(nid(1))).toBeUndefined(); expect(store.lookup(nid(2))).toBeDefined();
    fs.writeFileSync(store.file, await xlsx([good(3, "terminated")]));
    expect((await store.loadFromDisk()).ok).toBe(true);
    expect(store.lookup("1003")?.status).toBe("TERMINATED"); // by ID Card
    expect(fs.existsSync(store.file + ".tmp")).toBe(false);
  });
  it("serializes concurrent uploads: final index always equals final file", async () => {
    const { store } = await env();
    const [a, b, c] = await Promise.all([xlsx([good(1)]), xlsx([good(2)]), xlsx([good(3)])]);
    await Promise.all([store.replaceFromBuffer(a), store.replaceFromBuffer(b), store.replaceFromBuffer(c)]);
    expect(store.lookup(nid(3))).toBeDefined(); expect(store.lookup(nid(1))).toBeUndefined();
    const reloaded = new WorkerStore({ dataDir: path.dirname(path.dirname(store.file)) });
    await reloaded.loadFromDisk(); expect(reloaded.meta!.sha256).toBe(store.meta!.sha256);
  });
});

describe("worker API", () => {
  let e: Awaited<ReturnType<typeof env>>;
  beforeEach(async () => { e = await env(); });
  const up = async (rows: Row[], token = e.admin) => request(e.app).post("/api/admin/workers/upload").set("Authorization", `Bearer ${token}`)
    .set("Content-Type", "application/octet-stream").send(await xlsx(rows));

  it("enforces roles and requires a loaded database", async () => {
    expect((await request(e.app).get("/api/workers/lookup?id=1")).status).toBe(401);
    expect((await up([good(1)], e.user)).status).toBe(403);
    expect((await request(e.app).post("/api/admin/workers/reload").set("Authorization", `Bearer ${e.user}`)).status).toBe(403);
    expect((await request(e.app).get(`/api/workers/lookup?id=${nid(1)}`).set("Authorization", `Bearer ${e.user}`)).status).toBe(503);
  });
  it("uploads, looks up by National ID (Arabic digits too) and ID Card, and reports statuses", async () => {
    const r = await up([good(1), good(2, "terminated")]);
    expect(r.status).toBe(200); expect(r.body.meta).toMatchObject({ total: 2, valid: 1, terminated: 1 });
    const L = (id: string) => request(e.app).get("/api/workers/lookup").query({ id }).set("Authorization", `Bearer ${e.user}`);
    expect((await L(nid(1))).body).toMatchObject({ found: true, approved: true, worker: { name: "Worker 1", company: "Alpha Co" } });
    expect((await L("٢٩٥٠١٠١٠١٠٠٠٠١")).body.found).toBe(true);
    const t = (await L("1002")).body;
    expect(t).toMatchObject({ found: true, approved: false, worker: { status: "TERMINATED", terminationReason: "Absent" } });
    expect((await L("29501010100999")).body).toEqual({ found: false });
    expect((await request(e.app).get("/api/workers/lookup").set("Authorization", `Bearer ${e.user}`)).status).toBe(400);
  });
  it("rejects invalid uploads with 422 + row errors and keeps old data", async () => {
    await up([good(1)]);
    const bad = await up([good(2), ["abc", "", "", "", "", "", ""]]);
    expect(bad.status).toBe(422); expect(bad.body.errors[0].row).toBe(3);
    expect((await request(e.app).post("/api/admin/workers/upload").set("Authorization", `Bearer ${e.admin}`).send("not a zip")).status).toBe(400);
    const st = await request(e.app).get("/api/admin/workers/status").set("Authorization", `Bearer ${e.admin}`);
    expect(st.body.meta.total).toBe(1); expect(st.body.lastError).not.toBeNull();
    const sum = await request(e.app).get("/api/admin/workers/summary").set("Authorization", `Bearer ${e.admin}`);
    expect(sum.body.byCompany).toEqual({ "Alpha Co": { valid: 1, terminated: 0 } });
    expect(JSON.stringify(sum.body)).not.toContain(nid(1));
    const log = e.db.prepare("SELECT action FROM audit_log WHERE action LIKE 'WORKERS%'").all().map((x: any) => x.action);
    expect(log).toEqual(["WORKERS_UPLOADED", "WORKERS_UPLOADED_REJECTED"]);
  });
});

describe("alphanumeric company codes (ID Card column)", () => {
  const codeRows: Row[] = [
    [nid(1), "IR1234", "Ahmed", "Welder", "Alpha Co", "Valid", ""],
    [nid(2), "ZTC535", "Omar", "Driver", "Beta Co", "Valid", ""],
    [nid(3), "HC123456", "Ali", "Fitter", "Alpha Co", "terminated", "No show"],
  ];
  it("matches codes ignoring case, spaces, hyphens and Arabic-Indic digits; never matches empty input", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "code-")); const store = new WorkerStore({ dataDir: dir });
    expect((await store.replaceFromBuffer(await xlsx(codeRows))).ok).toBe(true);
    for (const q of ["IR1234", "ir1234", "Ir-1234", " IR 1234 ", "IR١٢٣٤", "ZTC535", "ztc-535", "HC123456", "hc١٢٣٤٥٦"]) expect(store.lookup(q), q).toBeDefined();
    expect(store.lookup("ir1234")!.name).toBe("Ahmed");
    expect(store.lookup("HC123456")!.status).toBe("TERMINATED");
    expect(store.lookup(nid(2))!.workerId).toBe("ZTC535");   // 14-digit National ID still works
    for (const q of ["", "---", "IR123", "ZTC5355", "XX1234"]) expect(store.lookup(q), q).toBeUndefined();
  });
  it("rejects two ID Cards that differ only by case/format; keeps the original text for display", async () => {
    const r = await parseWorkersBuffer(await xlsx([...codeRows, [nid(4), "ir-1234", "Dup", "", "Co", "Valid", ""]]));
    expect(r.errors.map((e) => `${e.row}:${e.field}`)).toEqual(["5:workerId"]);
    const ok = await parseWorkersBuffer(await xlsx(codeRows)); expect(ok.workers[0]!.workerId).toBe("IR1234");
  });
});

describe("similar IDs (did you mean)", () => {
  const rows2: Row[] = [[nid(1), "IR1234", "Ahmed", "", "Alpha Co", "Valid", ""], [nid(2), "ZTC535", "Omar", "", "Beta Co", "Valid", ""], [nid(900), "HC123456", "Ali", "", "Alpha Co", "terminated", "x"]];
  it("finds workers one character away, never exact matches or unrelated IDs", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sim-")); const s = new WorkerStore({ dataDir: dir }); await s.replaceFromBuffer(await xlsx(rows2));
    const names = (q: string) => s.similar(q).map((w) => w.name);
    expect(names(nid(1).slice(0, 13) + "7").sort()).toEqual(["Ahmed", "Omar"]); // last digit wrong: both neighbours are offered
    expect(names("2950101" + "09" + "00001")).toEqual(["Ahmed"]);            // one digit wrong (middle)
    expect(names("٢٩٥٠١٠١٠٩٠٠٠٠١")).toEqual(["Ahmed"]);                  // Arabic digits are normalised first
    expect(names(nid(1))).not.toContain("Ahmed");                            // the exact match itself is never "similar" (a neighbour may be)
    expect(names("29999999999999")).toEqual([]);
    expect(names("LZTC535")).toEqual(["Omar"]);                              // one extra letter
    expect(names("ZTC53")).toEqual(["Omar"]); expect(names("ZTD535")).toEqual(["Omar"]); expect(names("ir12345")).toEqual(["Ahmed"]);
    expect(names("ZTC535")).toEqual([]); expect(names("QQ99")).toEqual([]); expect(names("")).toEqual([]);
  });
  it("API: authenticated, validated, returns only name/company/status/IDs", async () => {
    const e = await env(); await up2(e);
    const g = (id: string, tok = e.user) => request(e.app).get("/api/workers/similar").query({ id }).set("Authorization", `Bearer ${tok}`);
    expect((await request(e.app).get("/api/workers/similar?id=ZTC53")).status).toBe(401);
    expect((await g("")).status).toBe(400);
    const r = await g("LZTC535"); expect(r.status).toBe(200); expect(r.body.matches).toMatchObject([{ name: "Omar", company: "Beta Co", status: "VALID" }]);
  });
  async function up2(e: Awaited<ReturnType<typeof env>>) { await e.store.replaceFromBuffer(await xlsx(rows2)); }
});
