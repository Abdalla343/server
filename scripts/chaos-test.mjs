// Real-process crash test: fires concurrent scans, SIGKILLs the server mid-flight (twice), restarts it,
// lets clients resend the same operationIds, then verifies every scan is in Excel EXACTLY once.
// Usage (after `npm run build`):  node scripts/chaos-test.mjs [scans=400] [killAfterMs=300]
import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs"; import os from "node:os"; import path from "node:path";
import ExcelJS from "exceljs";

const N = Number(process.argv[2] ?? 400), KILL_MS = Number(process.argv[3] ?? 300);
const dir = fs.mkdtempSync(path.join(os.tmpdir(), "chaos-")), PORT = 3100 + Math.floor(Math.random() * 500);
const env = { ...process.env, DATA_DIR: dir, PORT: String(PORT), LOG_LEVEL: "warn", ADMIN_PASSWORD: "chaos-pass-12345" };
const base = `http://127.0.0.1:${PORT}`, root = path.resolve(import.meta.dirname, "..");
const nid = (i) => "2950101" + "01" + String(i).padStart(5, "0");

spawnSync("node", [path.join(root, "dist/scripts/create-admin.js"), "chaos"], { env, cwd: root, stdio: "pipe" });
const startServer = async () => {
  const p = spawn("node", [path.join(root, "dist/src/index.js")], { env, cwd: root, stdio: "ignore" });
  for (let i = 0; i < 100; i++) { try { if ((await fetch(base + "/api/health")).ok) return p; } catch {} await new Promise((r) => setTimeout(r, 100)); }
  throw new Error("server did not start");
};
const kill9 = (p) => new Promise((r) => { p.once("exit", r); p.kill("SIGKILL"); });

const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("Workers");
ws.addRow(["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status", "reason of Terminated"]);
for (let i = 1; i <= N; i++) ws.addRow([nid(i), String(1000 + i), `Worker ${i}`, "Welder", "Alpha Co", "Valid", ""]);
const xlsx = Buffer.from(await wb.xlsx.writeBuffer());

let server = await startServer();
const token = (await (await fetch(base + "/api/auth/login", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ username: "chaos", password: env.ADMIN_PASSWORD }) })).json()).token;
const H = { authorization: `Bearer ${token}`, "content-type": "application/json" };
console.log("workers upload:", (await fetch(base + "/api/admin/workers/upload", { method: "POST", headers: { authorization: H.authorization }, body: xlsx })).status);

const acked = new Set();
const fire = (killMs) => {
  const reqs = Array.from({ length: N }, (_, i) => fetch(base + "/api/scans", { method: "POST", headers: H, body: JSON.stringify({ operationId: `chaos-op-${String(i + 1).padStart(5, "0")}`, idNumber: nid(i + 1), direction: "IN" }) })
    .then(async (r) => { const b = await r.json(); if (r.status === 200 && b.state === "COMPLETED") acked.add(b.operationId); return r.status; }).catch(() => "dropped"));
  return Promise.all([Promise.all(reqs), killMs ? new Promise((r) => setTimeout(r, killMs)).then(() => kill9(server)) : null]).then(([s]) => s);
};
const tally = (a) => a.reduce((m, x) => ((m[x] = (m[x] ?? 0) + 1), m), {});

console.log("round 1 (kill -9 after", KILL_MS, "ms):", tally(await fire(KILL_MS)), "acked:", acked.size);
server = await startServer();
console.log("round 2 (resend all, kill -9 again):", tally(await fire(Math.round(KILL_MS / 2))), "acked:", acked.size);
server = await startServer();
console.log("round 3 (resend all, no kill):", tally(await fire(0)), "acked:", acked.size);

const st = await (await fetch(base + "/api/queue/status", { headers: H })).json();
console.log("queue:", st.counts);
const date = (await (await fetch(base + "/api/attendance/days", { headers: H })).json()).days[0];
const rb = new ExcelJS.Workbook(); await rb.xlsx.readFile(path.join(dir, "attendance", date, `${date}CheckIn.xlsx`));
const ids = rb.worksheets[0].getSheetValues().slice(2).map((r) => r[8]);
const unique = new Set(ids), missingAcked = [...acked].filter((x) => !unique.has(x));
const ok = ids.length === N && unique.size === N && !missingAcked.length && !st.counts.FAILED && !st.counts.PENDING;
console.log(`rows in Excel: ${ids.length}, unique: ${unique.size}, expected: ${N}, acknowledged-but-missing: ${missingAcked.length}`);
await kill9(server);
console.log(ok ? "RESULT: PASS - every scan exactly once, nothing acknowledged was lost" : "RESULT: FAIL");
process.exit(ok ? 0 : 1);
