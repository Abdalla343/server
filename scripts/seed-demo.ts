// Demo data for UI work: node dist? no: run with tsx.  DATA_DIR=/tmp/uidemo npx tsx scripts/seed-demo.ts
import ExcelJS from "exceljs";
import fs from "node:fs";
import { loadConfig } from "../src/config.js";
import { openDb } from "../src/db.js";
import { hashPassword } from "../src/auth.js";
import { WorkerStore } from "../src/workers/store.js";
import { AttendanceWriter } from "../src/excel/attendance.js";
import { localParts } from "../src/time.js";

const cfg = loadConfig(), db = openDb(cfg.dataDir), store = new WorkerStore(cfg);
const companies = ["Nile Builders", "Delta Steel", "Sinai MEP", "Cairo Concrete", "Alex Scaffolding"];
const nid = (i: number) => "2950101" + "01" + String(i).padStart(5, "0");
const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("Workers");
ws.addRow(["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status", "reason of Terminated"]);
const N = 160, names = ["Ahmed", "Mohamed", "Mahmoud", "Omar", "Youssef", "Khaled", "Hassan", "Ali"], fam = ["Saad", "Fathy", "Nabil", "Gaber", "Zaki", "Ragab", "Mansour", "Ismail"];
for (let i = 1; i <= N; i++) ws.addRow([nid(i), String(5000 + i), `${names[i % 8]} ${fam[(i * 3) % 8]}`, ["Welder", "Driver", "Carpenter", "Fitter"][i % 4], companies[i % 5], i > 154 ? "terminated" : "Valid", i > 154 ? "Contract ended" : ""]);
await store.replaceFromBuffer(Buffer.from(await wb.xlsx.writeBuffer()));
const pw = await hashPassword("demo-password-1");
db.prepare("INSERT OR IGNORE INTO users (username,password_hash,role) VALUES ('boss',?, 'ADMIN'),('gate1',?, 'USER')").run(pw, pw);

const { date } = localParts(new Date(), cfg.timezone), day = `${date.slice(6)}-${date.slice(3, 5)}-${date.slice(0, 2)}`;
let seq = 0, rnd = 7; const r = () => (rnd = (rnd * 1103515245 + 12345) % 2147483648) / 2147483648;
const ins = db.prepare(`INSERT INTO operations (operation_id,direction,id_input,national_id,worker_id,name,job_title,company,worker_status,termination_reason,outcome,note,state,gate_id,user_id,ts_iso,date_key,day,time_text,completed_at)
  VALUES (?,?,?,?,?,?,?,?,?,?,?,?, 'COMPLETED',?,2,?,?,?,?,?)`);
const records: any = { IN: [], OUT: [] };
const hm = (h: number, m: number) => `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
function add(i: number | null, dir: "IN" | "OUT", h: number, m: number, outcome: string, note: string | null = null) {
  const w = i ? { nid: nid(i), id: String(5000 + i), name: `${names[i % 8]} ${fam[(i * 3) % 8]}`, job: ["Welder", "Driver", "Carpenter", "Fitter"][i % 4]!, co: companies[i % 5]!, st: i > 154 ? "TERMINATED" : "VALID", reason: i > 154 ? "Contract ended" : null } : null;
  const op = `seed-${String(++seq).padStart(6, "0")}`, time = hm(h, m), iso = new Date(`${day}T${time}:00+02:00`).toISOString();
  ins.run(op, dir, w?.nid ?? "29501019900777", w?.nid ?? null, w?.id ?? null, w?.name ?? null, w?.job ?? null, w?.co ?? null, w?.st ?? null, w?.reason ?? null, outcome, note, `Gate ${1 + (seq % 3)}`, iso, date, day, time, iso);
  if (outcome === "APPROVED" && w) records[dir].push({ operationId: op, nationalId: w.nid, workerId: w.id, name: w.name, jobTitle: w.job, company: w.co, status: "Valid", timeText: time, timestampIso: iso });
}
const shape = [[5, 4], [6, 22], [7, 41], [8, 28], [9, 12], [10, 3]]; // [hour, check-ins]
let w = 1;
db.transaction(() => {
  for (const [h, n] of shape as [number, number][]) for (let k = 0; k < n && w <= 150; k++, w++) add(w, "IN", h, Math.floor(r() * 60), "APPROVED");
  const out = [[12, 4], [14, 9], [15, 17], [16, 21], [17, 8]]; let o = 1;
  for (const [h, n] of out as [number, number][]) for (let k = 0; k < n && o < w - 40; k++, o++) add(o, "OUT", h, Math.floor(r() * 60), "APPROVED");
  for (let k = 0; k < 7; k++) add(3 + k, "IN", 7 + (k % 3), 50 + k, "ALREADY_REGISTERED", `IN 07:${10 + k}`);
  for (let k = 0; k < 3; k++) add(155 + k, "IN", 6 + k, 30 + k, "NOT_APPROVED");
  for (let k = 0; k < 5; k++) add(null, "IN", 6 + (k % 4), 10 + 7 * k, "NOT_FOUND");
})();
const writer = new AttendanceWriter(cfg.dataDir);
for (const dir of ["IN", "OUT"] as const) await writer.appendBatch(dir, date, records[dir]);
console.log(`seeded ${seq} scans for ${date}`);
db.close();
