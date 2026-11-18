import { describe, it, expect } from "vitest";
import ExcelJS from "exceljs";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { AttendanceWriter, attendancePath, AttendanceFileError, FileLockedError, type AttendanceRecord, type Direction } from "../src/excel/attendance.js";
import { localParts } from "../src/time.js";

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), "att-"));
const rec = (i: number, over: Partial<AttendanceRecord> = {}): AttendanceRecord => ({
  operationId: `op-${i}`, nationalId: "2950101010" + String(i).padStart(4, "0"), workerId: String(1000 + i), name: `Worker ${i}`,
  jobTitle: "Welder", company: "Alpha Co", status: "Valid", timeText: "12:00", timestampIso: "2026-11-14T10:00:00.000Z", ...over });
async function read(file: string) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file); const ws = wb.worksheets[0]!;
  return { ws, rows: ws.getSheetValues().slice(2) as any[][] };
}
const D = "14-11-2026";

describe("local time", () => {
  it("formats Cairo date/time as DD-MM-YYYY and HH:mm, DST-aware, correct at midnight", () => {
    expect(localParts(new Date("2026-11-14T10:05:00Z"), "Africa/Cairo")).toEqual({ date: "14-11-2026", time: "12:05" });   // UTC+2 (winter)
    expect(localParts(new Date("2026-07-01T21:30:00Z"), "Africa/Cairo")).toEqual({ date: "02-07-2026", time: "00:30" });   // UTC+3 (summer), next day
    expect(localParts(new Date("2026-11-14T21:59:00Z"), "Africa/Cairo")).toEqual({ date: "14-11-2026", time: "23:59" });
    expect(localParts(new Date("2026-11-14T22:00:00Z"), "Africa/Cairo")).toEqual({ date: "15-11-2026", time: "00:00" });
  });
});

describe("attendance writer", () => {
  it("creates correctly named files with your headers, hidden ID columns and plain-text cells", async () => {
    const dir = tmp(); const w = new AttendanceWriter(dir);
    await w.ensureDay(D);
    expect(fs.existsSync(path.join(dir, "attendance", D, `${D}CheckIn.xlsx`))).toBe(true);
    expect(fs.existsSync(path.join(dir, "attendance", D, `${D}CheckOut.xlsx`))).toBe(true);
    await w.appendBatch("IN", D, [rec(1, { timeText: "18:30" })]);
    await w.appendBatch("OUT", D, [rec(1, { operationId: "op-out-1", timeText: "07:05" })]);
    const inn = await read(attendancePath(dir, D, "IN")), out = await read(attendancePath(dir, D, "OUT"));
    expect((inn.ws.getRow(1).values as any[]).slice(1)).toEqual(["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status", "timeCheckIn", "Operation ID", "Timestamp (ISO)", "Scanned By", "Scanned By Username", "Operation Type", "Rejection Reason", "Operation Scan"]);
    expect(out.ws.getRow(1).getCell(7).value).toBe("timeCheckOut");
    expect(inn.ws.getColumn(8).hidden).toBe(true);
    const c = inn.ws.getRow(2);
    expect(c.getCell(1).value).toBe("29501010100001");          // text, not a number
    expect(c.getCell(7).value).toBe("18:30");                   // text, not a time serial
    expect(c.getCell(7).type).toBe(ExcelJS.ValueType.String);
    expect(out.ws.getRow(2).getCell(7).value).toBe("07:05");
  });

  it("appends, never overwrites, and skips duplicate Operation IDs (also inside one batch)", async () => {
    const dir = tmp(); const w = new AttendanceWriter(dir);
    expect(await w.appendBatch("IN", D, [rec(1), rec(2)])).toEqual({ appended: ["op-1", "op-2"], duplicates: [] });
    const r = await w.appendBatch("IN", D, [rec(2), rec(3), rec(3)]);
    expect(r).toEqual({ appended: ["op-3"], duplicates: ["op-2", "op-3"] });
    const { rows } = await read(attendancePath(dir, D, "IN"));
    expect(rows.map((x) => x[8])).toEqual(["op-1", "op-2", "op-3"]);
    // Same worker may legitimately appear again with a new Operation ID (re-entry later in the day).
    await w.appendBatch("IN", D, [rec(1, { operationId: "op-9", timeText: "15:00" })]);
    expect((await read(attendancePath(dir, D, "IN"))).rows.length).toBe(4);
  });

  it("crash BEFORE rename: original untouched; stray .tmp is cleaned; retry appends once", async () => {
    const dir = tmp();
    await new AttendanceWriter(dir).appendBatch("IN", D, [rec(1)]);
    const file = attendancePath(dir, D, "IN"); const before = fs.readFileSync(file);
    const crashing = new AttendanceWriter(dir, { hooks: { afterTempWrite: () => { throw new Error("power cut"); } } });
    await expect(crashing.appendBatch("IN", D, [rec(2)])).rejects.toThrow("power cut");
    expect(fs.readFileSync(file).equals(before)).toBe(true);
    fs.writeFileSync(file + ".tmp", "half-written garbage"); // what a real kill -9 would leave behind
    const w = new AttendanceWriter(dir); expect(w.cleanupTemps()).toBe(1);
    expect((await w.appendBatch("IN", D, [rec(2)])).appended).toEqual(["op-2"]);
    expect((await read(file)).rows.map((x) => x[8])).toEqual(["op-1", "op-2"]);
  });

  it("crash AFTER rename (DB not yet updated): retry is detected as duplicate, no second row", async () => {
    const dir = tmp();
    const crashing = new AttendanceWriter(dir, { hooks: { afterRename: () => { throw new Error("killed after save"); } } });
    await expect(crashing.appendBatch("OUT", D, [rec(1)])).rejects.toThrow("killed after save");
    const retry = await new AttendanceWriter(dir).appendBatch("OUT", D, [rec(1)]);
    expect(retry).toEqual({ appended: [], duplicates: ["op-1"] });
    expect((await read(attendancePath(dir, D, "OUT"))).rows.length).toBe(1);
  });

  it("locked file: retries then throws FileLockedError, original intact; transient lock succeeds", async () => {
    const dir = tmp(); await new AttendanceWriter(dir).appendBatch("IN", D, [rec(1)]);
    const file = attendancePath(dir, D, "IN"); const before = fs.readFileSync(file);
    const busy = Object.assign(new Error("busy"), { code: "EBUSY" });
    const stuck = new AttendanceWriter(dir, { retryDelayMs: 1, rename: () => { throw busy; } });
    await expect(stuck.appendBatch("IN", D, [rec(2)])).rejects.toBeInstanceOf(FileLockedError);
    expect(fs.readFileSync(file).equals(before)).toBe(true); expect(fs.existsSync(file + ".tmp")).toBe(false);
    let n = 0;
    const flaky = new AttendanceWriter(dir, { retryDelayMs: 1, rename: (a, b) => { if (n++ < 2) throw busy; fs.renameSync(a, b); } });
    expect((await flaky.appendBatch("IN", D, [rec(2)])).appended).toEqual(["op-2"]); expect(n).toBe(3);
  });

  it("refuses to touch corrupted or hand-damaged files", async () => {
    const dir = tmp(); const w = new AttendanceWriter(dir); await w.appendBatch("IN", D, [rec(1)]);
    const file = attendancePath(dir, D, "IN");
    // 1) hidden Operation ID header deleted by hand
    const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file); wb.worksheets[0]!.getRow(1).getCell(8).value = "oops"; await wb.xlsx.writeFile(file);
    let snap = fs.readFileSync(file);
    await expect(w.appendBatch("IN", D, [rec(2)])).rejects.toBeInstanceOf(AttendanceFileError);
    expect(fs.readFileSync(file).equals(snap)).toBe(true);
    // 2) not an xlsx at all
    fs.writeFileSync(file, "garbage"); snap = fs.readFileSync(file);
    await expect(w.appendBatch("IN", D, [rec(2)])).rejects.toBeInstanceOf(AttendanceFileError);
    expect(fs.readFileSync(file).equals(snap)).toBe(true);
  });

  it("keeps a .bak of the previous version and never leaves a .tmp behind", async () => {
    const dir = tmp(); const w = new AttendanceWriter(dir);
    await w.appendBatch("IN", D, [rec(1)]); await w.appendBatch("IN", D, [rec(2)]);
    const file = attendancePath(dir, D, "IN");
    expect((await read(file + ".bak")).rows.length).toBe(1);
    expect(fs.readdirSync(path.dirname(file)).filter((f) => f.endsWith(".tmp"))).toEqual([]);
  });

  it("benchmark: append cost as the file grows to 3,200 rows", async () => {
    const dir = tmp(); const w = new AttendanceWriter(dir); const out: string[] = [];
    let n = 0; const time = async (k: number) => { const t = performance.now(); await w.appendBatch("IN", D, Array.from({ length: k }, () => rec(++n))); return performance.now() - t; };
    await time(3180); // pre-fill in one batch
    out.push(`single append @3180 rows: ${(await time(1)).toFixed(0)} ms`);
    out.push(`batch of 20 @3181 rows: ${(await time(20)).toFixed(0)} ms`);
    console.log(out.join("\n"));
    expect((await read(attendancePath(dir, D, "IN"))).rows.length).toBe(3201);
  }, 60_000);
});

describe("legacy attendance files", () => {
  it("adds operator columns after custom columns and preserves old cells and worksheets", async () => {
    const dir = tmp(), file = attendancePath(dir, D, "IN");
    fs.mkdirSync(path.dirname(file), { recursive: true });
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet("CheckIn");
    ws.addRow(["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status", "timeCheckIn", "Operation ID", "Timestamp (ISO)", "Custom note"]);
    ws.addRow(["29501010100001", "A001", "Existing", "Welder", "Old Co", "Valid", "07:10", "legacy-op", "2026-11-14T05:10:00Z", "Keep me"]);
    wb.addWorksheet("Notes").getCell("A1").value = "Existing notes";
    await wb.xlsx.writeFile(file);
    await new AttendanceWriter(dir).appendBatch("IN", D, [rec(2, { scannedBy: "Gate Two", scannedUsername: "gate2" })]);
    const saved = new ExcelJS.Workbook(); await saved.xlsx.readFile(file);
    const sheet = saved.worksheets[0]!;
    expect(sheet.getRow(2).getCell(10).value).toBe("Keep me");
    expect(saved.getWorksheet("Notes")!.getCell("A1").value).toBe("Existing notes");
    expect(sheet.getRow(1).getCell(11).value).toBe("Scanned By");
    expect(sheet.getRow(3).getCell(11).value).toBe("Gate Two");
    expect(sheet.getRow(3).getCell(12).value).toBe("gate2");
    expect(sheet.getRow(3).getCell(13).value).toBe("Check-In");
    expect((await read(file + ".bak")).ws.getRow(2).getCell(10).value).toBe("Keep me");
  });
});
