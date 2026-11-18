import ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import { cellText } from "../workers/parse.js";

export type Direction = "IN" | "OUT";
export interface AttendanceRecord {
  operationId: string; nationalId: string; workerId: string; name: string; jobTitle: string;
  company: string; status: string; timeText: string; timestampIso: string; scannedBy?: string; scannedUsername?: string;
  scanMethod?: string; dateKey?: string; direction?: Direction; rejectionReason?: string;
}
export class AttendanceFileError extends Error { readonly code = "FILE_INVALID"; }
export class FileLockedError extends Error { readonly code = "FILE_LOCKED"; }

const VISIBLE = ["National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status"];
const timeHeader = (d: Direction) => (d === "IN" ? "timeCheckIn" : "timeCheckOut");
const OPID_COL = 8, ISO_COL = 9; // hidden columns
type FileKind = Direction | "REJECTED";
const headers = (d: FileKind) => d === "REJECTED"
  ? ["Employee ID", "National ID Number", "Name", "Date", "Time", "Operation Type", "Rejection Reason", "Operation ID", "Timestamp (ISO)", "Company", "Details"]
  : [...VISIBLE, timeHeader(d), "Operation ID", "Timestamp (ISO)"];
const OPERATOR_HEADERS = ["Scanned By", "Scanned By Username", "Operation Type", "Rejection Reason", "Operation Scan"];
export const rejectedPath = (dataDir: string) => path.join(dataDir, "attendance", "rejected", "Rejected.xlsx");

export interface WriterOptions {
  retries?: number; retryDelayMs?: number;
  operatorFor?: (operationId: string) => { scannedBy: string; scannedUsername: string; scanMethod?: string } | undefined;
  /** Test seams for fault injection. */
  rename?: (from: string, to: string) => void;
  hooks?: { afterTempWrite?: () => void; afterRename?: () => void };
}
export interface BatchResult { appended: string[]; duplicates: string[] }

/** Folder `DD-MM-YYYY`, files `DD-MM-YYYYCheckIn.xlsx` / `DD-MM-YYYYCheckOut.xlsx`. */
export function attendancePath(dataDir: string, date: string, d: Direction) {
  return path.join(dataDir, "attendance", date, `${date}${d === "IN" ? "CheckIn" : "CheckOut"}.xlsx`);
}

export class AttendanceWriter {
  constructor(private dataDir: string, private opt: WriterOptions = {}) {}
  private rename = (a: string, b: string) => (this.opt.rename ?? fs.renameSync)(a, b);

  /** Removes half-written temp files left by a crash. Call at startup. */
  cleanupTemps(): number {
    const root = path.join(this.dataDir, "attendance"); let n = 0;
    if (!fs.existsSync(root)) return 0;
    for (const day of fs.readdirSync(root))
      for (const f of fs.readdirSync(path.join(root, day))) if (f.endsWith(".xlsx.tmp")) { fs.rmSync(path.join(root, day, f)); n++; }
    return n;
  }

  /** Creates the day's folder and both (empty, header-only) files if missing. */
  async ensureDay(date: string) {
    for (const d of ["IN", "OUT"] as Direction[]) await this.appendBatch(d, date, []);
    await this.appendRejected([]);
  }

  /**
   * Appends records idempotently: any Operation ID already in the file is skipped.
   * The file is replaced only after the new version has been written, re-read and verified.
   * Throws (leaving the original untouched) if the file is invalid or locked.
   */
  appendBatch(direction: Direction, date: string, records: AttendanceRecord[]): Promise<BatchResult> {
    return this.appendFile(attendancePath(this.dataDir, date, direction), direction, records);
  }

  appendRejected(records: AttendanceRecord[]): Promise<BatchResult> {
    return this.appendFile(rejectedPath(this.dataDir), "REJECTED", records);
  }

  async removeOperations(direction: FileKind, date: string, operationIds: string[]): Promise<number> {
    const file = direction === "REJECTED" ? rejectedPath(this.dataDir) : attendancePath(this.dataDir, date, direction);
    if (!fs.existsSync(file)) return 0;
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.readFile(file); }
    catch (e: any) { throw new AttendanceFileError(e.message); }
    const ws = this.checkHeaders(wb, direction, file), remove = new Set(operationIds);
    const before = ws.actualRowCount - 1;
    let removed = 0;
    for (let row = ws.rowCount; row >= 2; row--) {
      if (remove.has(cellText(ws.getRow(row).getCell(OPID_COL).value))) {
        ws.spliceRows(row, 1);
        removed++;
      }
    }
    if (removed) await this.commit(wb, file, direction, this.readIds(ws), before, -removed, true);
    return removed;
  }

  private async appendFile(file: string, direction: FileKind, records: AttendanceRecord[]): Promise<BatchResult> {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    const existed = fs.existsSync(file);
    const wb = new ExcelJS.Workbook();
    let ws: ExcelJS.Worksheet;
    if (existed) {
      try { await wb.xlsx.readFile(file); } catch (e: any) { throw new AttendanceFileError(`${path.basename(file)} cannot be read: ${e.message}`); }
      ws = this.checkHeaders(wb, direction, file);
    } else ws = this.newSheet(wb, direction);

    const priorColumnCount = ws.columnCount;
    const operatorColumns = OPERATOR_HEADERS.map((name) => {
      let col = 0;
      ws.getRow(1).eachCell((cell, index) => { if (cellText(cell.value) === name) col = index; });
      if (!col) { col = ws.columnCount + 1; ws.getRow(1).getCell(col).value = name; }
      ws.getColumn(col).width = 24;
      ws.getColumn(col).numFmt = "@";
      return col;
    });
    const upgraded = ws.columnCount !== priorColumnCount;
    if (upgraded) {
      for (let index = 2; index <= ws.rowCount; index++) {
        const row = ws.getRow(index);
        const operator = this.opt.operatorFor?.(cellText(row.getCell(OPID_COL).value));
        if (operator && !cellText(row.getCell(operatorColumns[0]!).value)) {
          row.getCell(operatorColumns[0]!).value = operator.scannedBy;
          row.getCell(operatorColumns[1]!).value = operator.scannedUsername;
        }
        if (!cellText(row.getCell(operatorColumns[4]!).value)) row.getCell(operatorColumns[4]!).value = operator?.scanMethod ?? "Unknown";
        if (direction !== "REJECTED" && !cellText(row.getCell(operatorColumns[2]!).value))
          row.getCell(operatorColumns[2]!).value = direction === "IN" ? "Check-In" : "Check-Out";
      }
    }
    const ids = this.readIds(ws);
    const rowsBefore = ws.actualRowCount - 1;
    const result: BatchResult = { appended: [], duplicates: [] };
    for (const r of records) {
      if (ids.has(r.operationId)) { result.duplicates.push(r.operationId); continue; }
      ids.add(r.operationId); result.appended.push(r.operationId);
      const values = direction === "REJECTED"
        ? [r.workerId || r.nationalId, r.nationalId, r.name, r.dateKey, r.timeText,
          r.direction === "IN" ? "Check-In" : "Check-Out", r.rejectionReason, r.operationId, r.timestampIso, r.company, r.status]
        : [r.nationalId, r.workerId, r.name, r.jobTitle, r.company, r.status, r.timeText, r.operationId, r.timestampIso];
      const row = ws.addRow(values);
      row.getCell(operatorColumns[0]!).value = r.scannedBy ?? "";
      row.getCell(operatorColumns[1]!).value = r.scannedUsername ?? "";
      row.getCell(operatorColumns[4]!).value = r.scanMethod ?? "Unknown";
      if (direction !== "REJECTED") {
        row.getCell(operatorColumns[2]!).value = direction === "IN" ? "Check-In" : "Check-Out";
        row.getCell(operatorColumns[3]!).value = "";
      }
      row.eachCell((c) => { c.numFmt = "@"; }); // plain text everywhere: no 2.95E+13, no time conversion
    }
    if (existed && !upgraded && !result.appended.length) return result; // nothing new to write

    await this.commit(wb, file, direction, ids, rowsBefore, result.appended.length, existed);
    return result;
  }

  private newSheet(wb: ExcelJS.Workbook, d: FileKind) {
    const ws = wb.addWorksheet(d === "REJECTED" ? "Rejected" : d === "IN" ? "CheckIn" : "CheckOut");
    const h = ws.addRow(headers(d)); h.font = { bold: true };
    ws.views = [{ state: "frozen", ySplit: 1 }];
    [20, 14, 28, 18, 24, 12, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));
    ws.getColumn(OPID_COL).hidden = true; ws.getColumn(ISO_COL).hidden = true;
    for (let c = 1; c <= 9; c++) ws.getColumn(c).numFmt = "@";
    return ws;
  }

  private checkHeaders(wb: ExcelJS.Workbook, d: FileKind, file: string): ExcelJS.Worksheet {
    const ws = wb.worksheets[0];
    if (!ws) throw new AttendanceFileError(`${path.basename(file)} has no worksheet`);
    const want = headers(d), row = ws.getRow(1);
    for (let i = 0; i < want.length; i++)
      if (cellText(row.getCell(i + 1).value) !== want[i])
        throw new AttendanceFileError(`${path.basename(file)}: column ${i + 1} should be "${want[i]}" (found "${cellText(row.getCell(i + 1).value)}"). Refusing to write. Restore the file or its header.`);
    return ws;
  }

  private readIds(ws: ExcelJS.Worksheet): Set<string> {
    const ids = new Set<string>();
    for (let r = 2; r <= ws.rowCount; r++) { const v = cellText(ws.getRow(r).getCell(OPID_COL).value); if (v) ids.add(v); }
    return ids;
  }

  private async commit(wb: ExcelJS.Workbook, file: string, d: FileKind, expectIds: Set<string>, rowsBefore: number, added: number, existed: boolean) {
    const tmp = file + ".tmp";
    try {
      const buf = Buffer.from(await wb.xlsx.writeBuffer());
      const fd = fs.openSync(tmp, "w");
      try { fs.writeSync(fd, buf); fs.fsyncSync(fd); } finally { fs.closeSync(fd); }
      this.opt.hooks?.afterTempWrite?.();

      // Verify the temp file BEFORE it can replace anything.
      const chk = new ExcelJS.Workbook();
      await chk.xlsx.readFile(tmp);
      const cws = this.checkHeaders(chk, d, tmp);
      const got = this.readIds(cws);
      if (cws.actualRowCount - 1 !== rowsBefore + added || ![...expectIds].every((id) => got.has(id)))
        throw new AttendanceFileError(`Verification of ${path.basename(file)} failed (row count or Operation IDs mismatch); original kept.`);

      if (existed) fs.copyFileSync(file, file + ".bak");
      await this.renameWithRetry(tmp, file);
      this.opt.hooks?.afterRename?.();
    } finally {
      fs.rmSync(tmp, { force: true }); // no-op after a successful rename
    }
  }

  private async renameWithRetry(from: string, to: string) {
    const tries = this.opt.retries ?? 5;
    for (let i = 1; ; i++) {
      try { return this.rename(from, to); }
      catch (e: any) {
        if (!["EBUSY", "EPERM", "EACCES"].includes(e.code)) throw e;
        if (i >= tries) throw new FileLockedError(`${path.basename(to)} is locked (probably open in Excel). Nothing was lost; close the file and retry.`);
        await new Promise((r) => setTimeout(r, (this.opt.retryDelayMs ?? 200) * i));
      }
    }
  }
}
