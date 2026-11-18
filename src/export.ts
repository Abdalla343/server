import ExcelJS from "exceljs";
import type { Db } from "./db.js";

const COLS = ["Date", "National ID Number", "ID Card", "Name", "jobTitle", "nameCompany", "Status"];
/** Consolidated export for a date range, built from the completed operations. Plain-text cells, like the daily files. */
export async function buildExport(db: Db, from: string, to: string): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const sheet = (name: string, extra: string[]) => {
    const ws = wb.addWorksheet(name); ws.addRow([...COLS, ...extra]).font = { bold: true };
    ws.views = [{ state: "frozen", ySplit: 1 }]; return ws;
  };
  const add = (ws: ExcelJS.Worksheet, vals: unknown[]) => ws.addRow(vals).eachCell((c) => { c.numFmt = "@"; });
  const q = (dir: string) => db.prepare("SELECT * FROM active_operations WHERE direction=? AND outcome='APPROVED' AND state='COMPLETED' AND day BETWEEN ? AND ? ORDER BY id").all(dir, from, to) as any[];
  const wi = sheet("CheckIn", ["timeCheckIn", "Operation Type", "Scanned By", "Scanned By Username", "Rejection Reason", "Operation Scan"]), wo = sheet("CheckOut", ["timeCheckOut", "Operation Type", "Scanned By", "Scanned By Username", "Rejection Reason", "Operation Scan"]);
  for (const r of q("IN")) add(wi, [r.date_key, r.national_id, r.worker_id, r.name, r.job_title, r.company, "Valid", r.time_text, r.direction === "IN" ? "Check-In" : "Check-Out", r.scanned_by, r.scanned_username, "", r.scan_method === "SCANNER" ? "Scanner" : r.scan_method === "MANUAL" ? "Manual ID Entry" : "Unknown"]);
  for (const r of q("OUT")) add(wo, [r.date_key, r.national_id, r.worker_id, r.name, r.job_title, r.company, "Valid", r.time_text, r.direction === "IN" ? "Check-In" : "Check-Out", r.scanned_by, r.scanned_username, "", r.scan_method === "SCANNER" ? "Scanner" : r.scan_method === "MANUAL" ? "Manual ID Entry" : "Unknown"]);
  const wr = sheet("Rejected", ["Time", "Direction", "Result", "Reason", "Scanned By", "Scanned By Username", "Rejection Reason", "Operation Type", "Operation Scan"]);
  for (const r of db.prepare("SELECT * FROM active_operations WHERE outcome!='APPROVED' AND day BETWEEN ? AND ? ORDER BY id").all(from, to) as any[])
    add(wr, [r.date_key, r.national_id ?? r.id_input, r.worker_id || r.id_input, r.name, r.job_title, r.company, r.worker_status ?? "Unknown", r.time_text, r.direction, r.outcome, r.note === "NO_CHECK_IN" ? "No completed active check-in" : r.termination_reason ?? r.note ?? "", r.scanned_by, r.scanned_username, r.note === "NO_CHECK_IN" ? "No Check-In" : r.outcome === "NOT_APPROVED" ? "Terminated" : r.outcome === "NOT_FOUND" ? "Not Found" : "Already Registered", r.direction === "IN" ? "Check-In" : "Check-Out", r.scan_method === "SCANNER" ? "Scanner" : r.scan_method === "MANUAL" ? "Manual ID Entry" : "Unknown"]);
  return Buffer.from(await wb.xlsx.writeBuffer());
}
