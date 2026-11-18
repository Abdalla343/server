import ExcelJS from "exceljs";
import { isFormat14, nidSemanticProblem, normalizeCode, normalizeNid, toWesternDigits } from "../nationalId.js";

export type WorkerStatus = "VALID" | "TERMINATED";
export interface Worker {
  nationalId: string; workerId: string; name: string; jobTitle: string;
  company: string; status: WorkerStatus; terminationReason: string; row: number;
}
export interface Issue { row: number; field: string; message: string }
export interface ParseResult { workers: Worker[]; errors: Issue[]; warnings: Issue[]; sheet: string | null }

export const MAX_ROWS = 100_000;

type Field = "nationalId" | "workerId" | "name" | "jobTitle" | "company" | "status" | "reason";
/** Header aliases, compared after normalization (lowercase, letters/digits only). Real headers: see first of each list. */
const ALIASES: Record<Field, string[]> = {
  nationalId: ["nationalidnumber", "nationalid", "nid", "الرقمالقومي", "رقمقومي"],
  workerId: ["idcard", "idcardnumber", "workerid", "cardnumber", "رقمالكارنيه", "رقمالبطاقة"],
  name: ["name", "workername", "fullname", "الاسم"],
  jobTitle: ["jobtitle", "job", "position", "الوظيفة", "المهنة"],
  company: ["namecompany", "company", "companyname", "الشركة", "اسمالشركة"],
  status: ["status", "الحالة"],
  reason: ["reasonofterminated", "reasonofterminate", "terminationreason", "reason", "سببالانهاء", "سببالإنهاء"],
};
const REQUIRED: Field[] = ["nationalId", "name", "company", "status"];
const normHeader = (s: string) => s.toLowerCase().replace(/[^a-z0-9\u0600-\u06ff]/g, "");

export function cellText(v: ExcelJS.CellValue | undefined): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isInteger(v) ? v.toFixed(0) : String(v);
  if (typeof v === "string") return v.trim();
  if (typeof v === "boolean") return String(v);
  if (v instanceof Date) return v.toISOString();
  const o = v as any;
  if (Array.isArray(o.richText)) return o.richText.map((t: any) => t.text).join("").trim();
  if (o.result !== undefined) return cellText(o.result);
  if (o.text !== undefined) return cellText(o.text);
  if (o.error) return "";
  return String(v).trim();
}

const STATUS_MAP: Record<string, WorkerStatus> = { valid: "VALID", active: "VALID", terminated: "TERMINATED", terminate: "TERMINATED" };

export async function parseWorkersBuffer(buf: Buffer): Promise<ParseResult> {
  const out: ParseResult = { workers: [], errors: [], warnings: [], sheet: null };
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(buf as any); }
  catch { out.errors.push({ row: 0, field: "file", message: "Not a readable .xlsx file" }); return out; }

  // Find the first sheet whose first rows contain a National ID header.
  for (const ws of wb.worksheets) {
    for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
      const map = headerMap(ws.getRow(r));
      if (map.nationalId === undefined) continue;
      const missing = REQUIRED.filter((f) => map[f] === undefined);
      if (missing.length) {
        out.errors.push({ row: r, field: "header", message: `Missing required column(s): ${missing.join(", ")}` });
        return out;
      }
      for (const f of ["workerId", "jobTitle", "reason"] as Field[])
        if (map[f] === undefined) out.warnings.push({ row: r, field: f, message: "Optional column not found" });
      out.sheet = ws.name;
      readRows(ws, r, map, out);
      return out;
    }
  }
  out.errors.push({ row: 0, field: "header", message: "No sheet with a 'National ID Number' column found" });
  return out;
}

function headerMap(row: ExcelJS.Row): Partial<Record<Field, number>> {
  const map: Partial<Record<Field, number>> = {};
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    const h = normHeader(cellText(cell.value));
    for (const f of Object.keys(ALIASES) as Field[]) if (map[f] === undefined && ALIASES[f].includes(h)) map[f] = col;
  });
  return map;
}

function readRows(ws: ExcelJS.Worksheet, headerRow: number, map: Partial<Record<Field, number>>, out: ParseResult) {
  const seenNid = new Map<string, number>(), seenWid = new Map<string, number>();
  const get = (row: ExcelJS.Row, f: Field) => (map[f] === undefined ? "" : cellText(row.getCell(map[f]!).value));
  const last = ws.rowCount;
  if (last - headerRow > MAX_ROWS) { out.errors.push({ row: 0, field: "file", message: `More than ${MAX_ROWS} rows` }); return; }

  for (let r = headerRow + 1; r <= last; r++) {
    const row = ws.getRow(r);
    const nidRaw = get(row, "nationalId"), name = get(row, "name"), company = get(row, "company"), statusRaw = get(row, "status");
    if (!nidRaw && !name && !company && !statusRaw) continue; // blank row
    const err = (field: string, message: string) => out.errors.push({ row: r, field, message });

    const nationalId = normalizeNid(nidRaw);
    if (!isFormat14(nationalId)) err("nationalId", `"${nidRaw.slice(0, 20)}" is not 14 digits`);
    else {
      const sem = nidSemanticProblem(nationalId);
      if (sem) out.warnings.push({ row: r, field: "nationalId", message: `${sem} (loaded anyway)` });
      const dup = seenNid.get(nationalId);
      if (dup) err("nationalId", `Duplicate National ID (also on row ${dup})`); else seenNid.set(nationalId, r);
    }
    if (!name) err("name", "Name is empty");
    if (!company) err("company", "Company is empty");
    const status = STATUS_MAP[statusRaw.toLowerCase()];
    if (!status) err("status", `Status "${statusRaw.slice(0, 20)}" must be Valid or Terminated`);

    const workerId = toWesternDigits(get(row, "workerId")).trim();
    if (workerId) {
      const key = normalizeCode(workerId); // case/space/hyphen-insensitive, same rule as scanning
      if (!key) err("workerId", `ID Card "${workerId.slice(0, 20)}" has no letters or digits`);
      else {
        const dup = seenWid.get(key);
        if (dup) err("workerId", `Duplicate ID Card "${workerId}" (also on row ${dup})`); else seenWid.set(key, r);
      }
    }
    const terminationReason = get(row, "reason");
    if (status === "TERMINATED" && !terminationReason) out.warnings.push({ row: r, field: "reason", message: "Terminated without a reason" });

    if (nationalId && status) out.workers.push({ nationalId, workerId, name, jobTitle: get(row, "jobTitle"), company, status, terminationReason, row: r });
  }
  if (!out.errors.length && !out.workers.length) out.errors.push({ row: 0, field: "file", message: "File contains no workers" });
}
