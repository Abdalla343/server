import type { Db } from "./db.js";

/** Indexed projection of completed operations. Repeated allowed INs keep the original entry time. */
export function onSiteCount(db: Db) {
  return (db.prepare("SELECT COUNT(*) n FROM current_presence").get() as { n: number }).n;
}

export function onSitePage(db: Db, search: string, offset: number, limit: number, now: Date) {
  const term = search.trim().slice(0, 100);
  const match = `(?='' OR instr(lower(COALESCE(name,'')),lower(?))>0
    OR instr(lower(COALESCE(worker_id,'')),lower(?))>0
    OR instr(COALESCE(national_id,''),?)>0 OR instr(lower(COALESCE(company,'')),lower(?))>0)`;
  const args = [term, term, term, term, term];
  const filtered = (db.prepare(`SELECT COUNT(*) n FROM current_presence WHERE ${match}`).get(...args) as { n: number }).n;
  const rows = db.prepare(`SELECT national_id, worker_id, name, company, ts_iso, time_text, date_key,
      scanned_by, scanned_username FROM current_presence WHERE ${match} ORDER BY id LIMIT ? OFFSET ?`)
    .all(...args, limit, offset) as any[];
  return {
    total: onSiteCount(db), filtered, offset, limit, serverNow: now.toISOString(),
    employees: rows.map((r) => ({
      nationalId: r.national_id, employeeId: r.worker_id || r.national_id, name: r.name, company: r.company,
      checkInAt: r.ts_iso, checkInTime: r.time_text, checkInDate: r.date_key,
      secondsInside: Math.max(0, Math.floor((now.getTime() - Date.parse(r.ts_iso)) / 1000)),
      scannedBy: r.scanned_by, scannedUsername: r.scanned_username,
    })),
  };
}
