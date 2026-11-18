import type { Context } from "./context.js";
import { audit } from "./db.js";

/** Database tombstones and Excel cleanup jobs commit together, under the single writer lock. */
export async function deleteOperations(c: Context, userId: number, ids: string[] | "all") {
  const count = await c.processor.exclusive(() => c.db.transaction(() => {
    const actor = c.db.prepare("SELECT role, active FROM users WHERE id=?").get(userId) as any;
    if (!actor?.active || actor.role !== "ADMIN") throw new Error("Admin session is no longer active");
    c.db.exec("CREATE TEMP TABLE IF NOT EXISTS deletion_selection(operation_id TEXT PRIMARY KEY); DELETE FROM deletion_selection;");
    const select = c.db.prepare("INSERT OR IGNORE INTO deletion_selection VALUES (?)");
    if (ids === "all") c.db.exec("INSERT INTO deletion_selection SELECT operation_id FROM active_operations");
    else ids.forEach((id) => select.run(id));
    const files = c.db.prepare(`SELECT DISTINCT CASE WHEN outcome='APPROVED' THEN direction ELSE 'REJECTED' END direction,
      CASE WHEN outcome='APPROVED' THEN date_key ELSE '' END date_key
      FROM active_operations WHERE operation_id IN (SELECT operation_id FROM deletion_selection)`).all() as any[];
    const result = c.db.prepare("UPDATE operations SET deleted_at=? WHERE deleted_at IS NULL AND operation_id IN (SELECT operation_id FROM deletion_selection)")
      .run(c.clock().toISOString());
    const job = c.db.prepare(`INSERT INTO excel_deletions(file_key,direction,date_key) VALUES(?,?,?)
      ON CONFLICT(file_key) DO UPDATE SET state='PENDING',attempts=0,next_attempt_at=0,last_error=NULL`);
    files.forEach((file) => job.run(file.direction + "/" + file.date_key, file.direction, file.date_key));
    audit(c.db, userId, "OPERATIONS_DELETED", JSON.stringify({ scope: ids === "all" ? "all dates" : ids, count: result.changes }));
    return result.changes;
  }).immediate());
  c.processor.notifyDeleted();
  c.processor.kick();
  // Wait for this attempt to synchronize while leaving retries durable if a workbook is locked.
  await c.processor.exclusive(() => c.processor.flushDeletions());
  return { deleted: count, excelPending: c.processor.deletionStatus() };
}
