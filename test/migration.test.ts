import { describe, expect, it } from "vitest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { openDb } from "../src/db.js";
import { AttendanceWriter, rejectedPath } from "../src/excel/attendance.js";
import { QueueProcessor } from "../src/queue/processor.js";
import ExcelJS from "exceljs";

describe("upgrade from schema v2", () => {
  it("preserves existing operations and users, backfills operator snapshots and exports historical rejections once", async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "migration-"));
    let db = openDb(dir);
    // Recreate the exact v2 column layout before populating legacy data.
    db.exec(`DROP VIEW current_presence; DROP VIEW active_operations;
      DROP TABLE excel_deletions; DROP INDEX idx_ops_active_worker; DROP INDEX idx_ops_presence;
      ALTER TABLE operations DROP COLUMN scan_method;
      ALTER TABLE operations DROP COLUMN deleted_at;
      ALTER TABLE users DROP COLUMN display_name;
      ALTER TABLE operations DROP COLUMN id_type;
      ALTER TABLE operations DROP COLUMN scanned_by;
      ALTER TABLE operations DROP COLUMN scanned_username;
      PRAGMA user_version=2;
      INSERT INTO users(username,password_hash,role) VALUES ('oldgate','old-bcrypt-hash','USER');
      INSERT INTO operations(operation_id,direction,id_input,national_id,name,company,outcome,state,user_id,ts_iso,date_key,day,time_text)
      VALUES ('legacy-approved','IN','A1','29501010100001','Worker One','Old Company','APPROVED','COMPLETED',1,'2026-11-14T05:00:00Z','14-11-2026','2026-11-14','07:00'),
             ('legacy-rejected','OUT','XYZ',NULL,NULL,NULL,'NOT_FOUND','COMPLETED',1,'2026-11-14T06:00:00Z','14-11-2026','2026-11-14','08:00');`);
    db.close();
    db = openDb(dir);
    expect((db.prepare("SELECT username,display_name,password_hash FROM users").get())).toEqual({ username: "oldgate", display_name: "oldgate", password_hash: "old-bcrypt-hash" });
    expect(db.prepare("SELECT state,scanned_by,scanned_username FROM operations WHERE operation_id='legacy-rejected'").get())
      .toEqual({ state: "PENDING", scanned_by: "oldgate", scanned_username: "oldgate" });
    expect((db.prepare("SELECT COUNT(*) n FROM current_presence").get() as any).n).toBe(1);
    const processor = new QueueProcessor(db, new AttendanceWriter(dir));
    processor.start(); expect(await processor.waitIdle()).toBe(true); await processor.stop();
    db.close();
    db = openDb(dir);
    const again = new QueueProcessor(db, new AttendanceWriter(dir));
    again.start(); expect(await again.waitIdle()).toBe(true); await again.stop();
    const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(rejectedPath(dir));
    const sheet = wb.worksheets[0]!;
    expect(sheet.actualRowCount).toBe(2);
    expect(sheet.getRow(2).getCell(1).value).toBe("XYZ");
    expect(sheet.getRow(2).values).toEqual(expect.arrayContaining(["Not Found", "Check-Out", "oldgate"]));
    db.close();
  });
});
