import Database from "better-sqlite3";
import path from "node:path";

export type Db = Database.Database;

const MIGRATIONS: string[] = [
  `CREATE TABLE users (
     id INTEGER PRIMARY KEY AUTOINCREMENT,
     username TEXT NOT NULL UNIQUE COLLATE NOCASE,
     password_hash TEXT NOT NULL,
     role TEXT NOT NULL CHECK (role IN ('ADMIN','USER')),
     active INTEGER NOT NULL DEFAULT 1,
     token_version INTEGER NOT NULL DEFAULT 0,
     created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
     last_login_at TEXT
   );
   CREATE TABLE audit_log (
     id INTEGER PRIMARY KEY AUTOINCREMENT,
     ts TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
     user_id INTEGER, action TEXT NOT NULL, detail TEXT
   );
   CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);`,
  // v2: persistent scan queue + full scan audit (every scan attempt, including rejected ones)
  `CREATE TABLE operations (
     id INTEGER PRIMARY KEY AUTOINCREMENT,
     operation_id TEXT NOT NULL UNIQUE,
     direction TEXT NOT NULL CHECK (direction IN ('IN','OUT')),
     id_input TEXT NOT NULL,
     national_id TEXT, worker_id TEXT, name TEXT, job_title TEXT, company TEXT,
     worker_status TEXT, termination_reason TEXT,
     outcome TEXT NOT NULL CHECK (outcome IN ('APPROVED','NOT_APPROVED','NOT_FOUND','ALREADY_REGISTERED')),
     warning TEXT, note TEXT,
     state TEXT NOT NULL CHECK (state IN ('PENDING','PROCESSING','COMPLETED','FAILED')),
     attempts INTEGER NOT NULL DEFAULT 0, last_error TEXT,
     next_attempt_at INTEGER NOT NULL DEFAULT 0,
     gate_id TEXT, user_id INTEGER,
     ts_iso TEXT NOT NULL, date_key TEXT NOT NULL, day TEXT NOT NULL, time_text TEXT NOT NULL,
     created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
     completed_at TEXT
   );
   CREATE INDEX idx_ops_queue ON operations (state, next_attempt_at);
   CREATE INDEX idx_ops_worker_day ON operations (national_id, day, id);
   CREATE INDEX idx_ops_day ON operations (day, direction);`,
  // v3: immutable operator snapshots, rejected-file backfill and indexed live presence.
  `ALTER TABLE users ADD COLUMN display_name TEXT NOT NULL DEFAULT '';
   UPDATE users SET display_name=username WHERE display_name='';
   ALTER TABLE operations ADD COLUMN id_type TEXT NOT NULL DEFAULT 'auto';
   ALTER TABLE operations ADD COLUMN scanned_by TEXT NOT NULL DEFAULT '';
   ALTER TABLE operations ADD COLUMN scanned_username TEXT NOT NULL DEFAULT '';
   UPDATE operations SET scanned_by=COALESCE((SELECT display_name FROM users WHERE users.id=operations.user_id),''),
     scanned_username=COALESCE((SELECT username FROM users WHERE users.id=operations.user_id),'');
   UPDATE operations SET state='PENDING', completed_at=NULL, next_attempt_at=0
     WHERE outcome IN ('NOT_APPROVED','NOT_FOUND','ALREADY_REGISTERED');
   CREATE INDEX idx_ops_presence ON operations(national_id,id)
     WHERE outcome='APPROVED' AND state='COMPLETED';
   CREATE VIEW current_presence AS
   SELECT entry.* FROM operations entry
   WHERE entry.outcome='APPROVED' AND entry.state='COMPLETED' AND entry.direction='IN'
     AND entry.id=(
       SELECT MIN(o.id) FROM operations o
       WHERE o.national_id=entry.national_id AND o.outcome='APPROVED' AND o.state='COMPLETED'
         AND o.direction='IN' AND o.id > COALESCE((
           SELECT MAX(ex.id) FROM operations ex
           WHERE ex.national_id=entry.national_id AND ex.outcome='APPROVED'
             AND ex.state='COMPLETED' AND ex.direction='OUT'
         ),0)
     );`,
  // v4: submission method, admin deletion tombstones and durable Excel cleanup.
  `ALTER TABLE operations ADD COLUMN scan_method TEXT NOT NULL DEFAULT 'UNKNOWN'
     CHECK(scan_method IN ('SCANNER','MANUAL','UNKNOWN'));
   ALTER TABLE operations ADD COLUMN deleted_at TEXT;
   CREATE VIEW active_operations AS SELECT * FROM operations WHERE deleted_at IS NULL;
   CREATE TABLE excel_deletions (
     file_key TEXT PRIMARY KEY, direction TEXT NOT NULL, date_key TEXT NOT NULL,
     state TEXT NOT NULL DEFAULT 'PENDING', attempts INTEGER NOT NULL DEFAULT 0,
     next_attempt_at INTEGER NOT NULL DEFAULT 0, last_error TEXT
   );
   CREATE INDEX idx_ops_active_worker ON operations(national_id,id) WHERE deleted_at IS NULL;
   DROP VIEW current_presence;
   CREATE VIEW current_presence AS
   SELECT entry.* FROM active_operations entry
   WHERE entry.outcome='APPROVED' AND entry.state='COMPLETED' AND entry.direction='IN'
     AND entry.id=(
       SELECT MIN(o.id) FROM active_operations o
       WHERE o.national_id=entry.national_id AND o.outcome='APPROVED' AND o.state='COMPLETED'
         AND o.direction='IN' AND o.id > COALESCE((
           SELECT MAX(ex.id) FROM active_operations ex
           WHERE ex.national_id=entry.national_id AND ex.outcome='APPROVED'
             AND ex.state='COMPLETED' AND ex.direction='OUT'
         ),0)
     );`,

];

export function openDb(dataDir: string | ":memory:"): Db {
  const db = new Database(dataDir === ":memory:" ? ":memory:" : path.join(dataDir, "app.db"));
  db.pragma("journal_mode = WAL");
  db.pragma("synchronous = FULL"); // durability over speed: the queue must survive power loss
  db.pragma("foreign_keys = ON");
  db.pragma("busy_timeout = 5000");
  migrate(db);
  return db;
}

function migrate(db: Db) {
  const current = db.pragma("user_version", { simple: true }) as number;
  for (let v = current; v < MIGRATIONS.length; v++) {
    db.transaction(() => {
      db.exec(MIGRATIONS[v]!);
      db.pragma(`user_version = ${v + 1}`);
    })();
  }
}

export function audit(db: Db, userId: number | null, action: string, detail?: string) {
  db.prepare("INSERT INTO audit_log (user_id, action, detail) VALUES (?,?,?)").run(userId, action, detail ?? null);
}
