import type { Db } from "./db.js";
export type DuplicateMode = "block" | "warn" | "allow";
export const getSetting = (db: Db, key: string, def: string) =>
  ((db.prepare("SELECT value FROM settings WHERE key = ?").get(key) as { value: string } | undefined)?.value ?? def);
export const setSetting = (db: Db, key: string, value: string) =>
  void db.prepare("INSERT INTO settings (key, value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(key, value);
export const duplicateMode = (db: Db) => getSetting(db, "duplicateMode", "block") as DuplicateMode;
