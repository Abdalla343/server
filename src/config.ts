import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

export interface Config {
  port: number;
  host: string;
  dataDir: string;
  jwtSecret: string;
  jwtTtl: string;
  corsOrigins: string[];
  timezone: string;
  scanWaitMs: number; batchSize: number; maxAttempts: number;
  backupIntervalMin: number; backupDailyTime: string; backupAttendanceDays: number; backupExtraDir: string;
  httpsKey: string; httpsCert: string; webDir: string;
}

/** The JWT secret lives in the data dir (never in the repo). Generated on first run. */
function loadOrCreateSecret(dataDir: string): string {
  if (process.env.JWT_SECRET) return process.env.JWT_SECRET;
  const file = path.join(dataDir, "jwt.secret");
  if (fs.existsSync(file)) return fs.readFileSync(file, "utf8").trim();
  const secret = crypto.randomBytes(48).toString("hex");
  fs.writeFileSync(file, secret, { mode: 0o600 });
  return secret;
}

export function loadConfig(overrides: Partial<Config> = {}): Config {
  const dataDir = path.resolve(overrides.dataDir ?? process.env.DATA_DIR ?? "./data");
  fs.mkdirSync(dataDir, { recursive: true });
  return {
    port: Number(process.env.PORT ?? 3000),
    host: process.env.HOST ?? "0.0.0.0",
    dataDir,
    jwtSecret: loadOrCreateSecret(dataDir),
    jwtTtl: process.env.JWT_TTL ?? "12h",
    corsOrigins: (process.env.CORS_ORIGINS ?? "").split(",").map((s) => s.trim()).filter(Boolean),
    timezone: process.env.TIMEZONE ?? "Africa/Cairo",
    scanWaitMs: Number(process.env.SCAN_WAIT_MS ?? 8000),
    batchSize: Number(process.env.BATCH_SIZE ?? 200),
    maxAttempts: Number(process.env.MAX_ATTEMPTS ?? 8),
    backupIntervalMin: Number(process.env.BACKUP_INTERVAL_MIN ?? 60),
    backupDailyTime: process.env.BACKUP_DAILY_TIME ?? "23:30",
    backupAttendanceDays: Number(process.env.BACKUP_ATTENDANCE_DAYS ?? 7),
    backupExtraDir: process.env.BACKUP_EXTRA_DIR ?? "",
    httpsKey: process.env.HTTPS_KEY ?? "",
    httpsCert: process.env.HTTPS_CERT ?? "",
    webDir: path.resolve(process.env.WEB_DIR ?? "../web/dist"),
    ...overrides,
  };
}
