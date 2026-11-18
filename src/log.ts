import pino from "pino";
import fs from "node:fs";
import path from "node:path";

export type Log = pino.Logger;
export function createLogger(dataDir: string | null): Log {
  if (!dataDir) return pino({ level: "silent" });
  const dir = path.join(dataDir, "logs"); fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, "server.log");
  try { if (fs.statSync(file).size > 10 * 1024 * 1024) fs.renameSync(file, path.join(dir, `server-${Date.now()}.log`)); } catch {}
  // keep the 10 newest rotated logs
  const old = fs.readdirSync(dir).filter((f) => /^server-\d+\.log$/.test(f)).sort();
  old.slice(0, Math.max(0, old.length - 10)).forEach((f) => fs.rmSync(path.join(dir, f)));
  return pino({ level: process.env.LOG_LEVEL ?? "info" },
    pino.multistream([{ stream: process.stdout }, { stream: pino.destination({ dest: file, sync: true }) }]));
}
