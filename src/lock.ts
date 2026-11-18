import fs from "node:fs";
import path from "node:path";

/** Single-instance guard: only one backend may ever write the Excel files. */
export function acquireInstanceLock(dataDir: string): () => void {
  const file = path.join(dataDir, "server.lock");
  const tryCreate = () => { const fd = fs.openSync(file, "wx"); fs.writeSync(fd, String(process.pid)); fs.closeSync(fd); };
  try { tryCreate(); } catch (e: any) {
    if (e.code !== "EEXIST") throw e;
    const pid = Number(fs.readFileSync(file, "utf8"));
    let alive = false;
    try { process.kill(pid, 0); alive = pid !== process.pid; } catch (k: any) { alive = k.code === "EPERM"; }
    if (alive) throw new Error(`Another server instance is already running (PID ${pid}). Refusing to start.`);
    fs.rmSync(file); // stale lock from a crash
    tryCreate();
  }
  let released = false;
  const release = () => { if (!released) { released = true; try { fs.rmSync(file); } catch {} } };
  process.on("exit", release);
  return release;
}
