import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import { loadConfig } from "./config.js";
import { openDb } from "./db.js";
import { createApp } from "./app.js";
import { createContext } from "./context.js";
import { createLogger } from "./log.js";
import { acquireInstanceLock } from "./lock.js";
import { Scheduler } from "./scheduler.js";
import { WorkerStore } from "./workers/store.js";

if (fs.existsSync(".env")) process.loadEnvFile(".env"); // optional server/.env (Node 22)
const cfg = loadConfig();
const log = createLogger(cfg.dataDir);
const release = acquireInstanceLock(cfg.dataDir);
const db = openDb(cfg.dataDir);
const store = new WorkerStore(cfg);
const loaded = await store.loadFromDisk();
if (loaded.ok) log.info({ ...loaded.meta }, "workers loaded");
else log.warn({ reason: loaded.errors[0]?.message }, "WORKERS NOT LOADED: scanning disabled until an ADMIN uploads a valid Workers.xlsx");

const ctx = createContext(cfg, db, store, { log });
const app = createApp(cfg, db, store, ctx);
const useHttps = Boolean(cfg.httpsKey && cfg.httpsCert);
if (!useHttps) log.warn("Running over plain HTTP: browsers will NOT allow the camera on LAN addresses. Configure HTTPS_KEY / HTTPS_CERT.");
const server = useHttps
  ? https.createServer({ key: fs.readFileSync(cfg.httpsKey), cert: fs.readFileSync(cfg.httpsCert) }, app)
  : http.createServer(app);

const recovery = ctx.processor.start();
const scheduler = new Scheduler(cfg, db, ctx.writer, ctx.processor, ctx.backup, log);
scheduler.start();
server.listen(cfg.port, cfg.host, () => log.info({ url: `${useHttps ? "https" : "http"}://${cfg.host}:${cfg.port}`, data: cfg.dataDir, recovery }, "attendance server listening"));

let closing = false;
async function shutdown(code = 0) {
  if (closing) return; closing = true;
  log.info("shutting down: finishing the current save first");
  scheduler.stop(); server.close();
  await ctx.processor.stop(); // waits for an in-flight batch write
  db.close(); release(); process.exit(code);
}
process.on("SIGINT", () => void shutdown()); process.on("SIGTERM", () => void shutdown());
// Unknown state: stop cleanly and let the service manager restart us (unfinished operations are recovered on start).
process.on("uncaughtException", (e) => { log.fatal({ err: String(e.stack ?? e) }, "uncaught exception"); void shutdown(1); });
process.on("unhandledRejection", (e) => { log.fatal({ err: String(e) }, "unhandled rejection"); void shutdown(1); });
