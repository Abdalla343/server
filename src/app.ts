import express, { type NextFunction, type Request, type Response } from "express";
import helmet from "helmet";
import cors from "cors";
import type { Config } from "./config.js";
import type { Db } from "./db.js";
import { authRouter } from "./routes/auth.js";
import { usersRouter } from "./routes/users.js";
import { workersRouter } from "./routes/workers.js";
import { WorkerStore } from "./workers/store.js";
import { createContext, type Context } from "./context.js";
import { scansRouter } from "./routes/scans.js";
import { adminRouter } from "./routes/admin.js";
import rateLimit from "express-rate-limit";
import fs from "node:fs";
import path from "node:path";

export function createApp(cfg: Config, db: Db, store: WorkerStore = new WorkerStore(cfg), ctx: Context = createContext(cfg, db, store)) {
  const app = express();
  app.disable("x-powered-by");
  const https = Boolean(cfg.httpsKey && cfg.httpsCert);
  app.use(helmet({
    contentSecurityPolicy: { directives: {
      defaultSrc: ["'self'"], scriptSrc: ["'self'", "'wasm-unsafe-eval'"], workerSrc: ["'self'", "blob:"],
      imgSrc: ["'self'", "data:", "blob:"], mediaSrc: ["'self'", "blob:"], connectSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'"], upgradeInsecureRequests: https ? [] : null } },
    hsts: https,
  }));
  app.use((_q, res, next) => { res.setHeader("Permissions-Policy", "camera=(self)"); next(); });
  // Same-origin is the default (Express serves the web app). CORS only for explicit origins.
  if (cfg.corsOrigins.length) app.use(cors({ origin: cfg.corsOrigins }));
  app.use(express.json({ limit: "100kb" }));

  app.get("/api/health", (_q, res) => void res.json({ ok: true }));
  app.use("/api", rateLimit({ windowMs: 60_000, limit: 1200, standardHeaders: true, legacyHeaders: false }));
  app.use("/api/auth", authRouter(cfg, db));
  app.use("/api/admin/users", usersRouter(cfg, db));
  app.use("/api", workersRouter(cfg, db, store));
  app.use("/api", scansRouter(ctx));
  app.use("/api", adminRouter(ctx));
  app.use("/api", (_q, res) => void res.status(404).json({ error: "Not found" }));

  // Built web app (same origin => no CORS needed). Only this folder is served; the data folder is never exposed.
  if (fs.existsSync(path.join(cfg.webDir, "index.html"))) {
    app.use(express.static(cfg.webDir));
    app.use((req, res, next) => (req.method === "GET" ? res.sendFile(path.join(cfg.webDir, "index.html")) : next()));
  }

  app.use((err: any, req: Request, res: Response, _n: NextFunction) => {
    if (err?.type === "entity.parse.failed") return void res.status(400).json({ error: "Malformed JSON" });
    ctx.log.error({ err: String(err?.stack ?? err), path: req.path }, "unhandled error");
    res.status(500).json({ error: "Internal error" });
  });
  return app;
}
