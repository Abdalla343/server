import { Router } from "express";
import rateLimit from "express-rate-limit";
import bcrypt from "bcryptjs";
import { z } from "zod";
import { DUMMY_HASH, authenticate, hashPassword, passwordProblem, signToken } from "../auth.js";
import { audit, type Db } from "../db.js";
import type { Config } from "../config.js";

export function authRouter(cfg: Config, db: Db) {
  const r = Router();
  const limiter = rateLimit({ windowMs: 15 * 60_000, limit: 20, standardHeaders: true, legacyHeaders: false,
    message: { error: "Too many login attempts, try again later" } });
  const body = z.object({ username: z.string().min(1).max(64), password: z.string().min(1).max(128) });

  r.post("/login", limiter, async (req, res) => {
    const p = body.safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "Invalid request" });
    const u = db.prepare("SELECT * FROM users WHERE username = ?").get(p.data.username) as any;
    const ok = await bcrypt.compare(p.data.password, u?.password_hash ?? DUMMY_HASH);
    if (!u || !u.active || !ok) {
      audit(db, u?.id ?? null, "LOGIN_FAILED", p.data.username.slice(0, 64));
      return void res.status(401).json({ error: "Invalid username or password" });
    }
    db.prepare("UPDATE users SET last_login_at = strftime('%Y-%m-%dT%H:%M:%fZ','now') WHERE id = ?").run(u.id);
    audit(db, u.id, "LOGIN");
    res.json({ token: signToken(cfg, u), user: { id: u.id, username: u.username, displayName: u.display_name || u.username, role: u.role } });
  });

  const auth = authenticate(cfg, db);
  r.get("/me", auth, (req, res) => void res.json({ user: req.user }));

  // A user may change their own password; this invalidates all their existing tokens.
  r.post("/change-password", auth, async (req, res) => {
    const p = z.object({ currentPassword: z.string().max(128), newPassword: z.string().max(128) }).safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "Invalid request" });
    const bad = passwordProblem(p.data.newPassword);
    if (bad) return void res.status(400).json({ error: bad });
    const u = db.prepare("SELECT * FROM users WHERE id = ?").get(req.user!.id) as any;
    if (!(await bcrypt.compare(p.data.currentPassword, u.password_hash)))
      return void res.status(403).json({ error: "Current password is incorrect" });
    db.prepare("UPDATE users SET password_hash = ?, token_version = token_version + 1 WHERE id = ?")
      .run(await hashPassword(p.data.newPassword), u.id);
    audit(db, u.id, "PASSWORD_CHANGED");
    const fresh = db.prepare("SELECT id, token_version FROM users WHERE id = ?").get(u.id) as any;
    res.json({ token: signToken(cfg, fresh) });
  });
  return r;
}
