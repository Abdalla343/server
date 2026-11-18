import { Router } from "express";
import { z } from "zod";
import { authenticate, hashPassword, passwordProblem, requireRole } from "../auth.js";
import { audit, type Db } from "../db.js";
import type { Config } from "../config.js";

const cols = "id, username, display_name AS displayName, role, active, created_at, last_login_at";
const displayName = z.string().trim().min(1).max(100);

export function usersRouter(cfg: Config, db: Db) {
  const r = Router();
  r.use(authenticate(cfg, db), requireRole("ADMIN"));
  r.get("/", (_q, res) => void res.json({ users: db.prepare(`SELECT ${cols} FROM users ORDER BY id`).all() }));

  r.post("/", async (req, res) => {
    const p = z.object({
      username: z.string().regex(/^[A-Za-z0-9._-]{3,32}$/),
      displayName: displayName.optional(), password: z.string().max(128), role: z.enum(["ADMIN", "USER"]).default("USER"),
    }).safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "Invalid username, display name or role" });
    const bad = passwordProblem(p.data.password);
    if (bad) return void res.status(400).json({ error: bad });
    const hash = await hashPassword(p.data.password);
    try {
      const user = db.transaction(() => {
        // The requesting admin may have been revoked while the hash was being calculated.
        const actor = db.prepare("SELECT active, role FROM users WHERE id=?").get(req.user!.id) as any;
        if (!actor?.active || actor.role !== "ADMIN") return null;
        const info = db.prepare("INSERT INTO users (username, display_name, password_hash, role) VALUES (?,?,?,?)")
          .run(p.data.username, p.data.displayName ?? p.data.username, hash, p.data.role);
        audit(db, req.user!.id, "USER_CREATED", `${p.data.username}/${p.data.role}`);
        return db.prepare(`SELECT ${cols} FROM users WHERE id=?`).get(info.lastInsertRowid);
      }).immediate();
      if (!user) return void res.status(403).json({ error: "Forbidden" });
      res.status(201).json({ user });
    } catch (e: any) {
      if (String(e.code).startsWith("SQLITE_CONSTRAINT")) return void res.status(409).json({ error: "Username already exists" });
      throw e;
    }
  });

  r.patch("/:id", async (req, res) => {
    const id = Number(req.params.id);
    const p = z.object({
      role: z.enum(["ADMIN", "USER"]).optional(), active: z.boolean().optional(),
      displayName: displayName.optional(), password: z.string().max(128).optional(),
    }).safeParse(req.body);
    if (!Number.isSafeInteger(id) || id < 1 || !p.success) return void res.status(400).json({ error: "Invalid request" });
    let hash: string | null = null;
    if (p.data.password !== undefined) {
      const bad = passwordProblem(p.data.password);
      if (bad) return void res.status(400).json({ error: bad });
      hash = await hashPassword(p.data.password);
    }
    const result = db.transaction(() => {
      const actor = db.prepare("SELECT active, role FROM users WHERE id=?").get(req.user!.id) as any;
      if (!actor?.active || actor.role !== "ADMIN") return { status: 403, body: { error: "Forbidden" } };
      const target = db.prepare("SELECT * FROM users WHERE id=?").get(id) as any;
      if (!target) return { status: 404, body: { error: "User not found" } };
      const losesAdmin = target.role === "ADMIN" && target.active &&
        ((p.data.role && p.data.role !== "ADMIN") || p.data.active === false);
      if (losesAdmin && (db.prepare("SELECT COUNT(*) n FROM users WHERE role='ADMIN' AND active=1").get() as any).n <= 1)
        return { status: 409, body: { error: "Cannot remove the last active ADMIN" } };
      db.prepare(`UPDATE users SET role=COALESCE(?,role), active=COALESCE(?,active),
        display_name=COALESCE(?,display_name), password_hash=COALESCE(?,password_hash),
        token_version=token_version+1 WHERE id=?`)
        .run(p.data.role ?? null, p.data.active === undefined ? null : Number(p.data.active), p.data.displayName ?? null, hash, id);
      audit(db, req.user!.id, "USER_UPDATED", `${target.username}: ${Object.keys(p.data).join(",")}`);
      return { status: 200, body: { user: db.prepare(`SELECT ${cols} FROM users WHERE id=?`).get(id) } };
    }).immediate();
    res.status(result.status).json(result.body);
  });

  r.delete("/:id", (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isSafeInteger(id) || id < 1) return void res.status(400).json({ error: "Invalid user ID" });
    const result = db.transaction(() => {
      const target = db.prepare("SELECT * FROM users WHERE id=?").get(id) as any;
      if (!target) return { status: 404, body: { error: "User not found" } };
      if (id === req.user!.id) return { status: 409, body: { error: "You cannot delete your own account" } };
      if (target.role === "ADMIN" && target.active &&
          (db.prepare("SELECT COUNT(*) n FROM users WHERE role='ADMIN' AND active=1").get() as any).n <= 1)
        return { status: 409, body: { error: "Cannot remove the last active ADMIN" } };
      db.prepare("DELETE FROM users WHERE id=?").run(id);
      audit(db, req.user!.id, "USER_DELETED", target.username);
      return { status: 200, body: { deleted: id } };
    }).immediate();
    res.status(result.status).json(result.body);
  });
  return r;
}
