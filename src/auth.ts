import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import type { NextFunction, Request, Response } from "express";
import type { Config } from "./config.js";
import type { Db } from "./db.js";

export type Role = "ADMIN" | "USER";
export interface AuthUser { id: number; username: string; displayName?: string; role: Role }
declare module "express-serve-static-core" { interface Request { user?: AuthUser } }

const COST = 12;
export const hashPassword = (pw: string) => bcrypt.hash(pw, COST);
// Used to equalize timing when the username does not exist.
export const DUMMY_HASH = bcrypt.hashSync("dummy-password-for-timing", COST);

export function passwordProblem(pw: string): string | null {
  if (pw.length < 10) return "Password must be at least 10 characters";
  if (Buffer.byteLength(pw, "utf8") > 72) return "Password must be at most 72 UTF-8 bytes";
  return null;
}

export function signToken(cfg: Config, u: { id: number; token_version: number }): string {
  return jwt.sign({ sub: String(u.id), tv: u.token_version }, cfg.jwtSecret, {
    expiresIn: cfg.jwtTtl as jwt.SignOptions["expiresIn"],
    algorithm: "HS256",
  });
}

/** Verifies the JWT, then re-reads the user so role/disable/password-reset apply immediately. */
export function authenticate(cfg: Config, db: Db) {
  const q = db.prepare("SELECT id, username, display_name, role, active, token_version FROM users WHERE id = ?");
  return (req: Request, res: Response, next: NextFunction) => {
    const h = req.headers.authorization;
    if (!h?.startsWith("Bearer ")) return void res.status(401).json({ error: "Authentication required" });
    try {
      const p = jwt.verify(h.slice(7), cfg.jwtSecret, { algorithms: ["HS256"] }) as jwt.JwtPayload;
      const u = q.get(Number(p.sub)) as { id: number; username: string; display_name: string; role: Role; active: number; token_version: number } | undefined;
      if (!u || !u.active || u.token_version !== p.tv) return void res.status(401).json({ error: "Session no longer valid" });
      req.user = { id: u.id, username: u.username, displayName: u.display_name || u.username, role: u.role };
      next();
    } catch {
      res.status(401).json({ error: "Invalid or expired token" });
    }
  };
}

export const requireRole = (...roles: Role[]) => (req: Request, res: Response, next: NextFunction) =>
  req.user && roles.includes(req.user.role) ? next() : void res.status(403).json({ error: "Forbidden" });
