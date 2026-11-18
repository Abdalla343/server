import { describe, it, expect, beforeEach } from "vitest";
import request from "supertest";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createApp } from "../src/app.js";
import { loadConfig } from "../src/config.js";
import { openDb } from "../src/db.js";
import { hashPassword } from "../src/auth.js";
import { acquireInstanceLock } from "../src/lock.js";

async function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "att-"));
  const cfg = loadConfig({ dataDir: dir });
  const db = openDb(dir);
  db.prepare("INSERT INTO users (username,password_hash,role) VALUES ('boss',?, 'ADMIN')").run(await hashPassword("correct-horse-1"));
  return { dir, cfg, db, app: createApp(cfg, db) };
}
const login = async (app: any, u: string, p: string) => request(app).post("/api/auth/login").send({ username: u, password: p });

describe("auth", () => {
  let s: Awaited<ReturnType<typeof setup>>;
  beforeEach(async () => { s = await setup(); });

  it("rejects bad credentials generically and accepts good ones", async () => {
    expect((await login(s.app, "boss", "wrong-password")).status).toBe(401);
    expect((await login(s.app, "nobody", "whatever-123")).status).toBe(401);
    const ok = await login(s.app, "BOSS", "correct-horse-1"); // case-insensitive username
    expect(ok.status).toBe(200);
    expect(ok.body.user.role).toBe("ADMIN");
    expect(JSON.stringify(ok.body)).not.toContain("password");
  });

  it("protects endpoints and rejects forged tokens", async () => {
    expect((await request(s.app).get("/api/auth/me")).status).toBe(401);
    expect((await request(s.app).get("/api/admin/users").set("Authorization", "Bearer abc.def.ghi")).status).toBe(401);
  });

  it("enforces roles on the backend", async () => {
    const admin = (await login(s.app, "boss", "correct-horse-1")).body.token;
    const c = await request(s.app).post("/api/admin/users").set("Authorization", `Bearer ${admin}`)
      .send({ username: "gate1", password: "gate-password-1", role: "USER" });
    expect(c.status).toBe(201);
    expect((await request(s.app).post("/api/admin/users").set("Authorization", `Bearer ${admin}`)
      .send({ username: "GATE1", password: "gate-password-1", role: "USER" })).status).toBe(409);
    const user = (await login(s.app, "gate1", "gate-password-1")).body.token;
    expect((await request(s.app).get("/api/auth/me").set("Authorization", `Bearer ${user}`)).status).toBe(200);
    expect((await request(s.app).get("/api/admin/users").set("Authorization", `Bearer ${user}`)).status).toBe(403);
  });

  it("revokes tokens immediately on disable, and keeps at least one ADMIN", async () => {
    const admin = (await login(s.app, "boss", "correct-horse-1")).body.token;
    const H = { Authorization: `Bearer ${admin}` };
    const id = (await request(s.app).post("/api/admin/users").set(H).send({ username: "gate2", password: "gate-password-2", role: "USER" })).body.user.id;
    const ut = (await login(s.app, "gate2", "gate-password-2")).body.token;
    expect((await request(s.app).patch(`/api/admin/users/${id}`).set(H).send({ active: false })).status).toBe(200);
    expect((await request(s.app).get("/api/auth/me").set("Authorization", `Bearer ${ut}`)).status).toBe(401);
    expect((await login(s.app, "gate2", "gate-password-2")).status).toBe(401);
    expect((await request(s.app).patch("/api/admin/users/1").set(H).send({ role: "USER" })).status).toBe(409);
    expect((await request(s.app).patch("/api/admin/users/1").set(H).send({ active: false })).status).toBe(409);
  });

  it("password change invalidates the old token and enforces length", async () => {
    const old = (await login(s.app, "boss", "correct-horse-1")).body.token;
    const H = { Authorization: `Bearer ${old}` };
    expect((await request(s.app).post("/api/auth/change-password").set(H).send({ currentPassword: "nope", newPassword: "new-password-12" })).status).toBe(403);
    expect((await request(s.app).post("/api/auth/change-password").set(H).send({ currentPassword: "correct-horse-1", newPassword: "short" })).status).toBe(400);
    const r = await request(s.app).post("/api/auth/change-password").set(H).send({ currentPassword: "correct-horse-1", newPassword: "new-password-12" });
    expect(r.status).toBe(200);
    expect((await request(s.app).get("/api/auth/me").set(H)).status).toBe(401);
    expect((await request(s.app).get("/api/auth/me").set("Authorization", `Bearer ${r.body.token}`)).status).toBe(200);
  });

  it("stores only bcrypt hashes and persists the JWT secret", async () => {
    const row = s.db.prepare("SELECT password_hash FROM users").get() as any;
    expect(row.password_hash).toMatch(/^\$2[aby]\$12\$/);
    expect(loadConfig({ dataDir: s.dir }).jwtSecret).toBe(s.cfg.jwtSecret);
  });

  it("returns 400 on malformed JSON", async () => {
    const r = await request(s.app).post("/api/auth/login").set("Content-Type", "application/json").send("{bad");
    expect(r.status).toBe(400);
  });
});

describe("single-instance lock", () => {
  it("blocks a second live instance and recovers a stale lock", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "lock-"));
    fs.writeFileSync(path.join(dir, "server.lock"), "2147483000"); // dead PID -> stale
    const rel = acquireInstanceLock(dir);
    fs.writeFileSync(path.join(dir, "server.lock"), String(process.ppid)); // live other PID
    expect(() => acquireInstanceLock(dir)).toThrow(/already running/);
    rel();
  });
});

describe("security user lifecycle", () => {
  it("creates display names, hashes passwords, restricts admin operations and revokes deleted sessions", async () => {
    const s = await setup(), token = (await login(s.app, "boss", "correct-horse-1")).body.token;
    const H = { Authorization: `Bearer ${token}` };
    const created = await request(s.app).post("/api/admin/users").set(H)
      .send({ username: "security1", displayName: "Mahmoud Gate", password: "gate-password-12" });
    expect(created.status).toBe(201); expect(created.body.user).toMatchObject({ displayName: "Mahmoud Gate", role: "USER" });
    expect(created.body.user.password_hash).toBeUndefined();
    const gate = await login(s.app, "security1", "gate-password-12"), G = { Authorization: `Bearer ${gate.body.token}` };
    expect(gate.body.user.displayName).toBe("Mahmoud Gate");
    for (const url of ["/api/admin/users", "/api/admin/onsite", "/api/dashboard/summary", "/api/attendance/rejected", "/api/queue/status"])
      expect((await request(s.app).get(url).set(G)).status, url).toBe(403);
    expect((await request(s.app).post("/api/admin/users").set(G).send({ username: "attack", password: "password-123", role: "ADMIN" })).status).toBe(403);
    expect((await request(s.app).delete("/api/admin/users/1").set(G)).status).toBe(403);
    expect((await request(s.app).delete("/api/admin/users/1").set(H)).status).toBe(409);
    const id = created.body.user.id;
    expect((await request(s.app).delete("/api/admin/users/" + id).set(H)).status).toBe(200);
    expect((await request(s.app).get("/api/auth/me").set(G)).status).toBe(401);
    expect((await login(s.app, "security1", "gate-password-12")).status).toBe(401);
    expect((await request(s.app).get("/api/admin/users").set(H)).body.users).toHaveLength(1);
    s.db.close();
  });

  it("rejects passwords that bcrypt would truncate, including multibyte text", async () => {
    const s = await setup(), token = (await login(s.app, "boss", "correct-horse-1")).body.token;
    for (const password of ["a".repeat(73), "م".repeat(37)])
      expect((await request(s.app).post("/api/admin/users").set("Authorization", `Bearer ${token}`)
        .send({ username: "security2", displayName: "Gate Two", password })).status).toBe(400);
    s.db.close();
  });
});
