import readline from "node:readline/promises";
import { loadConfig } from "../src/config.js";
import { openDb } from "../src/db.js";
import { hashPassword, passwordProblem } from "../src/auth.js";

// Usage: npm run create-admin -- <username>   (password is prompted, or taken from ADMIN_PASSWORD)
const username = process.argv[2];
if (!username || !/^[A-Za-z0-9._-]{3,32}$/.test(username)) { console.error("Usage: npm run create-admin -- <username>"); process.exit(1); }
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
const pw = process.env.ADMIN_PASSWORD ?? await rl.question("Password (min 10 chars): ");
rl.close();
const bad = passwordProblem(pw);
if (bad) { console.error(bad); process.exit(1); }
const cfg = loadConfig();
const db = openDb(cfg.dataDir);
try {
  db.prepare("INSERT INTO users (username, password_hash, role) VALUES (?,?, 'ADMIN')").run(username, await hashPassword(pw));
  console.log(`ADMIN '${username}' created.`);
} catch (e: any) { console.error(String(e.code).startsWith("SQLITE_CONSTRAINT") ? "Username exists." : e); process.exit(1); }
