// Fallback HTTPS certificate (self-signed). Prefer mkcert (see README) because phones then trust it automatically.
// Usage: node scripts/gen-cert.mjs 192.168.1.100 [hostname]
import selfsigned from "selfsigned";
import fs from "node:fs"; import path from "node:path";
const ip = process.argv[2], host = process.argv[3] ?? "localhost";
if (!ip) { console.error("Usage: node scripts/gen-cert.mjs <server-LAN-IP> [hostname]"); process.exit(1); }
const pems = await selfsigned.generate([{ name: "commonName", value: host }], {
  days: 825, keySize: 2048, algorithm: "sha256",
  extensions: [{ name: "subjectAltName", altNames: [{ type: 2, value: host }, { type: 2, value: "localhost" }, { type: 7, ip }, { type: 7, ip: "127.0.0.1" }] }],
});
const dir = path.resolve(process.env.DATA_DIR ?? "./data", "tls"); fs.mkdirSync(dir, { recursive: true });
fs.writeFileSync(path.join(dir, "server.key"), pems.private, { mode: 0o600 }); fs.writeFileSync(path.join(dir, "server.crt"), pems.cert);
console.log(`Wrote ${path.join(dir, "server.key")} and ${path.join(dir, "server.crt")}\nSet HTTPS_KEY and HTTPS_CERT to these paths. Install server.crt as a trusted certificate on every gate device.`);
