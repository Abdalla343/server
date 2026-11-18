import { Router } from "express";
import { z } from "zod";
import { authenticate } from "../auth.js";
import type { Context } from "../context.js";

const body = z.object({
  operationId: z.string().regex(/^[A-Za-z0-9_-]{8,64}$/),
  scanMethod: z.enum(["SCANNER", "MANUAL", "UNKNOWN"]).optional(),
  idType: z.enum(["auto", "employee"]).optional(),
  idNumber: z.string().trim().min(1).max(40),
  direction: z.enum(["IN", "OUT"]),
  gateId: z.string().trim().max(32).optional(),
});

export function scansRouter(c: Context) {
  const r = Router(), auth = authenticate(c.cfg, c.db);
  r.post("/scans", auth, async (req, res) => {
    const p = body.safeParse(req.body);
    if (!p.success) return void res.status(400).json({ error: "Invalid scan request" });
    const out = await c.scans.submit(req.user!, p.data);
    res.status(out.status).json(out.body);
  });
  r.get("/scans/:operationId", auth, (req, res) => {
    const id = String(req.params.operationId);
    const owner = c.scans.owner(id);
    if (owner === null || (req.user!.role !== "ADMIN" && owner !== req.user!.id)) return void res.status(404).json({ error: "Not found" });
    const v = c.scans.get(id)!;
    res.status((v as any).state === "COMPLETED" ? 200 : (v as any).state === "FAILED" ? 500 : 202).json(v);
  });
  return r;
}
