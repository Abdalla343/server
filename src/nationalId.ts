/** Digit + Egyptian National ID helpers. */
const AR = /[\u0660-\u0669]/g, FA = /[\u06F0-\u06F9]/g;
/** Converts Arabic-Indic / Persian digits to 0-9. */
export const toWesternDigits = (s: string) =>
  s.replace(AR, (c) => String(c.charCodeAt(0) - 0x0660)).replace(FA, (c) => String(c.charCodeAt(0) - 0x06f0));

/** Normalizes a raw National ID: Western digits, no spaces/separators. Returns "" if nothing usable. */
export const normalizeNid = (raw: string) => toWesternDigits(raw).replace(/[\s\u200e\u200f\u00a0-]/g, "");
/**
 * Canonical form of a company worker code (the "ID Card" column, or a scanned code such as IR1234 / ZTC535 / HC123456):
 * Western digits, UPPER CASE, letters and digits only. "ir-1234", "IR 1234" and "ير١٢٣٤"-style digit mixes all compare equal.
 */
export const normalizeCode = (raw: string) => toWesternDigits(raw).toUpperCase().replace(/[^A-Z0-9]/g, "");
export const isFormat14 = (s: string) => /^\d{14}$/.test(s);

const GOV = new Set(["01","02","03","04","11","12","13","14","15","16","17","18","19","21","22","23","24","25","26","27","28","29","31","32","33","34","35","88"]);

/** Semantic check: century digit, real birth date (not future), governorate code. No official checksum exists. */
export function nidSemanticProblem(nid: string, now = new Date()): string | null {
  if (!isFormat14(nid)) return "must be exactly 14 digits";
  const c = nid[0]!;
  if (c !== "2" && c !== "3") return "century digit must be 2 or 3";
  const y = (c === "2" ? 1900 : 2000) + Number(nid.slice(1, 3)), m = Number(nid.slice(3, 5)), d = Number(nid.slice(5, 7));
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return "invalid birth date";
  if (dt.getTime() > now.getTime()) return "birth date is in the future";
  if (!GOV.has(nid.slice(7, 9))) return "unknown governorate code";
  return null;
}

export const maskNid = (nid: string) => (nid.length === 14 ? `${nid.slice(0, 3)}********${nid.slice(-3)}` : "***");
