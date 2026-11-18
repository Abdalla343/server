/** Local date/time helpers. All business dates use the configured timezone (default Africa/Cairo). */
export interface LocalParts { date: string; time: string } // date = DD-MM-YYYY, time = HH:mm

export function localParts(d: Date, tz: string): LocalParts {
  const p = Object.fromEntries(
    new Intl.DateTimeFormat("en-GB", { timeZone: tz, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" })
      .formatToParts(d).map((x) => [x.type, x.value]),
  ) as Record<string, string>;
  return { date: `${p.day}-${p.month}-${p.year}`, time: `${p.hour}:${p.minute}` };
}
