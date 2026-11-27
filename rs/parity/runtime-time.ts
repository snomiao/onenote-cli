#!/usr/bin/env bun
// What the JavaScript runtime makes of local time, for the test in
// rs/src/time.rs that compares every time zone the binary reads with it
// (`cargo test` in rs/).
//
//   TZ=<zone> bun rs/parity/runtime-time.ts < questions > answers
//
// It is started once per zone with TZ set, as `onenote` is. One question a
// line, one answer a line:
//
//   U <seconds>                 new Date(seconds * 1000).toLocaleString(),
//                               which is how src/index.ts prints "last synced"
//   S <milliseconds> <n><unit>  what parseSince in src/index.ts makes of
//                               `--since <n><unit>` at that moment
//
// `onenote` works out one date and ends. This process works out thousands,
// and the runtime remembers the clock changes it has come across in a way
// that goes wrong where two of them are less than three weeks apart (bun
// 1.4.0: asked in a row, `S 1086148800456 1d` under America/Argentina/Tucuman
// is an hour off; asked alone it is not). So what it remembers is thrown
// away before every `S` question, by setting the zone again.

import { setTimeZone } from "bun:jsc";

const zone = process.env.TZ;
if (!zone) throw new Error("TZ has to name the zone");

/** A year of local times, to see that setting the zone again changes nothing. */
const sample = () => Array.from({ length: 366 }, (_, day) => new Date(1_700_000_000_000 + day * 86_400_000).toString()).join("\n");
const atStart = sample();
setTimeZone(zone);
if (sample() !== atStart) throw new Error(`${zone}: TZ in the environment and setTimeZone() are not the same zone`);

/** parseSince's branch for a span (src/index.ts:823), with the moment given. */
function since(now: number, value: string): string {
  const rel = value.match(/^(\d+)([dmy])$/i);
  if (!rel) return "not a span";
  const n = parseInt(rel[1], 10);
  const unit = rel[2].toLowerCase();
  const d = new Date(now);
  if (unit === "d") d.setDate(d.getDate() - n);
  else if (unit === "m") d.setMonth(d.getMonth() - n);
  else if (unit === "y") d.setFullYear(d.getFullYear() - n);
  return d.toISOString();
}

const answers: string[] = [];
for (const line of (await Bun.stdin.text()).split("\n")) {
  if (!line) continue;
  const [kind, when, value] = line.split(" ");
  if (kind === "U") answers.push(new Date(Number(when) * 1000).toLocaleString());
  else if (kind === "S") {
    setTimeZone(zone);
    answers.push(since(Number(when), value));
  }
  else answers.push(`not a question: ${line}`);
}
// Bun.write waits until everything is written; console.log may not.
await Bun.write(Bun.stdout, answers.map((answer) => `${answer}\n`).join(""));
