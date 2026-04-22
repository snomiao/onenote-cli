import React, { useState, useEffect } from "react";
import { Box, Text, useApp } from "ink";

export type SyncEvent =
  | { type: "total"; total: number; notebooks: number }
  | { type: "section"; index: number; notebook: string; section: string; size?: string }
  | { type: "retag"; index: number; notebook: string; section: string }
  | { type: "tags"; count: number }
  | { type: "done"; notebook: string; section: string; pages: number; status: "ok" | "failed" | "skipped" }
  | { type: "complete"; downloaded: number; retagged: number; skipped: number };

interface State {
  total: number;
  notebooks: number;
  done: number;
  downloaded: number;
  retagged: number;
  skipped: number;
  current: { notebook: string; section: string; size?: string; retag?: boolean } | null;
  log: string[];
  complete: boolean;
  startMs: number;
}

const BAR_WIDTH = 24;

function bar(ratio: number): string {
  const filled = Math.round(ratio * BAR_WIDTH);
  return "█".repeat(filled) + "░".repeat(BAR_WIDTH - filled);
}

function fmtTime(ms: number): string {
  const s = Math.round(ms / 1000);
  if (s < 60) return `${s}s`;
  return `${Math.floor(s / 60)}m${s % 60}s`;
}

function SyncUI({ events }: { events: AsyncIterable<SyncEvent> }) {
  const { exit } = useApp();
  const [state, setState] = useState<State>({
    total: 0, notebooks: 0, done: 0, downloaded: 0, retagged: 0, skipped: 0,
    current: null, log: [], complete: false, startMs: Date.now(),
  });

  useEffect(() => {
    (async () => {
      for await (const ev of events) {
        setState((s) => {
          switch (ev.type) {
            case "total":
              return { ...s, total: ev.total, notebooks: ev.notebooks };
            case "section":
              return { ...s, current: { notebook: ev.notebook, section: ev.section, size: ev.size } };
            case "retag":
              return { ...s, current: { notebook: ev.notebook, section: ev.section, retag: true } };
            case "tags": {
              const last = s.log[s.log.length - 1];
              const next = last?.startsWith("  tags: ")
                ? [...s.log.slice(0, -1), `  tags: ${ev.count} pages tagged`]
                : [...s.log, `  tags: ${ev.count} pages tagged`];
              return { ...s, log: next.slice(-6) };
            }
            case "done": {
              const icon = ev.status === "ok" ? "✓" : ev.status === "failed" ? "✗" : "–";
              const line = `${icon} ${ev.notebook}/${ev.section} (${ev.pages} pages)`;
              return {
                ...s,
                done: s.done + 1,
                downloaded: s.downloaded + (ev.status === "ok" ? 1 : 0),
                skipped: s.skipped + (ev.status === "skipped" ? 1 : 0),
                log: [...s.log, line].slice(-6),
                current: null,
              };
            }
            case "complete":
              return { ...s, complete: true, downloaded: ev.downloaded, retagged: ev.retagged, skipped: ev.skipped };
          }
        });
      }
      exit();
    })();
  }, []);

  const elapsed = Date.now() - state.startMs;
  const ratio = state.total > 0 ? state.done / state.total : 0;
  const speed = state.done > 0 ? state.done / (elapsed / 1000) : 0;
  const eta = speed > 0 && !state.complete ? (state.total - state.done) / speed : 0;
  const pct = Math.round(ratio * 100);

  return (
    <Box flexDirection="column" paddingTop={1}>
      {/* Header */}
      <Box>
        <Text bold color="cyan">Syncing OneNote  </Text>
        <Text color="green">{bar(ratio)}</Text>
        <Text>  </Text>
        <Text bold>{state.done}/{state.total}</Text>
        <Text color="gray">  {pct}%</Text>
      </Box>

      {/* Current section */}
      {state.current && (
        <Box marginTop={1}>
          <Text color="yellow">  ⟳ </Text>
          <Text>{state.current.notebook}/</Text>
          <Text bold>{state.current.section}</Text>
          {state.current.size && <Text color="gray">  {state.current.size}</Text>}
          {state.current.retag && <Text color="blue">  (retag)</Text>}
        </Box>
      )}

      {/* Speed / ETA */}
      <Box marginTop={state.current ? 0 : 1}>
        <Text color="gray">  </Text>
        <Text color="gray">speed: </Text>
        <Text>{speed > 0 ? `${speed.toFixed(1)} sec/s` : "–"}</Text>
        <Text color="gray">  elapsed: </Text>
        <Text>{fmtTime(elapsed)}</Text>
        {!state.complete && eta > 0 && (
          <>
            <Text color="gray">  ETA: </Text>
            <Text color="yellow">{fmtTime(eta * 1000)}</Text>
          </>
        )}
      </Box>

      {/* Recent log */}
      <Box flexDirection="column" marginTop={1}>
        {state.log.map((line, i) => (
          <Text key={i} color={line.startsWith("✗") ? "red" : line.startsWith("✓") ? "green" : "gray"}>
            {"  "}{line}
          </Text>
        ))}
      </Box>

      {/* Complete summary */}
      {state.complete && (
        <Box marginTop={1}>
          <Text bold color="green">✓ Sync complete  </Text>
          <Text color="gray">
            {state.downloaded} downloaded  {state.retagged} retagged  {state.skipped} up-to-date
            {"  in "}{fmtTime(elapsed)}
          </Text>
        </Box>
      )}
    </Box>
  );
}

export async function runSyncUI(
  syncFn: (emit: (ev: SyncEvent) => void) => Promise<void>
): Promise<void> {
  // Fall back to plain text if not a TTY
  if (!process.stdout.isTTY) {
    const emit = (ev: SyncEvent) => {
      if (ev.type === "total") process.stdout.write(`${ev.total} sections across ${ev.notebooks} notebooks\n`);
      if (ev.type === "section") process.stdout.write(`  [${ev.index}] ${ev.notebook}/${ev.section}${ev.size ? ` (${ev.size})` : ""}\n`);
      if (ev.type === "retag") process.stdout.write(`  [${ev.index}] ${ev.notebook}/${ev.section} (retag only)\n`);
      if (ev.type === "done" && ev.status === "ok") process.stdout.write(`    [ok] ${ev.section} (${ev.pages} pages)\n`);
      if (ev.type === "done" && ev.status === "failed") process.stdout.write(`    [failed] ${ev.section}\n`);
      if (ev.type === "complete") process.stdout.write(`Sync complete. ${ev.downloaded} downloaded, ${ev.retagged} retagged, ${ev.skipped} up-to-date.\n`);
    };
    await syncFn(emit);
    return;
  }

  const queue: SyncEvent[] = [];
  const state = { resolve: null as ((v: void) => void) | null, done: false };

  const emit = (ev: SyncEvent) => {
    queue.push(ev);
    state.resolve?.(undefined);
    state.resolve = null;
  };

  async function* makeStream(): AsyncIterable<SyncEvent> {
    while (!state.done || queue.length > 0) {
      while (queue.length > 0) yield queue.shift()!;
      if (!state.done) await new Promise<void>((r) => { state.resolve = r; });
    }
  }

  const { render } = await import("ink");
  const stream = makeStream();
  const { waitUntilExit } = render(<SyncUI events={stream} />);

  await syncFn(emit);
  state.done = true;
  state.resolve?.(undefined);
  await waitUntilExit();
}
