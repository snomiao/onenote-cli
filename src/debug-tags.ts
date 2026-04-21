/**
 * Debug tool: scan .one binary files and dump all property IDs found near NoteTags property
 * to help identify shape IDs for #question, #star, etc.
 *
 * Usage:
 *   bun run src/debug-tags.ts [section.one] [--page <title-fragment>]
 *   bun run src/debug-tags.ts --all
 *
 * Output: for each NoteTag occurrence, lists all 4-byte prid candidates in ±300 bytes
 */

import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import { dirname } from "node:path";

const PKG_ROOT = dirname(import.meta.dir);
const CACHE_ROOT = join(PKG_ROOT, ".onenote", "cache");

// MS-ONESTORE property markers
const NOTE_TAGS_PROP = Buffer.from([0x89, 0x34, 0x00, 0x40]); // 0x40003489: NoteTags property container
const ACTION_ITEM_STATUS = Buffer.from([0x70, 0x34, 0x00, 0x10]); // 0x10003470: ActionItemStatus
const SIZE_MARKER = Buffer.from([0x10, 0x00, 0x00, 0x00]); // page GUID anchor marker

// Known property IDs for annotation
const KNOWN_PRIDS: Record<number, string> = {
  0x10003470: "ActionItemStatus",
  0x40003489: "NoteTags",
  0x1c003481: "NoteTagHighlightColor?",
  0x1c003483: "NoteTagTextColor?",
  0x0c003482: "NoteTagShape?",
  0x14003484: "NoteTagCreated?",
  0x14003485: "NoteTagModified?",
  0x1c003486: "NoteTagFontColor?",
  0x2000346f: "ActionItemSchemaVersion?",
  0x10003472: "ActionItemType?",
  0x14003471: "ActionItemCompletionTime?",
  0x10346f: "SchemaRevision?",
};

function hex32(n: number) {
  return "0x" + n.toString(16).padStart(8, "0");
}

function scanFile(buf: Buffer, label: string, pageFilter?: string) {
  const pageAnchors: { offset: number; guidHex: string }[] = [];

  // Extract page anchors (SIZE_MARKER + 16-byte GUID)
  let p = 0;
  while (p < buf.length - 20) {
    const m = buf.indexOf(SIZE_MARKER, p);
    if (m < 0) break;
    const guidHex = buf.slice(m + 4, m + 20).toString("hex");
    // Only accept if it looks like a real GUID (not all zeros)
    if (!/^0{32}$/.test(guidHex)) {
      pageAnchors.push({ offset: m, guidHex });
    }
    p = m + 1;
  }

  // Find page name anchors: look for UTF-16LE strings near page GUID markers
  // (simplified: just use offsets)
  const getPageGuidAt = (offset: number): string => {
    for (let i = pageAnchors.length - 1; i >= 0; i--) {
      if (pageAnchors[i].offset <= offset) return pageAnchors[i].guidHex.slice(0, 8);
    }
    return "????????";
  };

  // Scan for all NOTE_TAGS_PROP occurrences
  const tagOccurrences: { offset: number; props: Map<number, number[]> }[] = [];

  let pos = 0;
  while (true) {
    const idx = buf.indexOf(NOTE_TAGS_PROP, pos);
    if (idx < 0) break;

    const windowStart = Math.max(0, idx - 16);
    const windowEnd = Math.min(buf.length, idx + 400);
    const window = buf.slice(windowStart, windowEnd);

    // Extract all 4-byte values in window (aligned to 4)
    const props = new Map<number, number[]>();
    for (let i = 0; i + 4 <= window.length; i += 4) {
      const v = window.readUInt32LE(i);
      // Filter: top byte must be a known type prefix (0x10, 0x14, 0x18, 0x1c, 0x20, 0x24, 0x40, 0x0c)
      const top = (v >>> 24) & 0xff;
      if ([0x10, 0x14, 0x18, 0x1c, 0x20, 0x24, 0x40, 0x0c, 0x08].includes(top)) {
        const offs = windowStart + i;
        if (!props.has(v)) props.set(v, []);
        props.get(v)!.push(offs);
      }
    }

    tagOccurrences.push({ offset: idx, props });
    pos = idx + 1;
  }

  if (tagOccurrences.length === 0) return;

  console.log(`\n=== ${label} ===`);
  console.log(`  ${tagOccurrences.length} NoteTags occurrences found`);

  // Aggregate: count how many occurrences each prop appears in
  const propFreq = new Map<number, number>();
  for (const occ of tagOccurrences) {
    for (const [v] of occ.props) {
      propFreq.set(v, (propFreq.get(v) ?? 0) + 1);
    }
  }

  // Print sorted by frequency (most common first)
  console.log("\n  Property ID frequencies (top byte = type prefix, next 3 = prop#):");
  const sorted = [...propFreq.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30);

  for (const [v, count] of sorted) {
    const known = KNOWN_PRIDS[v] ?? "";
    const freq = `(${count}/${tagOccurrences.length} occurrences)`;
    console.log(`    ${hex32(v)}  ${freq}  ${known}`);
  }

  // For each occurrence, print AIS status and all prid values
  console.log("\n  Per-occurrence detail:");
  for (let i = 0; i < tagOccurrences.length; i++) {
    const occ = tagOccurrences[i];
    const pageHex = getPageGuidAt(occ.offset);

    // Find AIS in vicinity
    const vicinity = buf.slice(occ.offset, Math.min(buf.length, occ.offset + 300));
    const aisRel = vicinity.indexOf(ACTION_ITEM_STATUS);
    let statusStr = "no-AIS";
    if (aisRel >= 0) {
      const aisAbs = occ.offset + aisRel;
      if (aisAbs + 14 <= buf.length) {
        const status = buf.readUInt16LE(aisAbs + 12);
        statusStr = status === 0 ? "☐ unchecked" : status === 1 ? "☑ checked" : `status=${status}`;
      }
    }

    // Collect ALL 4-byte aligned values near this occurrence with following value
    const rawStart = Math.max(0, occ.offset - 4);
    const raw = buf.slice(rawStart, Math.min(buf.length, occ.offset + 300));
    const allVals: string[] = [];
    for (let j = 0; j + 8 <= raw.length; j += 4) {
      const v = raw.readUInt32LE(j);
      const next = raw.readUInt32LE(j + 4);
      const top = (v >>> 24) & 0xff;
      if ([0x10, 0x14, 0x18, 0x1c, 0x20, 0x24, 0x40, 0x0c, 0x08].includes(top)) {
        const known = KNOWN_PRIDS[v] ? `[${KNOWN_PRIDS[v]}]` : "";
        allVals.push(`${hex32(v)}${known}=${hex32(next)}`);
      }
    }

    console.log(`    [${i + 1}] @${occ.offset} page~${pageHex}  ${statusStr}`);
    console.log(`         prids: ${allVals.join("  ")}`);
  }
}

async function findOneFiles(dir: string): Promise<string[]> {
  const files: string[] = [];
  try {
    const entries = await readdir(dir, { withFileTypes: true });
    for (const e of entries) {
      const full = join(dir, e.name);
      if (e.isDirectory()) {
        files.push(...(await findOneFiles(full)));
      } else if (e.name.endsWith(".one")) {
        files.push(full);
      }
    }
  } catch {}
  return files;
}

async function main() {
  const args = process.argv.slice(2);
  const allMode = args.includes("--all");
  const pageIdx = args.indexOf("--page");
  const pageFilter = pageIdx >= 0 ? args[pageIdx + 1] : undefined;
  const fileArgs = args.filter((a) => !a.startsWith("--") && a !== (pageFilter ?? "___"));

  let files: string[];
  if (fileArgs.length > 0) {
    files = fileArgs;
  } else if (allMode) {
    files = await findOneFiles(CACHE_ROOT);
  } else {
    // Default: only files that have NOTE_TAGS_PROP hits
    files = await findOneFiles(CACHE_ROOT);
  }

  console.log(`Scanning ${files.length} .one file(s) for NoteTags property patterns...`);

  let totalOcc = 0;
  for (const f of files) {
    const buf = await readFile(f);
    const label = f.replace(CACHE_ROOT + "/", "");
    // Quick check: has any NoteTags at all?
    if (buf.indexOf(NOTE_TAGS_PROP) < 0) continue;
    totalOcc++;
    scanFile(buf, label, pageFilter);
  }
  console.log(`\nTotal files with NoteTags: ${totalOcc}`);
}

main().catch(console.error);
