function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatInline(markdown: string): string {
  const codeSpans: string[] = [];
  let text = markdown.replace(/`([^`]+)`/g, (_match, code: string) => {
    const token = `\u0000CODE${codeSpans.length}\u0000`;
    codeSpans.push(`<code>${escapeHtml(code)}</code>`);
    return token;
  });

  text = escapeHtml(text);
  text = text
    .replace(
      /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g,
      '<a href="$2">$1</a>'
    )
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/__([^_]+)__/g, "<strong>$1</strong>")
    .replace(/(^|[^\w])\*([^*\n]+)\*(?!\w)/g, "$1<em>$2</em>")
    .replace(/(^|[^\w])_([^_\n]+)_(?!\w)/g, "$1<em>$2</em>");

  return text.replace(/\u0000CODE(\d+)\u0000/g, (_match, index: string) => {
    return codeSpans[Number(index)] ?? "";
  });
}

function flushParagraph(lines: string[], out: string[]) {
  if (lines.length === 0) return;
  out.push(`<p>${lines.map((line) => formatInline(line)).join("<br />")}</p>`);
  lines.length = 0;
}

function closeList(openList: "ul" | "ol" | null, out: string[]): null {
  if (!openList) return null;
  out.push(`</${openList}>`);
  return null;
}

export function markdownToHtml(markdown: string): string {
  const normalized = markdown.replace(/\r\n/g, "\n").trim();
  if (!normalized) return "<p></p>";

  const out: string[] = [];
  const lines = normalized.split("\n");
  const paragraph: string[] = [];
  let openList: "ul" | "ol" | null = null;

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const trimmed = line.trim();

    if (!trimmed) {
      flushParagraph(paragraph, out);
      openList = closeList(openList, out);
      continue;
    }

    const fenceMatch = trimmed.match(/^```([\w-]+)?$/);
    if (fenceMatch) {
      flushParagraph(paragraph, out);
      openList = closeList(openList, out);
      const codeLines: string[] = [];
      let j = i + 1;
      while (j < lines.length && !lines[j].trim().startsWith("```")) {
        codeLines.push(lines[j]);
        j += 1;
      }
      const language = fenceMatch[1] ? ` class="language-${escapeHtml(fenceMatch[1])}"` : "";
      out.push(`<pre><code${language}>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
      i = j;
      continue;
    }

    const headingMatch = trimmed.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      flushParagraph(paragraph, out);
      openList = closeList(openList, out);
      const level = headingMatch[1].length;
      out.push(`<h${level}>${formatInline(headingMatch[2].trim())}</h${level}>`);
      continue;
    }

    if (/^([-*_])(?:\s*\1){2,}$/.test(trimmed)) {
      flushParagraph(paragraph, out);
      openList = closeList(openList, out);
      out.push("<hr />");
      continue;
    }

    const bulletMatch = trimmed.match(/^[-*+]\s+(.+)$/);
    if (bulletMatch) {
      flushParagraph(paragraph, out);
      if (openList !== "ul") {
        openList = closeList(openList, out);
        out.push("<ul>");
        openList = "ul";
      }
      out.push(`<li>${formatInline(bulletMatch[1])}</li>`);
      continue;
    }

    const orderedMatch = trimmed.match(/^\d+\.\s+(.+)$/);
    if (orderedMatch) {
      flushParagraph(paragraph, out);
      if (openList !== "ol") {
        openList = closeList(openList, out);
        out.push("<ol>");
        openList = "ol";
      }
      out.push(`<li>${formatInline(orderedMatch[1])}</li>`);
      continue;
    }

    if (trimmed.startsWith(">")) {
      flushParagraph(paragraph, out);
      openList = closeList(openList, out);
      out.push(`<blockquote><p>${formatInline(trimmed.replace(/^>\s?/, ""))}</p></blockquote>`);
      continue;
    }

    paragraph.push(trimmed);
  }

  flushParagraph(paragraph, out);
  closeList(openList, out);
  return out.join("\n");
}
