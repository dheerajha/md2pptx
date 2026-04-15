/**
 * parser.js — Markdown parser
 * Extracts sections, metadata, tables, code blocks, lists, and numerical data.
 */

// Custom parser below — no external markdown library needed

function parseMarkdown(raw) {
  const lines = raw.split("\n");
  const sections = [];
  let current = null;
  let title = null;
  let subtitle = null;
  let inCode = false;
  let codeBlock = [];
  let codeLang = "";

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Code blocks
    if (line.startsWith("```")) {
      if (!inCode) {
        inCode = true;
        codeLang = line.slice(3).trim();
        codeBlock = [];
      } else {
        inCode = false;
        if (current) {
          current.blocks.push({ type: "code", lang: codeLang, content: codeBlock.join("\n") });
        }
        codeBlock = [];
        codeLang = "";
      }
      continue;
    }
    if (inCode) { codeBlock.push(line); continue; }

    // H1 — document title
    if (/^# /.test(line)) {
      title = line.replace(/^# /, "").trim();
      continue;
    }

    // H2 — major sections
    if (/^## /.test(line)) {
      if (current) sections.push(current);
      current = { heading: line.replace(/^## /, "").trim(), level: 2, blocks: [] };
      continue;
    }

    // H3 — sub-sections
    if (/^### /.test(line)) {
      if (!current) {
        current = { heading: line.replace(/^### /, "").trim(), level: 3, blocks: [] };
      } else {
        current.blocks.push({ type: "subheading", content: line.replace(/^### /, "").trim() });
      }
      continue;
    }

    // Tables
    if (/^\|/.test(line)) {
      const tableLines = [];
      let j = i;
      while (j < lines.length && /^\|/.test(lines[j])) {
        tableLines.push(lines[j]);
        j++;
      }
      i = j - 1;
      const table = parseTable(tableLines);
      if (table) {
        if (!current) current = { heading: "Data", level: 2, blocks: [] };
        current.blocks.push({ type: "table", ...table });
      }
      continue;
    }

    // Bullet / numbered lists
    if (/^[-*+] /.test(line) || /^\d+\. /.test(line)) {
      const items = [];
      let j = i;
      while (j < lines.length && (/^[-*+] /.test(lines[j]) || /^\d+\. /.test(lines[j]) || /^  [-*+] /.test(lines[j]))) {
        const indented = /^  /.test(lines[j]);
        const text = lines[j].replace(/^(\s*[-*+\d.]+\s*)/, "").trim();
        items.push({ text, indent: indented ? 1 : 0 });
        j++;
      }
      i = j - 1;
      if (!current) current = { heading: "", level: 2, blocks: [] };
      current.blocks.push({ type: "list", items });
      continue;
    }

    // Blockquote
    if (/^> /.test(line)) {
      const text = line.replace(/^> /, "").trim();
      if (current) current.blocks.push({ type: "quote", content: text });
      continue;
    }

    // Horizontal rule → ignore
    if (/^---+$/.test(line.trim()) || /^\*\*\*+$/.test(line.trim())) continue;

    // Paragraph
    const trimmed = line.trim();
    if (trimmed) {
      if (!title && sections.length === 0 && !current) {
        subtitle = trimmed;
        continue;
      }
      if (!current) current = { heading: "", level: 2, blocks: [] };
      // Check for inline bold as potential key stat
      const statMatch = trimmed.match(/\*\*([^*]+)\*\*/);
      const numMatch = trimmed.match(/\b(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s*(%|x|X|billion|million|trillion|thousand|k\b|M\b|B\b)?/gi);

      const lastBlock = current.blocks[current.blocks.length - 1];
      if (lastBlock && lastBlock.type === "paragraph") {
        lastBlock.content += " " + trimmed;
      } else {
        current.blocks.push({ type: "paragraph", content: trimmed, hasStat: !!numMatch, stat: numMatch ? numMatch[0] : null });
      }
    }
  }

  if (current) sections.push(current);

  return { title: title || "Presentation", subtitle, sections, raw };
}

function parseTable(lines) {
  const dataLines = lines.filter(l => !/^[|\s-:]+$/.test(l));
  if (dataLines.length < 2) return null;

  const parseRow = (l) => l.split("|").map(c => c.trim()).filter((c, i, a) => i > 0 && i < a.length - 1);

  const headers = parseRow(dataLines[0]);
  const rows = dataLines.slice(1).map(parseRow);

  // Detect numeric columns; treat first col as label if it looks like years or categories
  const numericCols = headers.map((h, ci) => {
    if (ci === 0) {
      // If first col is 4-digit years or text labels, treat as label not numeric
      const isYear = rows.every(r => r[0] && /^\d{4}$/.test(r[0].trim()));
      if (isYear) return false;
    }
    const isNumeric = rows.every(r => r[ci] && /^[-+]?[\d,.\s%$€£+KMB]+$/i.test(r[ci].trim()) && /\d/.test(r[ci]));
    if (!isNumeric) return false;
    // Exclude all-negative columns (e.g. reduction targets like -48%) — not good for charting
    const allNeg = rows.every(r => parseFloat((r[ci] || "0").replace(/[^0-9.-]/g, "")) < 0);
    return !allNeg;
  });

  return { headers, rows, numericCols };
}

module.exports = { parseMarkdown };
