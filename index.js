#!/usr/bin/env node
/**
 * md2pptx — Markdown to PPTX Converter
 * Modular pipeline: parse → structure → enrich → render
 */

const fs = require("fs");
const path = require("path");
const { parseMarkdown } = require("./src/parser");
const { buildStory } = require("./src/structurer");
const { renderPresentation } = require("./src/renderer");

async function convert(inputPath, outputPath) {
  const stat = fs.statSync(inputPath);
  if (stat.size > 5 * 1024 * 1024) {
    throw new Error("Input file exceeds 5 MB limit.");
  }

  const raw = fs.readFileSync(inputPath, "utf8");
  console.log(`📄  Parsing: ${inputPath}`);
  const parsed = parseMarkdown(raw);

  console.log(`🧠  Structuring storyline (${parsed.sections.length} sections found)...`);
  const story = buildStory(parsed);

  const finalOutput = outputPath || inputPath.replace(/\.md$/, ".pptx");
  console.log(`🎨  Rendering presentation → ${finalOutput}`);
  await renderPresentation(story, finalOutput);

  console.log(`✅  Done! Output: ${finalOutput}`);
  return finalOutput;
}

// CLI usage
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.error("Usage: node index.js <input.md> [output.pptx]");
    process.exit(1);
  }
  convert(args[0], args[1]).catch((err) => {
    console.error("Error:", err.message);
    process.exit(1);
  });
}

module.exports = { convert };
