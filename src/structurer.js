/**
 * structurer.js — Story Builder
 * Converts raw parsed sections into a slide-ready story with 10-15 slides.
 * Decisions: slide count, chart detection, infographic candidates, data slides.
 */

function buildStory(parsed) {
  const { title, subtitle, sections } = parsed;
  const slides = [];

  // 1. Title slide
  slides.push({
    type: "title",
    title,
    subtitle: subtitle || extractSubtitle(sections),
  });

  // 2. Agenda — if 4+ major sections
  const majorSections = sections.filter(s => s.level === 2 && s.heading);
  if (majorSections.length >= 4) {
    slides.push({
      type: "agenda",
      title: "Agenda",
      items: majorSections.map(s => s.heading),
    });
  }

  // 3. Executive Summary — first meaningful paragraphs
  const execSummary = buildExecSummary(sections);
  if (execSummary) slides.push(execSummary);

  // 4. Section content slides
  const contentSlides = buildContentSlides(sections);
  slides.push(...contentSlides);

  // 5. Chart / data slides extracted
  const dataSlides = extractDataSlides(sections);
  slides.push(...dataSlides);

  // 6. Conclusion / Key Takeaways
  const conclusion = buildConclusion(sections, title);
  slides.push(conclusion);

  // Enforce 10-15 slide limit
  const trimmed = trimSlides(slides);

  return { title, slides: trimmed, palette: choosePalette(title) };
}

function extractSubtitle(sections) {
  for (const s of sections) {
    for (const b of s.blocks) {
      if (b.type === "paragraph" && b.content.length > 20 && b.content.length < 180) {
        return b.content.replace(/\*\*/g, "");
      }
    }
  }
  return "";
}

function cleanMd(str) {
  return str.replace(/\*\*([^*]+)\*\*/g, "$1").replace(/\*([^*]+)\*/g, "$1").replace(/__([^_]+)__/g, "$1").replace(/_([^_]+)_/g, "$1");
}

function buildExecSummary(sections) {
  const points = [];
  for (const s of sections) {
    for (const b of s.blocks) {
      if (b.type === "paragraph" && b.content.length > 30) {
        const clean = cleanMd(b.content);
        points.push(clean);
        if (points.length >= 4) break;
      }
      if (b.type === "list" && b.items.length > 0) {
        b.items.slice(0, 3).forEach(item => points.push(item.text));
        if (points.length >= 4) break;
      }
    }
    if (points.length >= 4) break;
  }
  if (points.length === 0) return null;
  return { type: "executive_summary", title: "Executive Summary", points: points.slice(0, 5) };
}

function buildContentSlides(sections) {
  const slides = [];

  for (const section of sections) {
    if (!section.heading) continue;

    // Check if section has a table (will become data slide)
    const tables = section.blocks.filter(b => b.type === "table");
    const lists = section.blocks.filter(b => b.type === "list");
    const paragraphs = section.blocks.filter(b => b.type === "paragraph");
    const subheadings = section.blocks.filter(b => b.type === "subheading");
    const quotes = section.blocks.filter(b => b.type === "quote");

    // Large list → icon-card layout
    if (lists.length > 0 && lists[0].items.length >= 3 && paragraphs.length <= 1) {
      const items = lists[0].items.slice(0, 6);
      if (items.length >= 4) {
        slides.push({
          type: "icon_cards",
          title: section.heading,
          intro: paragraphs[0]?.content,
          items: items.map(i => ({ text: cleanMd(i.text) })),
        });
        continue;
      }
    }

    // Section with subheadings → two-column or steps
    if (subheadings.length >= 2) {
      const items = subheadings.map(sh => {
        const idx = section.blocks.indexOf(sh);
        const desc = section.blocks.slice(idx + 1).find(b => b.type === "paragraph" || b.type === "list");
        return {
          heading: cleanMd(sh.content),
          desc: desc?.type === "paragraph" ? cleanMd(desc.content).substring(0, 150) :
                desc?.type === "list" ? desc.items.slice(0, 2).map(i => cleanMd(i.text)).join("; ") : ""
        };
      });
      slides.push({
        type: "two_column_cards",
        title: section.heading,
        items: items.slice(0, 4),
      });
      continue;
    }

    // Quote callout
    if (quotes.length > 0) {
      slides.push({
        type: "quote_callout",
        title: section.heading,
        quote: quotes[0].content,
        body: paragraphs[0]?.content,
      });
      continue;
    }

    // Standard content slide
    const bodyText = paragraphs.slice(0, 3).map(p => p.content.replace(/\*\*([^*]+)\*\*/g, "$1")).join("\n\n");
    const bullets = lists.length > 0 ? lists[0].items.slice(0, 5) : [];

    slides.push({
      type: "content",
      title: section.heading,
      body: bodyText,
      bullets,
      stat: findStat(section),
    });
  }

  return slides;
}

function findStat(section) {
  for (const b of section.blocks) {
    if (b.type === "paragraph" && b.hasStat && b.stat) {
      return b.stat;
    }
  }
  return null;
}

function extractDataSlides(sections) {
  const slides = [];

  for (const section of sections) {
    for (const block of section.blocks) {
      if (block.type === "table") {
        const hasNumeric = block.numericCols.some(Boolean);
        const slideType = hasNumeric ? "chart_table" : "data_table";
        slides.push({
          type: slideType,
          title: section.heading || "Data Overview",
          table: block,
          chartType: inferChartType(block),
        });
      }
    }
  }

  return slides.slice(0, 3); // Cap at 3 data slides
}

function inferChartType(table) {
  const numCols = table.numericCols.filter(Boolean).length;
  const rowCount = table.rows.length;
  if (rowCount <= 2 && numCols >= 3) return "line";
  if (numCols === 1 && rowCount <= 6) return "bar";
  if (numCols >= 2 && rowCount <= 8) return "bar";
  return "bar";
}

function buildConclusion(sections, title) {
  // Find conclusion-like section first
  const concSection = sections.find(s =>
    /conclusion|takeaway|summary|next step|recommendation|outlook/i.test(s.heading)
  );

  const points = [];
  if (concSection) {
    for (const b of concSection.blocks) {
      if (b.type === "list") points.push(...b.items.slice(0, 5).map(i => i.text));
      if (b.type === "paragraph") points.push(cleanMd(b.content).substring(0, 180));
      if (points.length >= 5) break;
    }
  }

  // Fallback: collect key points from across doc
  if (points.length < 3) {
    for (const s of sections) {
      const list = s.blocks.find(b => b.type === "list");
      if (list) {
        points.push(...list.items.slice(0, 2).map(i => i.text));
        if (points.length >= 5) break;
      }
    }
  }

  return {
    type: "conclusion",
    title: "Key Takeaways",
    points: points.slice(0, 5),
  };
}

function trimSlides(slides) {
  // Must have: title, exec summary, conclusion = 3 fixed
  // 10-15 total
  if (slides.length <= 15) return slides;

  // Keep: title (0), agenda (1 if exists), exec_summary, conclusion
  const fixed = slides.filter(s => ["title", "conclusion"].includes(s.type));
  const agenda = slides.filter(s => s.type === "agenda");
  const execSum = slides.filter(s => s.type === "executive_summary");
  const data = slides.filter(s => ["chart_table", "data_table"].includes(s.type));
  const content = slides.filter(s => !["title", "agenda", "executive_summary", "conclusion", "chart_table", "data_table"].includes(s.type));

  const maxContent = 15 - fixed.length - agenda.length - execSum.length - Math.min(data.length, 2);
  const result = [
    ...slides.filter(s => s.type === "title"),
    ...agenda,
    ...execSum,
    ...content.slice(0, maxContent),
    ...data.slice(0, 2),
    ...slides.filter(s => s.type === "conclusion"),
  ];

  return result.slice(0, 15);
}

function choosePalette(title) {
  const t = title.toLowerCase();
  if (/tech|ai|data|digital|software|cloud/.test(t)) return "midnight";
  if (/finance|revenue|invest|market|fund/.test(t)) return "ocean";
  if (/health|medical|pharma|care/.test(t)) return "teal";
  if (/sustain|green|environment|climate/.test(t)) return "forest";
  if (/brand|market|growth|sales/.test(t)) return "coral";
  return "charcoal";
}

module.exports = { buildStory };
