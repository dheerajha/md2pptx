/**
 * renderer.js — PptxGenJS Presentation Builder
 * Uses defineSlideMaster() for every layout — consistent branding across all slides.
 * Rich chart types: multi-series bar, line, doughnut with proper axis styling.
 */

const pptxgen = require("pptxgenjs");
const { PALETTES, FONTS } = require("./themes");

const W = 10, H = 5.625;

// ─── SLIDE MASTER DEFINITIONS ─────────────────────────────────────────────────
function defineMasters(pres, pal) {
  // 1. TITLE_COVER — first slide only
  pres.defineSlideMaster({
    title: "TITLE_COVER",
    background: { color: pal.dark },
    objects: [
      { rect: { x: 0, y: 0, w: 0.18, h: H, fill: { color: pal.accent } } },
      { oval: { x: 6.8, y: -0.8, w: 4.5, h: 4.5, fill: { color: pal.primary, transparency: 55 }, line: { color: pal.primary, transparency: 55 } } },
      { oval: { x: 8.2, y: 2.8, w: 2.2, h: 2.2, fill: { color: pal.accent, transparency: 70 }, line: { color: pal.accent, transparency: 70 } } },
      { rect: { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: pal.primary } } },
      {
        placeholder: {
          options: {
            name: "main_title", type: "title",
            x: 0.55, y: 1.5, w: 7.2, h: 1.6,
            fontSize: 40, fontFace: FONTS.heading, bold: true, color: "FFFFFF",
            align: "left", valign: "middle", margin: 0,
          },
        },
      },
      {
        placeholder: {
          options: {
            name: "sub_title", type: "body",
            x: 0.55, y: 3.2, w: 7.2, h: 0.9,
            fontSize: 16, fontFace: FONTS.body, bold: false, color: pal.secondary,
            align: "left", valign: "top", margin: 0,
          },
        },
      },
    ],
  });

  // 2. DARK — conclusion, exec summary, quote, dark content
  pres.defineSlideMaster({
    title: "DARK",
    background: { color: pal.dark },
    objects: [
      { rect: { x: 0, y: 0, w: 0.12, h: H, fill: { color: pal.accent } } },
      { rect: { x: 0, y: H - 0.18, w: W, h: 0.18, fill: { color: pal.primary } } },
    ],
    slideNumber: { x: 9.3, y: H - 0.17, w: 0.5, h: 0.16, fontSize: 8, color: pal.secondary, align: "right" },
  });

  // 3. LIGHT — content, agenda, icon cards, two-col
  pres.defineSlideMaster({
    title: "LIGHT",
    background: { color: pal.light },
    objects: [
      { rect: { x: 0, y: 0, w: W, h: 0.95, fill: { color: pal.primary } } },
      { rect: { x: 0, y: 0.95, w: W, h: 0.04, fill: { color: pal.accent } } },
      { rect: { x: 0, y: H - 0.18, w: W, h: 0.18, fill: { color: pal.primary } } },
      {
        placeholder: {
          options: {
            name: "hdr", type: "title",
            x: 0.4, y: 0.08, w: 9.2, h: 0.8,
            fontSize: 24, fontFace: FONTS.heading, bold: true, color: "FFFFFF",
            align: "left", valign: "middle", margin: 0,
          },
        },
      },
    ],
    slideNumber: { x: 9.3, y: H - 0.17, w: 0.5, h: 0.16, fontSize: 8, color: pal.secondary, align: "right" },
  });

  // 4. CHART — data/chart slides
  pres.defineSlideMaster({
    title: "CHART",
    background: { color: pal.light },
    objects: [
      { rect: { x: 0, y: 0, w: W, h: 0.95, fill: { color: pal.primary } } },
      { rect: { x: 0, y: 0.95, w: W, h: 0.04, fill: { color: pal.accent } } },
      { rect: { x: 0, y: H - 0.18, w: W, h: 0.18, fill: { color: pal.primary } } },
      {
        placeholder: {
          options: {
            name: "hdr", type: "title",
            x: 0.4, y: 0.08, w: 9.2, h: 0.8,
            fontSize: 24, fontFace: FONTS.heading, bold: true, color: "FFFFFF",
            align: "left", valign: "middle", margin: 0,
          },
        },
      },
    ],
    slideNumber: { x: 9.3, y: H - 0.17, w: 0.5, h: 0.16, fontSize: 8, color: pal.secondary, align: "right" },
  });
}

// ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────
async function renderPresentation(story, outputPath) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "md2pptx";
  pres.title = story.title;
  pres.subject = story.title;

  const pal = PALETTES[story.palette] || PALETTES.charcoal;

  // MUST define all masters before adding any slides
  defineMasters(pres, pal);

  for (const slide of story.slides) {
    await renderSlide(pres, slide, pal);
  }

  await pres.writeFile({ fileName: outputPath });
}

async function renderSlide(pres, data, pal) {
  switch (data.type) {
    case "title":             return renderTitle(pres, data, pal);
    case "agenda":            return renderAgenda(pres, data, pal);
    case "executive_summary": return renderExecSummary(pres, data, pal);
    case "content":           return renderContent(pres, data, pal);
    case "icon_cards":        return renderIconCards(pres, data, pal);
    case "two_column_cards":  return renderTwoColCards(pres, data, pal);
    case "quote_callout":     return renderQuoteCallout(pres, data, pal);
    case "chart_table":       return renderChartTable(pres, data, pal);
    case "data_table":        return renderDataTable(pres, data, pal);
    case "conclusion":        return renderConclusion(pres, data, pal);
    default:                  return renderContent(pres, data, pal);
  }
}

// ─── TITLE ────────────────────────────────────────────────────────────────────
function renderTitle(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "TITLE_COVER" });
  slide.addText(data.title, { placeholder: "main_title" });
  if (data.subtitle) {
    slide.addText(cleanMd(data.subtitle).substring(0, 160), { placeholder: "sub_title" });
  }
}

// ─── AGENDA ───────────────────────────────────────────────────────────────────
function renderAgenda(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "LIGHT" });
  slide.addText("Agenda", { placeholder: "hdr" });

  const items = data.items.slice(0, 8);
  const half = Math.ceil(items.length / 2);
  const colW = 4.25;
  const boxH = 0.58;
  const startY = 1.12;
  const gap = 0.15;

  items.forEach((item, i) => {
    const col = i < half ? 0 : 1;
    const row = i < half ? i : i - half;
    const x = 0.35 + col * (colW + 0.55);
    const y = startY + row * (boxH + gap);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: colW, h: boxH,
      fill: { color: "FFFFFF" },
      line: { color: pal.primary, width: 0.75 },
      shadow: { type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.07 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.42, h: boxH,
      fill: { color: pal.primary }, line: { color: pal.primary },
    });
    slide.addText(String(i + 1), {
      x, y, w: 0.42, h: boxH,
      fontSize: 14, fontFace: FONTS.heading, bold: true,
      color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    slide.addText(item, {
      x: x + 0.5, y, w: colW - 0.58, h: boxH,
      fontSize: 13, fontFace: FONTS.body, color: pal.text,
      valign: "middle", margin: 0,
    });
  });
}

// ─── EXECUTIVE SUMMARY ────────────────────────────────────────────────────────
function renderExecSummary(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "DARK" });

  slide.addText("Executive Summary", {
    x: 0.35, y: 0.2, w: 9.3, h: 0.72,
    fontSize: 28, fontFace: FONTS.heading, bold: true,
    color: "FFFFFF", valign: "middle", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.0, w: 9.3, h: 0.04,
    fill: { color: pal.accent }, line: { color: pal.accent },
  });

  const points = data.points.slice(0, 4);
  const cardH = (H - 1.25) / points.length - 0.1;

  points.forEach((point, i) => {
    const y = 1.1 + i * (cardH + 0.1);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: cardH,
      fill: { color: pal.primary, transparency: 35 },
      line: { color: pal.accent, width: 0.5 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 0.07, h: cardH,
      fill: { color: pal.accent }, line: { color: pal.accent },
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.56, y: y + cardH / 2 - 0.2, w: 0.4, h: 0.4,
      fill: { color: pal.accent }, line: { color: pal.accent },
    });
    slide.addText(String(i + 1), {
      x: 0.56, y: y + cardH / 2 - 0.2, w: 0.4, h: 0.4,
      fontSize: 12, fontFace: FONTS.heading, bold: true,
      color: pal.dark, align: "center", valign: "middle", margin: 0,
    });

    const maxChars = Math.floor(cardH * 270);
    let txt = point;
    if (txt.length > maxChars) txt = txt.substring(0, maxChars).replace(/\s\S+$/, "") + "…";

    slide.addText(txt, {
      x: 1.1, y: y + 0.06, w: 8.4, h: cardH - 0.12,
      fontSize: 13, fontFace: FONTS.body, color: "FFFFFF",
      valign: "middle", margin: 0, wrap: true,
    });
  });
}

// ─── CONTENT ──────────────────────────────────────────────────────────────────
function renderContent(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "LIGHT" });
  slide.addText(data.title, { placeholder: "hdr" });

  const hasStat = data.stat;
  const hasBullets = data.bullets && data.bullets.length > 0;
  const rawBody = cleanMd(data.body || "");
  const hasBody = rawBody.trim().length > 0;
  const contentW = hasStat ? 6.2 : 9.3;
  let startY = 1.1;

  if (hasStat) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.85, y: 1.05, w: 2.8, h: 2.6,
      fill: { color: pal.primary }, line: { color: pal.primary },
      shadow: { type: "outer", blur: 10, offset: 4, angle: 135, color: "000000", opacity: 0.18 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.85, y: 1.05, w: 2.8, h: 0.06,
      fill: { color: pal.accent }, line: { color: pal.accent },
    });
    const statVal = data.stat.length > 14 ? data.stat.substring(0, 14) : data.stat;
    slide.addText(statVal, {
      x: 6.85, y: 1.42, w: 2.8, h: 1.0,
      fontSize: 32, fontFace: FONTS.heading, bold: true,
      color: pal.accent, align: "center", margin: 0,
    });
    slide.addText("KEY METRIC", {
      x: 6.85, y: 2.52, w: 2.8, h: 0.35,
      fontSize: 10, fontFace: FONTS.body, bold: true,
      color: pal.secondary, align: "center", margin: 0, charSpacing: 2,
    });
  }

  let y = startY;
  if (hasBody && !hasBullets) {
    slide.addText(rawBody.substring(0, 700), {
      x: 0.4, y, w: contentW, h: H - y - 0.3,
      fontSize: 15, fontFace: FONTS.body, color: pal.text,
      valign: "top", margin: 0, wrap: true,
    });
  } else if (hasBullets) {
    if (hasBody) {
      slide.addText(rawBody.substring(0, 240), {
        x: 0.4, y, w: contentW, h: 1.0,
        fontSize: 13, fontFace: FONTS.body, color: pal.text,
        valign: "top", margin: 0, wrap: true,
      });
      y += 1.05;
    }
    const bulletItems = data.bullets.slice(0, 6).map((b, i, arr) => ({
      text: cleanMd(b.text).substring(0, 140),
      options: {
        bullet: { code: "25B8", color: pal.accent },
        breakLine: i < arr.length - 1,
        fontSize: 14, fontFace: FONTS.body, color: pal.text,
        paraSpaceAfter: 8, indentLevel: b.indent || 0,
      },
    }));
    slide.addText(bulletItems, {
      x: 0.4, y, w: contentW, h: H - y - 0.28,
      valign: "top", margin: 0,
    });
  }
}

// ─── ICON CARDS ───────────────────────────────────────────────────────────────
function renderIconCards(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "LIGHT" });
  slide.addText(data.title, { placeholder: "hdr" });

  const items = data.items.slice(0, 6);
  const cols = items.length <= 3 ? items.length : 3;
  const rows = Math.ceil(items.length / cols);
  const cardW = (W - 0.7) / cols - 0.2;
  const cardH = (H - 1.35) / rows - 0.18;
  const colors = pal.chartColors;

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = 0.35 + col * (cardW + 0.2);
    const y = 1.15 + row * (cardH + 0.18);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: "FFFFFF" },
      line: { color: colors[i % colors.length], width: 1.5 },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.09 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: 0.08,
      fill: { color: colors[i % colors.length] }, line: { color: colors[i % colors.length] },
    });
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.14, y: y + 0.14, w: 0.38, h: 0.38,
      fill: { color: colors[i % colors.length] }, line: { color: colors[i % colors.length] },
    });
    slide.addText(String(i + 1), {
      x: x + 0.14, y: y + 0.14, w: 0.38, h: 0.38,
      fontSize: 12, fontFace: FONTS.heading, bold: true,
      color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    slide.addText(cleanMd(item.text).substring(0, 110), {
      x: x + 0.1, y: y + 0.58, w: cardW - 0.2, h: cardH - 0.7,
      fontSize: 12, fontFace: FONTS.body, color: pal.text,
      valign: "top", margin: 0, wrap: true,
    });
  });
}

// ─── TWO COLUMN CARDS ─────────────────────────────────────────────────────────
function renderTwoColCards(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "LIGHT" });
  slide.addText(data.title, { placeholder: "hdr" });

  const items = data.items.slice(0, 4);
  const cardW = 4.5;
  const cardH = (H - 1.3) / Math.ceil(items.length / 2) - 0.2;
  const colors = pal.chartColors;

  items.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.35 + col * (cardW + 0.3);
    const y = 1.1 + row * (cardH + 0.2);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: "FFFFFF" },
      line: { color: "E5E7EB", width: 0.5 },
      shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.07 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.09, h: cardH,
      fill: { color: colors[i % colors.length] }, line: { color: colors[i % colors.length] },
    });
    slide.addText(item.heading, {
      x: x + 0.2, y: y + 0.12, w: cardW - 0.28, h: 0.42,
      fontSize: 15, fontFace: FONTS.heading, bold: true, color: pal.primary, margin: 0,
    });
    if (item.desc) {
      slide.addText(cleanMd(item.desc).substring(0, 180), {
        x: x + 0.2, y: y + 0.58, w: cardW - 0.3, h: cardH - 0.7,
        fontSize: 12, fontFace: FONTS.body, color: pal.text,
        valign: "top", margin: 0, wrap: true,
      });
    }
  });
}

// ─── QUOTE CALLOUT ────────────────────────────────────────────────────────────
function renderQuoteCallout(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "DARK" });

  slide.addText(data.title, {
    x: 0.35, y: 0.18, w: 9.3, h: 0.72,
    fontSize: 24, fontFace: FONTS.heading, bold: true, color: "FFFFFF",
    valign: "middle", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.08, w: 9.3, h: 2.8,
    fill: { color: pal.primary, transparency: 25 },
    line: { color: pal.accent, width: 1.2 },
  });
  slide.addText("\u201C", {
    x: 0.42, y: 0.92, w: 0.9, h: 0.9,
    fontSize: 72, fontFace: "Georgia", color: pal.accent, margin: 0,
  });
  slide.addText(cleanMd(data.quote).substring(0, 300), {
    x: 0.95, y: 1.22, w: 8.5, h: 2.45,
    fontSize: 17, fontFace: "Georgia", italic: true,
    color: "FFFFFF", valign: "middle", margin: 0, wrap: true,
  });
  if (data.body) {
    slide.addText(cleanMd(data.body).substring(0, 200), {
      x: 0.35, y: 4.05, w: 9.3, h: 0.9,
      fontSize: 13, fontFace: FONTS.body, color: pal.secondary,
      valign: "top", margin: 0, wrap: true,
    });
  }
}

// ─── CHART + TABLE ────────────────────────────────────────────────────────────
function renderChartTable(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "CHART" });
  slide.addText(data.title, { placeholder: "hdr" });

  const table = data.table;

  // All numeric columns (skip col 0 if it looks like labels/years)
  const numericColIdxs = table.headers
    .map((_, ci) => ci)
    .filter(ci => ci !== 0 && table.numericCols[ci]);
  const valueCols = numericColIdxs.length > 0 ? numericColIdxs : [1];

  const labels = !table.numericCols[0]
    ? table.rows.map(r => String(r[0] || ""))
    : table.headers.slice(1);

  // Multi-series data
  const chartData = valueCols.slice(0, 4).map(ci => ({
    name: table.headers[ci] || `Series ${ci}`,
    labels,
    values: table.rows.map(r => parseFloat((r[ci] || "0").replace(/[^0-9.-]/g, "")) || 0),
  }));

  const rowCount = table.rows.length;
  const seriesCount = chartData.length;

  let chosenType, chartOpts;
  if (seriesCount === 1 && rowCount <= 7) {
    // Doughnut for single-series small sets — most visually striking
    chosenType = pres.charts.DOUGHNUT;
    chartOpts = {
      chartColors: pal.chartColors,
      chartArea: { fill: { color: "FFFFFF" } },
      showLabel: true,
      showPercent: true,
      showLegend: true,
      legendPos: "b",
      legendFontSize: 10,
      legendColor: "374151",
      dataLabelFontSize: 10,
      dataLabelColor: "FFFFFF",
      holeSize: 50,
      showTitle: false,
    };
  } else if (seriesCount >= 2) {
    // Grouped bar for multi-series
    chosenType = pres.charts.BAR;
    chartOpts = {
      barDir: "col",
      barGrouping: "clustered",
      chartColors: pal.chartColors,
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: false },
      plotArea: { fill: { color: "FAFAFA" } },
      catAxisLabelColor: "64748B",
      valAxisLabelColor: "64748B",
      catAxisLineShow: false,
      valAxisLineShow: false,
      valGridLine: { color: "E2E8F0", style: "solid", size: 0.5 },
      catGridLine: { style: "none" },
      showValue: false,
      showLegend: true,
      legendPos: "b",
      legendFontSize: 10,
      legendColor: "374151",
      showTitle: false,
    };
  } else if (rowCount > 7) {
    // Line for time-series
    chosenType = pres.charts.LINE;
    chartOpts = {
      chartColors: pal.chartColors,
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: false },
      plotArea: { fill: { color: "FAFAFA" } },
      catAxisLabelColor: "64748B",
      valAxisLabelColor: "64748B",
      catAxisLineShow: false,
      valAxisLineShow: false,
      valGridLine: { color: "E2E8F0", style: "solid", size: 0.5 },
      catGridLine: { style: "none" },
      lineSize: 2.5,
      lineSmooth: true,
      showMarker: true,
      markerSize: 5,
      showValue: false,
      showLegend: true,
      legendPos: "b",
      legendFontSize: 10,
      showTitle: false,
    };
  } else {
    // Standard column bar
    chosenType = pres.charts.BAR;
    chartOpts = {
      barDir: "col",
      barGrouping: "clustered",
      chartColors: pal.chartColors,
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: false },
      plotArea: { fill: { color: "FAFAFA" } },
      catAxisLabelColor: "64748B",
      valAxisLabelColor: "64748B",
      catAxisLineShow: false,
      valAxisLineShow: false,
      valGridLine: { color: "E2E8F0", style: "solid", size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true,
      dataLabelFontSize: 9,
      dataLabelColor: "374151",
      dataLabelPosition: "outEnd",
      showLegend: false,
      showTitle: false,
    };
  }

  slide.addChart(chosenType, chartData, {
    x: 0.35, y: 1.05, w: 5.9, h: 4.2,
    ...chartOpts,
  });

  renderTablePanel(slide, pres, table, 6.55, 1.05, 3.1, pal);
}

// ─── DATA TABLE ONLY ──────────────────────────────────────────────────────────
function renderDataTable(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "CHART" });
  slide.addText(data.title, { placeholder: "hdr" });
  renderTablePanel(slide, pres, data.table, 0.4, 1.05, 9.25, pal);
}

// ─── TABLE PANEL (shared) ─────────────────────────────────────────────────────
function renderTablePanel(slide, pres, table, x, y, w, pal) {
  const maxRows = Math.min(table.rows.length, 9);
  const colCount = table.headers.length;
  const colW = w / colCount;

  const tableData = [
    table.headers.map(h => ({
      text: h,
      options: {
        fill: { color: pal.primary }, color: "FFFFFF",
        bold: true, fontSize: 10.5, fontFace: FONTS.heading,
        align: "center", valign: "middle",
      },
    })),
    ...table.rows.slice(0, maxRows).map((row, ri) =>
      row.slice(0, colCount).map((cell, ci) => ({
        text: String(cell),
        options: {
          fill: { color: ri % 2 === 0 ? "FFFFFF" : "F1F5F9" },
          color: pal.text, fontSize: 10, fontFace: FONTS.body,
          align: table.numericCols[ci] ? "right" : "left",
          valign: "middle",
        },
      }))
    ),
  ];

  slide.addTable(tableData, {
    x, y, w,
    colW: table.headers.map(() => colW),
    border: { pt: 0.5, color: "D1D5DB" },
    rowH: 0.32,
  });
}

// ─── CONCLUSION ───────────────────────────────────────────────────────────────
function renderConclusion(pres, data, pal) {
  const slide = pres.addSlide({ masterName: "DARK" });

  // Decorative circles
  slide.addShape(pres.shapes.OVAL, {
    x: 6.5, y: -1.2, w: 5.5, h: 5.5,
    fill: { color: pal.primary, transparency: 55 },
    line: { color: pal.primary, transparency: 55 },
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 8.5, y: 3.2, w: 2.5, h: 2.5,
    fill: { color: pal.accent, transparency: 70 },
    line: { color: pal.accent, transparency: 70 },
  });

  slide.addText(data.title, {
    x: 0.35, y: 0.2, w: 6.5, h: 0.75,
    fontSize: 30, fontFace: FONTS.heading, bold: true, color: "FFFFFF", margin: 0,
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.05, w: 2.8, h: 0.05,
    fill: { color: pal.accent }, line: { color: pal.accent },
  });

  const points = data.points.slice(0, 5);
  const itemH = (H - 1.3) / points.length - 0.1;

  points.forEach((point, i) => {
    const y = 1.18 + i * (itemH + 0.1);

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 6.5, h: itemH,
      fill: { color: pal.primary, transparency: i % 2 === 0 ? 55 : 70 },
      line: { color: "FFFFFF", transparency: 85 },
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.48, y: y + itemH / 2 - 0.2, w: 0.4, h: 0.4,
      fill: { color: pal.accent }, line: { color: pal.accent },
    });
    slide.addText(String(i + 1), {
      x: 0.48, y: y + itemH / 2 - 0.2, w: 0.4, h: 0.4,
      fontSize: 12, fontFace: FONTS.heading, bold: true,
      color: pal.dark, align: "center", valign: "middle", margin: 0,
    });

    let txt = point;
    if (txt.length > 160) {
      const sentEnd = txt.substring(0, 210).lastIndexOf(". ");
      txt = sentEnd > 80 ? txt.substring(0, sentEnd + 1) : txt.substring(0, 185).replace(/\s\S+$/, "") + "…";
    }
    slide.addText(txt, {
      x: 1.05, y, w: 5.65, h: itemH,
      fontSize: 13.5, fontFace: FONTS.body, color: "FFFFFF",
      valign: "middle", margin: 0, wrap: true,
    });
  });
}

// ─── SHARED UTIL ──────────────────────────────────────────────────────────────
function cleanMd(str) {
  return (str || "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/_([^_]+)_/g, "$1");
}

module.exports = { renderPresentation };
