const fs = require("fs");
const path = require("path");
const docx = require("docx");
const PptxGenJS = require("pptxgenjs");
const root = path.resolve(__dirname, "..", "..");
const plan = JSON.parse(fs.readFileSync(path.join(root, "build", "plan.json")));
const selected = new Set(process.argv.slice(2));
const C = { dark: "18242D", copper: "B5683A", cream: "F3F0E9", ink: "26343C", muted: "53616A", white: "FFFFFF" };

async function word(item, draft, output) {
  const { Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, HeadingLevel, BorderStyle } = docx;
  const children = [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(draft.title)] }),
    new Paragraph({ spacing: { after: 220 }, children: [new TextRun({
      text: draft.publication_date + (item.audience === "public" ? "" : " | Internal"),
      color: C.muted, size: 20
    })] })
  ];
  for (const section of draft.sections) {
    if (item.type !== "press_release" || ["company_boilerplate", "media_contact"].includes(section.role)) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true,
        children: [new TextRun(section.heading)] }));
    }
    for (const text of section.text.split(/\n\n|\n/)) {
      children.push(new Paragraph({ keepLines: true, spacing: { after: 130, line: 290 },
        children: [new TextRun(text)] }));
    }
  }
  const document = new Document({
    creator: "AstroDig LLC", title: draft.title, description: "AstroDig LLC business document",
    styles: {
      default: { document: { run: { font: "Arial", size: 23, color: C.ink } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", run: { font: "Arial", size: 36, bold: true, color: "000000" },
          paragraph: { spacing: { after: 220 }, keepNext: true } },
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Arial", size: 27, bold: true, color: "000000" },
          paragraph: { outlineLevel: 0, spacing: { before: 210, after: 140 }, keepNext: true } }
      ]
    },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 },
        margin: { top: 1200, bottom: 1200, left: 1440, right: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 5, color: C.copper, space: 5 } },
        children: [new TextRun({ text: "ASTRODIG LLC", font: "Arial", size: 18, bold: true, color: C.dark })]
      })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ children: [
        new TextRun({ text: item.audience === "public" ? "AstroDig LLC | " : "AstroDig LLC | Internal | ", size: 18, color: C.muted }),
        new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.muted })
      ] })] }) },
      children
    }]
  });
  fs.writeFileSync(output, await Packer.toBuffer(document));
}

function box(slide, text, x, y, w, h, options = {}) {
  slide.addText(text, { x, y, w, h, fontFace: "Calibri", fontSize: 18,
    color: C.ink, margin: 0, valign: "top", breakLine: false, ...options });
}

async function presentation(item, draft, output) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "AstroDig LLC";
  pptx.subject = draft.title;
  pptx.title = draft.title;
  pptx.company = "AstroDig LLC";
  pptx.lang = "en-US";
  pptx.theme = { headFontFace: "Trebuchet MS", bodyFontFace: "Calibri", lang: "en-US" };
  for (const [n, data] of draft.slides.entries()) {
    const slide = pptx.addSlide();
    const dark = ["cover", "closing"].includes(data.layout);
    slide.background = { color: dark ? C.dark : C.cream };
    box(slide, "ASTRODIG LLC", 0.65, 0.38, 7, 0.25,
      { fontSize: 12, bold: true, color: dark ? "DEB895" : C.copper, charSpacing: 2 });
    box(slide, data.layout === "cover" ? draft.title : data.heading, 0.65,
      data.layout === "cover" ? 1.12 : 0.90, 12, data.layout === "cover" ? 2.1 : 1.1,
      { fontFace: "Trebuchet MS", fontSize: data.layout === "cover" ? 30 : 32, bold: true,
        color: dark ? C.white : C.dark });
    box(slide, `${draft.publication_date} | ${item.audience === "board" ? "Board confidential" : "Internal"} | ${n + 1}`,
      0.65, 6.95, 12, 0.2, { fontSize: 10, color: dark ? "BCC6CD" : C.muted });
    if (data.layout === "cover") {
      box(slide, data.body.join("\n\n"), 0.7, 3.45, 8.2, 2.0,
        { fontSize: 24, color: C.white });
      slide.addShape(pptx.ShapeType.ellipse, { x: 10.15, y: 3.5, w: 2.1, h: 2.1,
        line: { color: C.copper, width: 4 }, fill: { color: C.dark } });
      box(slide, data.badge || "DEVELOPMENT", 10.30, 4.24, 1.8, 0.7,
        { align: "center", fontSize: 14, bold: true, color: "DEB895" });
    } else if (data.layout === "chart") {
      slide.addChart(pptx.ChartType.bar, [{ name: "USD", labels: data.chart.labels, values: data.chart.values }], {
        x: 0.85, y: 2.08, w: 8.5, h: 4.2, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11,
        showTitle: false, showLegend: false, showValue: true, dataLabelFormatCode: "#,##0;(#,##0)",
        dataLabelPosition: "outEnd", dataLabelColor: C.dark, dataLabelBkgrdColor: C.cream,
        chartColors: [C.copper], showBorder: false, showCatName: false, showCatNameInLegend: false,
        showMarker: false, showLine: false, catAxisLabelColor: C.ink, valAxisLabelColor: C.muted,
        valAxisNumFmt: "#,##0;(#,##0)", showCatNameInLabel: false,
        showSerName: false, showPercent: false, showValue: true,
        catAxisTitle: "Recognized USD", showCatName: false
      });
      box(slide, data.body.join("\n\n"), 9.8, 2.25, 2.6, 3.65, { fontSize: 19 });
    } else if (data.layout === "workflow") {
      const count = data.body.length;
      const w = (12 - (count - 1) * 0.32) / count;
      data.body.forEach((text, i) => {
        const x = 0.65 + i * (w + 0.32);
        slide.addShape(pptx.ShapeType.rect, { x, y: 2.3, w, h: 3.72,
          fill: { color: i % 2 ? "E6DED2" : "FFFFFF" }, line: { color: "D2C7B8", width: 1 } });
        box(slide, String(i + 1).padStart(2, "0"), x + 0.18, 2.50, w - 0.36, 0.6,
          { fontSize: 30, color: C.copper, bold: true });
        box(slide, text, x + 0.18, 3.35, w - 0.36, 2.3,
          { fontSize: count > 3 ? 17 : 19 });
      });
    } else if (data.layout === "grid") {
      data.body.forEach((text, i) => {
        const x = 0.65 + (i % 2) * 6.16;
        const y = 2.1 + Math.floor(i / 2) * 1.42;
        slide.addShape(pptx.ShapeType.rect, { x, y, w: 5.82, h: 1.14,
          fill: { color: "FFFFFF" }, line: { color: "D2C7B8", width: 1 } });
        box(slide, text, x + 0.22, y + 0.18, 5.38, 0.82, { fontSize: 17 });
      });
    } else if (data.layout === "capabilities") {
      const gap = 0.22;
      const h = Math.min(0.90, (4.4 - (data.body.length - 1) * gap) / data.body.length);
      data.body.forEach((text, i) => {
        const y = 2.1 + i * (h + gap);
        slide.addShape(pptx.ShapeType.rect, { x: 0.65, y, w: 0.62, h,
          fill: { color: C.copper }, line: { color: C.copper } });
        box(slide, String(i + 1), 0.69, y + 0.16, 0.54, 0.4,
          { fontSize: 20, color: C.white, bold: true, align: "center" });
        box(slide, text, 1.60, y + 0.05, 10.45, h - 0.04,
          { fontSize: data.body.length > 4 ? 17 : 18 });
      });
    } else if (["cards", "stats"].includes(data.layout)) {
      const w = (12 - (data.body.length - 1) * 0.38) / data.body.length;
      data.body.forEach((text, i) => {
        const x = 0.65 + i * (w + 0.38);
        slide.addShape(pptx.ShapeType.rect, { x, y: 2.25, w, h: 3.75,
          fill: { color: "FFFFFF" }, line: { color: "D2C7B8", width: 1 } });
        box(slide, String(i + 1).padStart(2, "0"), x + 0.22, 2.47, w - 0.44, 0.62,
          { fontSize: 32, color: C.copper, bold: true });
        box(slide, text, x + 0.22, 3.35, w - 0.44, 2.36,
          { fontSize: data.layout === "stats" ? 22 : 18 });
      });
    } else {
      slide.addShape(pptx.ShapeType.rect, { x: 0.65, y: 2.1, w: 0.14, h: 4.13,
        fill: { color: C.copper }, line: { color: C.copper } });
      box(slide, data.body.join("\n\n"), 1.12, 2.15, 10.85, 4.0,
        { fontSize: 23, color: dark ? C.white : C.ink });
    }
    slide.addNotes(data.body.join("\n\n"));
  }
  await pptx.writeFile({ fileName: output });
}

(async () => {
  for (const item of plan.documents) {
    if (selected.size && !selected.has(item.id)) continue;
    if (item.format === "xlsx") continue;
    const output = path.join(root, "documents", item.filename);
    if (fs.existsSync(output)) continue;
    const draft = JSON.parse(fs.readFileSync(path.join(root, "build", "drafts", `${item.id}.json`)));
    if (item.format === "docx") await word(item, draft, output);
    else await presentation(item, draft, output);
    console.log(item.id, item.format, item.filename);
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
