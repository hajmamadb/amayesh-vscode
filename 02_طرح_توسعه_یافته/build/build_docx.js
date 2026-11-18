/*
 * ساخت فایل Word (DOCX) از فایل‌های Markdown فصل‌ها
 * اجرای دستور:  node build_docx.js
 * خروجی:        ..\سند_آمایش_استانی_طرح_توسعه_یافته.docx
 */
"use strict";

const fs = require("fs");
const path = require("path");

/* بارگذاری بسته docx: ترجیح به node_modules محلی؛ در نبود آن، مسیر پیشین */
let docx;
try {
  docx = require("docx");
} catch (e) {
  docx = require("D:/آمایش2 زد/01_طرح_توسعه_یافته/build/node_modules/docx/build/index.cjs");
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, HeadingLevel, PageBreak,
  Footer, PageNumber, ShadingType, VerticalAlign,
} = docx;

const FONT = "B Nazanin";
const FONT_MONO = "Consolas";
const BODY_SIZE = 24;   // 12pt
const TABLE_SIZE = 24;  // 12pt (هم‌اندازه متن برای نگارش رسمی)
const CODE_SIZE = 18;   // 9pt

const mdDir = path.join(__dirname, "..", "md");
const outPath = path.join(__dirname, "..", "سند_آمایش_استانی_طرح_توسعه_یافته.docx");

const faDigits = (s) => String(s).replace(/\d/g, (d) => "۰۱۲۳۴۵۶۷۸۹"[+d]);

/* ---------- تجزیه متن درون‌سطری: **بولد** و `کد` ---------- */
function parseRuns(text, base) {
  const runs = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const tok = m[0];
    if (tok.startsWith("**")) {
      runs.push(new TextRun({ text: tok.slice(2, -2), bold: true, ...base }));
    } else {
      runs.push(new TextRun({ text: tok.slice(1, -1), font: FONT_MONO, size: base.size || BODY_SIZE }));
    }
    last = m.index + tok.length;
  }
  if (last < text.length) runs.push(new TextRun({ text: text.slice(last), ...base }));
  if (!runs.length) runs.push(new TextRun({ text: "", ...base }));
  return runs;
}

/* ---------- تجزیه یک فایل Markdown به بلوک‌ها ---------- */
function parseFile(lines) {
  const blocks = [];
  let i = 0;
  while (i < lines.length) {
    const raw = lines[i];
    const line = raw.replace(/\r$/, "");
    const t = line.trim();

    if (t === "") { i++; continue; }

    // بلوک کد
    if (t.startsWith("```")) {
      const code = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith("```")) {
        code.push(lines[i].replace(/\r$/, ""));
        i++;
      }
      i++; // خط پایانی ```
      blocks.push({ type: "code", lines: code });
      continue;
    }

    // جدول
    if (t.startsWith("|")) {
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) {
        const cells = lines[i].trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
        if (!cells.every((c) => /^:?-{2,}:?$/.test(c))) rows.push(cells);
        i++;
      }
      blocks.push({ type: "table", rows });
      continue;
    }

    if (t.startsWith("### ")) { blocks.push({ type: "h3", text: t.slice(4) }); i++; continue; }
    if (t.startsWith("## "))  { blocks.push({ type: "h2", text: t.slice(3) }); i++; continue; }
    if (t.startsWith("# "))   { blocks.push({ type: "h1", text: t.slice(2) }); i++; continue; }
    if (/^(-{3,}|\*{3,})$/.test(t)) { blocks.push({ type: "hr" }); i++; continue; }
    if (t.startsWith("> "))   { blocks.push({ type: "quote", text: t.slice(2) }); i++; continue; }

    // فهرست شماره‌دار:  1. متن
    let mNum = t.match(/^(\d+)[.)]\s+(.*)$/);
    if (mNum) { blocks.push({ type: "num", n: mNum[1], text: mNum[2] }); i++; continue; }

    // گلوله
    if (t.startsWith("- ")) { blocks.push({ type: "bullet", text: t.slice(2) }); i++; continue; }

    blocks.push({ type: "p", text: t });
    i++;
  }
  return blocks;
}

/* ---------- تبدیل بلوک‌ها به عناصر DOCX ---------- */
const TABLE_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "8EAADB" };

function buildTable(rows) {
  const header = rows[0] || [];
  const ncol = Math.max(...rows.map((r) => r.length));
  const tableRows = rows.map((cells, ri) => {
    const padded = cells.slice(0, ncol);
    while (padded.length < ncol) padded.push("");
    return new TableRow({
      tableHeader: ri === 0,
      cantSplit: true,
      children: padded.map((c) => new TableCell({
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: "DCE6F1" } : undefined,
        children: [
          new Paragraph({
            bidirectional: true,
            alignment: AlignmentType.RIGHT,
            spacing: { after: 0, line: 260 },
            children: parseRuns(c, { font: FONT, size: TABLE_SIZE, bold: ri === 0 }),
          }),
        ],
      })),
    });
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: TABLE_BORDER, bottom: TABLE_BORDER, left: TABLE_BORDER, right: TABLE_BORDER,
      insideHorizontal: TABLE_BORDER, insideVertical: TABLE_BORDER,
    },
    rows: tableRows,
  });
}

function buildBlock(b) {
  switch (b.type) {
    case "h1":
      return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        bidirectional: true,
        alignment: AlignmentType.RIGHT,
        spacing: { before: 240, after: 200, line: 360 },
        children: [new TextRun({ text: b.text, bold: true, font: FONT, size: 32, color: "1F3864" })],
      });
    case "h2":
      return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        bidirectional: true,
        alignment: AlignmentType.RIGHT,
        spacing: { before: 240, after: 120, line: 360 },
        children: [new TextRun({ text: b.text, bold: true, font: FONT, size: 28, color: "2E5496" })],
      });
    case "h3":
      return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        bidirectional: true,
        alignment: AlignmentType.RIGHT,
        spacing: { before: 200, after: 100, line: 360 },
        children: [new TextRun({ text: b.text, bold: true, font: FONT, size: 26, color: "2E5496" })],
      });
    case "p":
      return new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 120, line: 360 },
        children: parseRuns(b.text, { font: FONT, size: BODY_SIZE }),
      });
    case "bullet":
      return new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.JUSTIFIED,
        indent: { right: 480 },
        spacing: { after: 80, line: 360 },
        children: [
          new TextRun({ text: "•  ", font: FONT, size: BODY_SIZE }),
          ...parseRuns(b.text, { font: FONT, size: BODY_SIZE }),
        ],
      });
    case "num":
      return new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.JUSTIFIED,
        indent: { right: 480 },
        spacing: { after: 80, line: 360 },
        children: [
          new TextRun({ text: `${faDigits(b.n)}.  `, font: FONT, size: BODY_SIZE }),
          ...parseRuns(b.text, { font: FONT, size: BODY_SIZE }),
        ],
      });
    case "quote":
      return new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.JUSTIFIED,
        indent: { right: 600 },
        spacing: { before: 120, after: 120, line: 360 },
        border: { right: { style: BorderStyle.SINGLE, size: 18, color: "2E5496", space: 8 } },
        children: parseRuns(b.text, { font: FONT, size: BODY_SIZE, color: "404040" }),
      });
    case "hr":
      return new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF", space: 1 } },
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "" })],
      });
    case "table":
      return buildTable(b.rows);
    case "code":
      return b.lines.map((l) => new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 0, line: 260 },
        shading: { type: ShadingType.CLEAR, fill: "F2F2F2" },
        children: [new TextRun({ text: l.length ? l : " ", font: FONT_MONO, size: CODE_SIZE })],
      }));
    default:
      return new Paragraph({ children: [new TextRun({ text: b.text })] });
  }
}

/* ---------- ساخت سند ---------- */
function main() {
  const files = fs.readdirSync(mdDir).filter((f) => f.endsWith(".md")).sort();
  const children = [];

  files.forEach((f, fi) => {
    const content = fs.readFileSync(path.join(mdDir, f), "utf8");
    const lines = content.split(/\r?\n/);
    const blocks = parseFile(lines);

    if (fi > 0) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }

    for (const b of blocks) {
      const el = buildBlock(b);
      if (Array.isArray(el)) children.push(...el);
      else children.push(el);
    }
  });

  const doc = new Document({
    creator: "دفتر آمایش سرزمین استان",
    title: "سند آمایش استانی — طرح توسعه‌یافته",
    description: "طرح توسعه‌یافته سامانه آمایش استانی (۱۸ فصل + پیوست‌ها)",
    styles: {
      default: {
        document: { run: { font: FONT, size: BODY_SIZE } },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 }, // ۲ سانتی‌متر
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "صفحه ", font: FONT, size: 18 }),
                  new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 }),
                ],
              }),
            ],
          }),
        },
        children,
      },
    ],
  });

  Packer.toBuffer(doc).then((buf) => {
    fs.writeFileSync(outPath, buf);
    const kb = (buf.length / 1024).toFixed(1);
    console.log(`OK: ${outPath}`);
    console.log(`فایل‌های ورودی: ${files.length}   اندازه خروجی: ${kb} KB`);
  }).catch((err) => {
    console.error("خطا در ساخت DOCX:", err);
    process.exit(1);
  });
}

main();
