/**
 * common.js
 * ドキュメント生成の共通スタイル・部品。
 * 各ドキュメントは spec.js / usermanual.js / setup.js / testchecklist.js に定義し、
 * build.js からまとめて生成する。
 */
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak, PageOrientation,
} = require('docx');

const FONT = 'Yu Gothic';
const CONTENT_WIDTH = 9000;            // A4縦・余白1inch時の本文幅（DXA）
const CONTENT_WIDTH_LANDSCAPE = 14600; // A4横・余白0.75inch時の本文幅（DXA）

const COLOR = {
  heading: '1F3864',
  sub: '2E5395',
  accent: '44546A',
  noteFill: 'FFF4E5',
  warnFill: 'FDE7E9',
  tableHeadFill: 'DCE6F1',
  zebraFill: 'F5F7FA',
};

function styles() {
  return {
    default: {
      document: { run: { font: FONT, size: 21 }, paragraph: { spacing: { line: 300, after: 100 } } },
      title: {
        run: { font: FONT, size: 48, bold: true, color: COLOR.heading },
        paragraph: { spacing: { after: 240 }, alignment: AlignmentType.CENTER },
      },
      heading1: {
        run: { font: FONT, size: 30, bold: true, color: COLOR.heading },
        paragraph: { spacing: { before: 400, after: 160 } },
      },
      heading2: {
        run: { font: FONT, size: 25, bold: true, color: COLOR.sub },
        paragraph: { spacing: { before: 280, after: 120 } },
      },
      heading3: {
        run: { font: FONT, size: 22, bold: true, color: COLOR.accent },
        paragraph: { spacing: { before: 200, after: 100 } },
      },
    },
  };
}

function numbering() {
  return {
    config: [
      {
        reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '●', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 240 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '○', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 960, hanging: 240 } } } },
        ],
      },
      {
        reference: 'steps',
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 300 } } } },
        ],
      },
    ],
  };
}

const h1 = t => new Paragraph({ text: t, heading: HeadingLevel.HEADING_1 });
const h2 = t => new Paragraph({ text: t, heading: HeadingLevel.HEADING_2 });
const h3 = t => new Paragraph({ text: t, heading: HeadingLevel.HEADING_3 });

function p(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold: !!opts.bold, italics: !!opts.italics, color: opts.color })],
    spacing: opts.spacing || { after: 120 },
    alignment: opts.alignment,
  });
}

const bullet = (t, level = 0) =>
  new Paragraph({ children: [new TextRun(t)], numbering: { reference: 'bullets', level }, spacing: { after: 60 } });

const step = t =>
  new Paragraph({ children: [new TextRun(t)], numbering: { reference: 'steps', level: 0 }, spacing: { after: 60 } });

/** 補足・注意ボックス。kind='warn' で警告色になる */
function box(title, lines, kind = 'note', width = CONTENT_WIDTH) {
  const fill = kind === 'warn' ? COLOR.warnFill : COLOR.noteFill;
  const children = [
    new Paragraph({ children: [new TextRun({ text: title, bold: true })], spacing: { after: 60 } }),
    ...lines.map(l => new Paragraph({ children: [new TextRun(l)], spacing: { after: 40 } })),
  ];
  return new Table({
    columnWidths: [width],
    width: { size: width, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: width, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children,
      })],
    })],
  });
}

function cell(text, width, opts = {}) {
  const runs = String(text).split('\n').map(line =>
    new Paragraph({
      children: [new TextRun({ text: line, bold: !!opts.bold, size: opts.size || 19 })],
      spacing: { after: 0 },
      alignment: opts.alignment,
    })
  );
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    children: runs,
  });
}

/** headers: string[], rows: string[][], widths: number[]（表の幅は合計値になる） */
function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const headRow = new TableRow({
    tableHeader: true,
    children: headers.map((hText, i) => cell(hText, widths[i], { bold: true, fill: COLOR.tableHeadFill })),
  });
  const bodyRows = rows.map((r, ri) => new TableRow({
    children: r.map((c, ci) => cell(c, widths[ci], { fill: ri % 2 === 1 ? COLOR.zebraFill : undefined })),
  }));
  return new Table({
    columnWidths: widths,
    width: { size: total, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'AAB4C4' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAB4C4' },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'AAB4C4' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'AAB4C4' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'C6CFDC' },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'C6CFDC' },
    },
    rows: [headRow, ...bodyRows],
  });
}

const spacer = (after = 160) => new Paragraph({ text: '', spacing: { after } });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

/** 表紙 */
function coverPage(title, subtitle, meta, width = CONTENT_WIDTH) {
  const out = [
    new Paragraph({ text: '', spacing: { after: 1800 } }),
    new Paragraph({
      children: [new TextRun({ text: title, bold: true, size: 48, color: COLOR.heading })],
      alignment: AlignmentType.CENTER, spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: subtitle, size: 26, color: COLOR.sub })],
      alignment: AlignmentType.CENTER, spacing: { after: 1200 },
    }),
  ];
  const widths = [Math.round(width / 3), width - Math.round(width / 3)];
  out.push(new Table({
    columnWidths: widths,
    width: { size: width, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'DDE3EC' },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: meta.map(([k, v]) => new TableRow({
      children: [cell(k, widths[0], { bold: true }), cell(v, widths[1])],
    })),
  }));
  out.push(pageBreak());
  return out;
}

/** @param {{landscape?: boolean}} opts */
function buildDoc(sections, opts = {}) {
  const margin = opts.landscape
    ? { top: 1080, bottom: 1080, left: 1080, right: 1080 }
    : { top: 1440, bottom: 1440, left: 1440, right: 1440 };
  const page = opts.landscape
    ? { margin, size: { orientation: PageOrientation.LANDSCAPE } }
    : { margin };
  return new Document({
    styles: styles(),
    numbering: numbering(),
    sections: [{ properties: { page }, children: sections }],
  });
}

async function save(doc, path) {
  const fs = require('fs');
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(path, buffer);
  console.log('wrote', path);
}

module.exports = {
  h1, h2, h3, p, bullet, step, box, table, spacer, pageBreak, coverPage,
  buildDoc, save, CONTENT_WIDTH, CONTENT_WIDTH_LANDSCAPE, COLOR,
};
