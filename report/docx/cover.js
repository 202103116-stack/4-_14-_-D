// 표지 + 목차 (템플릿과 실제 보고서 공용)
const {
  AlignmentType, BorderStyle, Paragraph, TextRun, TableOfContents, PageBreak,
  Table, TableRow, TableCell, WidthType, VerticalAlign, ShadingType,
} = require('docx');
const { font, HUFS_B, HUFS_M, HUFS_L, COLOR } = require('./styles');

const blank = (n = 1) => Array.from({ length: n }, () => new Paragraph({ spacing: { after: 0, line: 360 }, children: [] }));

function cover({ kicker, title, subtitle, info, date, team }) {
  const line = (pos) => new Paragraph({
    spacing: { before: 0, after: 0, line: 120 },
    border: { [pos]: { style: BorderStyle.THICK_THIN_MEDIUM_GAP, size: 24, color: COLOR.accent, space: 1 } },
    children: [],
  });
  const infoW = [1900, 5300];
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const thin = { style: BorderStyle.SINGLE, size: 4, color: COLOR.rule };
  const infoTable = new Table({
    width: { size: 7200, type: WidthType.DXA },
    columnWidths: infoW,
    alignment: AlignmentType.CENTER,
    borders: { top: { style: BorderStyle.SINGLE, size: 12, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 12, color: '000000' }, left: none, right: none, insideHorizontal: thin, insideVertical: thin },
    rows: info.map(([k, v]) => new TableRow({
      children: [
        new TableCell({
          width: { size: infoW[0], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.CLEAR, color: 'auto', fill: COLOR.headFill },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [new Paragraph({ style: 'TableHead', children: [new TextRun(k)] })],
        }),
        new TableCell({
          width: { size: infoW[1], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
          margins: { top: 80, bottom: 80, left: 160, right: 100 },
          children: [new Paragraph({ style: 'TableText', children: [new TextRun(v)] })],
        }),
      ],
    })),
  });
  return [
    ...blank(3),
    new Paragraph({ style: 'CoverSub', spacing: { after: 240 }, children: [new TextRun({ text: kicker, font: font(HUFS_M), color: '404040' })] }),
    line('top'),
    new Paragraph({ style: 'CoverTitle', spacing: { before: 360, after: 200 }, children: [new TextRun(title)] }),
    new Paragraph({ style: 'CoverSub', spacing: { after: 360 }, children: [new TextRun(subtitle)] }),
    line('bottom'),
    ...blank(6),
    infoTable,
    ...blank(3),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: date, font: font(HUFS_L), size: 26 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: team, font: font(HUFS_B), size: 34 })] }),
  ];
}

function toc() {
  return [
    new Paragraph({ style: 'TocTitle', pageBreakBefore: true, children: [new TextRun('목  차')] }),
    new TableOfContents('목차', { hyperlink: true, headingStyleRange: '1-2' }),
  ];
}

module.exports = { cover, toc };
