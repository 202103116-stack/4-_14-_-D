// 14조 팀보고서 공통 스타일: 공공기관 보고서 양식(A4)
// 본문 휴먼명조, 제목·표지·캡션 한국외대체(B/M/L), 표 안 글자 맑은 고딕
// 템플릿(make_template.js)과 실제 보고서(make_report.js)가 이 파일 하나를 같이 쓴다.
const {
  AlignmentType, BorderStyle, Footer, Header, LevelFormat, LevelSuffix, PageNumber,
  Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType,
  VerticalAlign, SequentialIdentifier, ImageRun, HeightRule,
} = require('docx');

const SERIF = '휴먼명조';
const SANS = '맑은 고딕';
// 한국외대체: 굵기별로 별도 글꼴(전부 Regular)이라 bold를 켜지 않는다(가짜 굵게 방지)
const HUFS_B = 'HUFS B';
const HUFS_M = 'HUFS M';
const HUFS_L = 'HUFS L';
const font = (name) => ({ ascii: name, eastAsia: name, hAnsi: name, cs: name });

// 단위: 1mm = 56.7 DXA, 1pt = 2 half-point(size) / 20 DXA(spacing)
const MM = 56.7;
const PAGE = {
  size: { width: 11906, height: 16838 }, // A4
  margin: {
    top: Math.round(20 * MM), bottom: Math.round(15 * MM),
    left: Math.round(20 * MM), right: Math.round(20 * MM),
    header: Math.round(10 * MM), footer: Math.round(8 * MM),
  },
};
const CONTENT_WIDTH = PAGE.size.width - PAGE.margin.left - PAGE.margin.right; // 9638 DXA = 170mm

const COLOR = {
  text: '000000',
  gray: '595959',
  rule: '808080',
  headFill: 'D9E2F3', // 표 머리글 음영(연한 남색)
  noteFill: 'F2F2F2',
  accent: '1F3864', // 제목 1 구분선·표지 선
};

const LINE_160 = 384; // 줄 간격 160% (240 = 100%)

const styles = {
  default: {
    document: {
      run: { font: font(SERIF), size: 22, color: COLOR.text }, // 11pt
      paragraph: { spacing: { line: LINE_160, before: 0, after: 60 }, alignment: AlignmentType.JUSTIFIED },
    },
    heading1: {
      run: { font: font(HUFS_B), size: 34, bold: false, color: COLOR.text }, // 17pt 한국외대체 B
      paragraph: {
        spacing: { before: 360, after: 160, line: 300 },
        keepNext: true, keepLines: true,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: COLOR.accent, space: 4 } },
        alignment: AlignmentType.LEFT,
      },
    },
    heading2: {
      run: { font: font(HUFS_B), size: 27, bold: false, color: COLOR.text }, // 13.5pt 한국외대체 B
      paragraph: { spacing: { before: 240, after: 100, line: 300 }, keepNext: true, keepLines: true, alignment: AlignmentType.LEFT },
    },
    heading3: {
      run: { font: font(HUFS_M), size: 24, bold: false, color: COLOR.text }, // 12pt 한국외대체 M
      paragraph: { spacing: { before: 180, after: 80, line: 300 }, keepNext: true, keepLines: true, alignment: AlignmentType.LEFT },
    },
    heading4: {
      run: { font: font(SERIF), size: 22, bold: true, color: COLOR.text }, // 11pt
      paragraph: { spacing: { before: 120, after: 60, line: 300 }, keepNext: true, keepLines: true, alignment: AlignmentType.LEFT },
    },
  },
  paragraphStyles: [
    {
      id: 'Caption', name: 'caption', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: font(HUFS_M), size: 19, bold: false, color: COLOR.text }, // 9.5pt 한국외대체 M
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 60, after: 240, line: 276 } },
    },
    {
      id: 'CoverTitle', name: '표지 제목', basedOn: 'Normal', quickFormat: true,
      run: { font: font(HUFS_B), size: 56, bold: false }, // 28pt 한국외대체 B
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120, line: 360 } },
    },
    {
      id: 'CoverSub', name: '표지 부제', basedOn: 'Normal', quickFormat: true,
      run: { font: font(HUFS_L), size: 28, color: COLOR.gray }, // 14pt 한국외대체 L
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60, line: 300 } },
    },
    {
      id: 'TocTitle', name: '목차 제목', basedOn: 'Normal', quickFormat: true,
      run: { font: font(HUFS_B), size: 36, bold: false, characterSpacing: 40 }, // 18pt 한국외대체 B
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 360 } },
    },
    {
      id: 'TableHead', name: '표 머리글', basedOn: 'Normal', quickFormat: true,
      run: { font: font(SANS), size: 20, bold: true }, // 10pt
      paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0, line: 276 } },
    },
    {
      id: 'TableText', name: '표 본문', basedOn: 'Normal', quickFormat: true,
      run: { font: font(SANS), size: 20 }, // 10pt
      paragraph: { alignment: AlignmentType.LEFT, spacing: { before: 0, after: 0, line: 276 } },
    },
    {
      id: 'Note', name: '참고 박스', basedOn: 'Normal', quickFormat: true,
      run: { font: font(SANS), size: 19 }, // 9.5pt
      paragraph: {
        alignment: AlignmentType.LEFT,
        spacing: { before: 120, after: 160, line: 312 },
        indent: { left: 113, right: 113 },
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: COLOR.noteFill },
        border: {
          top: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF', space: 4 },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF', space: 4 },
          left: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF', space: 4 },
          right: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF', space: 4 },
        },
      },
    },
    {
      id: 'Source', name: '출처', basedOn: 'Normal', quickFormat: true,
      run: { font: font(SANS), size: 16, color: COLOR.gray }, // 8pt
      paragraph: { alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 160, line: 240 } },
    },
    {
      id: 'Figure', name: '그림', basedOn: 'Normal', quickFormat: true,
      paragraph: { alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 0, line: 240 } },
    },
  ],
};

// 제목 번호: Ⅰ. → 1. → 가. → 1)  (postprocess.py가 제목 1~4 스타일에 연결해 자동 번호로 만든다)
// 글머리: □ → ○ → - → ·
const numbering = {
  config: [
    {
      reference: 'heading',
      levels: [
        { level: 0, format: LevelFormat.UPPER_ROMAN, text: '%1.', suffix: LevelSuffix.SPACE, style: { paragraph: { indent: { left: 0, hanging: 0 } } } },
        { level: 1, format: LevelFormat.DECIMAL, text: '%2.', suffix: LevelSuffix.SPACE, style: { paragraph: { indent: { left: 0, hanging: 0 } } } },
        { level: 2, format: LevelFormat.GANADA, text: '%3.', suffix: LevelSuffix.SPACE, style: { paragraph: { indent: { left: 200, hanging: 0 } } } },
        { level: 3, format: LevelFormat.DECIMAL, text: '%4)', suffix: LevelSuffix.SPACE, style: { paragraph: { indent: { left: 400, hanging: 0 } } } },
      ],
    },
    {
      reference: 'bullet',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '□', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 340, hanging: 340 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '○', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 680, hanging: 300 } } } },
        { level: 2, format: LevelFormat.BULLET, text: '-', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 260 } } } },
        { level: 3, format: LevelFormat.BULLET, text: '·', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1300, hanging: 240 } } } },
      ],
    },
  ],
};

// ── 문단 도우미 ─────────────────────────────────────────
// 인라인 **굵게** 를 TextRun으로 나눔
function runs(text, opts = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map((p) => (p.startsWith('**') ? new TextRun({ text: p.slice(2, -2), bold: true, ...opts }) : new TextRun({ text: p, ...opts })));
}
const H = (level, text) => new Paragraph({ heading: `Heading${level}`, children: runs(text) });
const P = (text, opts = {}) => new Paragraph({ children: runs(text), ...opts });
const B = (text, level = 0) => new Paragraph({ numbering: { reference: 'bullet', level }, children: runs(text), alignment: AlignmentType.LEFT });
const NOTE = (text) => new Paragraph({ style: 'Note', children: runs('※ ' + text) });
const SOURCE = (text) => new Paragraph({ style: 'Source', children: runs(text) });

// [표 n] / [그림 n] — SEQ 필드라서 Word에서 자동 번호
const tableCaption = (title) => new Paragraph({
  style: 'Caption', keepNext: true, spacing: { before: 200, after: 80 },
  children: [new TextRun('[표 '), new SequentialIdentifier('표'), new TextRun('] ' + title)],
});
const figureCaption = (title) => new Paragraph({
  style: 'Caption',
  children: [new TextRun('[그림 '), new SequentialIdentifier('그림'), new TextRun('] ' + title)],
});

// 표: 위·아래 굵은 선, 안쪽 가는 선, 머리글 음영 (공공기관 표 양식)
function table(headers, rows, widths, { align } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const thin = { style: BorderStyle.SINGLE, size: 4, color: COLOR.rule };
  const thick = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
  const cell = (text, w, head, colIdx) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: COLOR.headFill } : undefined,
    children: String(text).split('\n').map((line) => new Paragraph({
      style: head ? 'TableHead' : 'TableText',
      alignment: head ? AlignmentType.CENTER : (align && align[colIdx]) || AlignmentType.LEFT,
      children: runs(line),
    })),
  });
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    alignment: AlignmentType.CENTER,
    borders: { top: thick, bottom: thick, left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }, insideHorizontal: thin, insideVertical: thin },
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => cell(h, widths[i], true, i)) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false, i)) })),
    ],
  });
}

// 그림: 폭(mm) 기준으로 비율 유지해 넣음
function figure(data, pxW, pxH, widthMm, type = 'png') {
  const wPx = Math.round(widthMm * 3.78); // 96dpi 기준 mm→px
  const hPx = Math.round(wPx * pxH / pxW);
  return new Paragraph({ style: 'Figure', children: [new ImageRun({ type, data, transformation: { width: wPx, height: hPx } })] });
}

function figurePlaceholder(text = '(그림·스크린샷 자리)', heightMm = 45, widthMm = 110) {
  const dash = { style: BorderStyle.DASHED, size: 6, color: '7F7F7F' };
  const w = Math.round(widthMm * MM);
  return new Table({
    width: { size: w, type: WidthType.DXA }, columnWidths: [w], alignment: AlignmentType.CENTER,
    borders: { top: dash, bottom: dash, left: dash, right: dash, insideHorizontal: dash, insideVertical: dash },
    rows: [new TableRow({
      height: { value: Math.round(heightMm * MM), rule: HeightRule.EXACT }, cantSplit: true,
      children: [new TableCell({
        width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ style: 'Figure', keepNext: true, children: [new TextRun({ text, font: font(SANS), size: 20, color: '7F7F7F' })] })],
      })],
    })],
  });
}

const RUNNING_TITLE = '빅데이터분석 4주차 팀과제  |  14조 · 미션 D';
const headers = {
  default: new Header({
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { after: 0, line: 240 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR.rule, space: 4 } },
      children: [new TextRun({ text: RUNNING_TITLE, font: font(HUFS_L), size: 17, color: COLOR.gray })],
    })],
  }),
  first: new Header({ children: [new Paragraph({ children: [] })] }),
};
const footers = {
  default: new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 240 },
      children: [new TextRun({ children: ['- ', PageNumber.CURRENT, ' -'], font: font(SANS), size: 18 })],
    })],
  }),
  first: new Footer({ children: [new Paragraph({ children: [] })] }),
};

const frontFooter = new Footer({ children: [new Paragraph({ children: [] })] });
// 앞부분(표지·목차): 표지는 머리글·쪽 번호 없음, 목차 쪽은 머리글만
const frontSection = (children) => ({
  properties: { page: PAGE, titlePage: true },
  headers, footers: { default: frontFooter, first: footers.first }, children,
});
// 본문: 쪽 번호 1부터
const bodySection = (children) => ({
  properties: { page: { ...PAGE, pageNumbers: { start: 1 } } },
  headers: { default: headers.default }, footers: { default: footers.default }, children,
});

module.exports = {
  SERIF, SANS, HUFS_B, HUFS_M, HUFS_L, font, MM, PAGE, CONTENT_WIDTH, COLOR, styles, numbering,
  runs, H, P, B, NOTE, SOURCE, tableCaption, figureCaption, table, figure,
  headers, footers, frontSection, bodySection, figurePlaceholder, HeightRule,
};
