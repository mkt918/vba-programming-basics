const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, BorderStyle, ShadingType, AlignmentType, WidthType,
} = require("docx");
const fs   = require("fs");
const path = require("path");

const C = {
  HEADER_BG:  "4a6fa5",
  HEADER_FG:  "FFFFFF",
  BUG_BG:     "FFEEEE",
  SECTION_BG: "EEF2F8",
  WHITE:      "FFFFFF",
};

const FONT      = "BIZ UDGothic";
const PAGE_W    = 11906;
const MARGIN    = 1000;
const CONTENT_W = PAGE_W - MARGIN * 2;

// ── ユーティリティ（1Aと共通）─────────────────────────
function cell(text, opts = {}) {
  const { bg = C.WHITE, bold = false, size = 20, color = "000000", w = null, align = AlignmentType.LEFT } = opts;
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  return new TableCell({
    width: w ? { size: w, type: WidthType.DXA } : undefined,
    shading: { fill: bg, type: ShadingType.CLEAR },
    borders: { top: border, bottom: border, left: border, right: border },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text, font: FONT, size, bold, color })] })],
  });
}
function hcell(text, w = null) {
  return cell(text, { bg: C.HEADER_BG, color: C.HEADER_FG, bold: true, size: 20, w });
}
function para(text, opts = {}) {
  const { bold = false, size = 20, color = "000000", spacing = 120, indent = 0 } = opts;
  return new Paragraph({
    indent: indent ? { left: indent } : undefined,
    spacing: { before: spacing, after: spacing },
    children: [new TextRun({ text, font: FONT, size, bold, color })],
  });
}
function sectionHeader(title) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.HEADER_BG } },
    shading: { fill: C.SECTION_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "■ " + title, font: FONT, size: 22, bold: true, color: C.HEADER_BG })],
  });
}
function emptyLines(n = 1) {
  return Array.from({ length: n }, () =>
    new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "" })] })
  );
}
function answerBox(lines = 4) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: Array.from({ length: lines }, () =>
      new TableRow({
        height: { value: 400, rule: "exact" },
        children: [new TableCell({
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { fill: C.WHITE, type: ShadingType.CLEAR },
          borders: { top: border, bottom: border, left: border, right: border },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: "" })] })],
        })],
      })
    ),
  });
}

// ── 1B 固有：VBAコード（バグ行: 行7）─────────────────
// バグ: Cells(1, 1).Value = a（行番号が固定されている）
// 正解: Cells(a, 1).Value = a
const CODE_LINES = [
  { no: "1",  code: "Sub DisplaySequence()" },
  { no: "2",  code: "    Dim n As Integer" },
  { no: "3",  code: "    Dim a As Integer" },
  { no: "4",  code: '    n = InputBox("いくつまで表示しますか？")' },
  { no: "5",  code: "    For a = 1 To n" },
  { no: "6",  code: "        Cells(1, 1).Value = a", bug: true },
  { no: "7",  code: "    Next a" },
  { no: "8",  code: "End Sub" },
];

function codeTable() {
  const colNO   = 600;
  const colCode = CONTENT_W - colNO;
  const border  = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [colNO, colCode],
    rows: [
      new TableRow({ children: [hcell("行", colNO), hcell("コード", colCode)] }),
      ...CODE_LINES.map(({ no, code, bug }) => {
        const bg = bug ? C.BUG_BG : C.WHITE;
        return new TableRow({
          children: [
            new TableCell({
              width: { size: colNO, type: WidthType.DXA },
              shading: { fill: bg, type: ShadingType.CLEAR },
              borders, margins: { top: 60, bottom: 60, left: 100, right: 100 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: no, font: "Courier New", size: 18 })] })],
            }),
            new TableCell({
              width: { size: colCode, type: WidthType.DXA },
              shading: { fill: bg, type: ShadingType.CLEAR },
              borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
              children: [new Paragraph({ children: [
                new TextRun({ text: code, font: "Courier New", size: 18 }),
                ...(bug ? [new TextRun({ text: "  ← バグあり", font: FONT, size: 18, color: "CC0000", bold: true })] : []),
              ]})],
            }),
          ],
        });
      }),
    ],
  });
}

// 仕様例（n=5の場合）
function specTable() {
  const col = Math.floor(CONTENT_W / 2);
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [col, col],
    rows: [
      new TableRow({ children: [hcell("セル", col), hcell("表示される値", col)] }),
      ...["1","2","3","4","5"].map((v, i) =>
        new TableRow({ children: [
          cell(`A${i+1}`, { w: col }),
          cell(v, { w: col }),
        ]})
      ),
    ],
  });
}

// ── ドキュメント組み立て ────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: 16838 },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    children: [
      // タイトルブロック
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [Math.floor(CONTENT_W * 0.6), Math.floor(CONTENT_W * 0.4)],
        rows: [new TableRow({ children: [
          new TableCell({
            shading: { fill: C.HEADER_BG, type: ShadingType.CLEAR },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "26-06-15 ループ基礎 1B", font: FONT, size: 28, bold: true, color: C.HEADER_FG })] }),
              new Paragraph({ children: [new TextRun({ text: "連番表示（For〜Next）", font: FONT, size: 22, color: "CCDDFF" })] }),
            ],
          }),
          new TableCell({
            shading: { fill: C.HEADER_BG, type: ShadingType.CLEAR },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "クラス：________  番号：____", font: FONT, size: 18, color: C.HEADER_FG })] }),
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "氏名：__________________", font: FONT, size: 18, color: C.HEADER_FG })] }),
              new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "日付：____年____月____日", font: FONT, size: 18, color: C.HEADER_FG })] }),
            ],
          }),
        ]})],
      }),

      ...emptyLines(1),

      sectionHeader("状況説明"),
      para("太郎さんは「n を入力すると 1 から n まで順番に A列に表示する」プログラムを作りました。", { indent: 200 }),
      para("しかし実行すると、A列に正しく連番が並びません。", { indent: 200 }),
      para("グループで原因を調べて、正しいプログラムに直してみましょう。", { indent: 200, bold: true }),

      ...emptyLines(1),

      sectionHeader("プログラムの仕様"),
      para("①  ユーザーに「n（何番まで表示するか）」を入力してもらう", { indent: 200 }),
      para("②  A1 〜 A(n) のセルに、1, 2, 3, … n と順番に表示する", { indent: 200 }),
      para("③  例）n = 5 の場合 →", { indent: 200 }),
      ...emptyLines(1),
      specTable(),
      ...emptyLines(1),

      sectionHeader("バグのあるプログラム"),
      para("※ 薄赤の行にバグが仕込まれています。", { indent: 200, color: "CC0000" }),
      ...emptyLines(1),
      codeTable(),
      ...emptyLines(1),

      sectionHeader("グループワーク"),

      para("【課題 1】　予想してみよう", { bold: true, size: 22 }),
      para("n = 5 で実行したとき、A列の結果はどうなると思いますか？グループで予想してから確認しましょう。", { indent: 200 }),
      ...emptyLines(1),
      answerBox(3),
      ...emptyLines(1),

      para("【課題 2】　バグを発見しよう", { bold: true, size: 22 }),
      para("何行目が間違っていますか？　なぜそうなるのか、理由を日本語で説明しましょう。", { indent: 200 }),
      ...emptyLines(1),
      answerBox(4),
      ...emptyLines(1),

      para("【課題 3】　修正しよう", { bold: true, size: 22 }),
      para("正しいコードに書き直して、Excel で動作を確認しましょう。", { indent: 200 }),
      ...emptyLines(1),
      answerBox(3),
    ],
  }],
});

const outPath = path.join(__dirname, "..", "worksheets", "step1", "1b_sequence-display.docx");
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log("生成完了:", outPath);
}).catch(err => { console.error(err); process.exit(1); });
