const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, BorderStyle, ShadingType, AlignmentType, WidthType,
} = require("docx");
const fs   = require("fs");
const path = require("path");

const C = {
  HEADER_BG:  "4a6fa5",
  HEADER_FG:  "FFFFFF",
  HOLE_BG:    "FFFFEE",
  SECTION_BG: "EEF2F8",
  WHITE:      "FFFFFF",
  LIGHT_GREEN:"F0FFF0",
};

const FONT      = "BIZ UDGothic";
const PAGE_W    = 11906;
const MARGIN    = 1000;
const CONTENT_W = PAGE_W - MARGIN * 2;

// ── ユーティリティ ────────────────────────────────────
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

// ── 課題①固有：仕様例表（n=5）──────────────────────
function specTable() {
  const c1 = 1800, c2 = 1800;
  const col = Math.floor(CONTENT_W / 2);
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [col, col],
    rows: [
      new TableRow({ children: [hcell("A列（連番）", col), hcell("B列（2倍）", col)] }),
      ...[[1,2],[2,4],[3,6],[4,8],[5,10]].map(([a, b]) =>
        new TableRow({ children: [
          cell(String(a), { w: col }),
          cell(String(b), { w: col }),
        ]})
      ),
    ],
  });
}

// ── トレース表 ─────────────────────────────────────
function traceTable() {
  const cols = [1200, 1500, 2000, 2000, 2000, 1206]; // sum = 9906 = CONTENT_W（調整）
  const totalW = cols.reduce((a,b) => a+b, 0);
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const holeBorder = { style: BorderStyle.SINGLE, size: 2, color: "AAAAAA" };
  const holeBorders = { top: holeBorder, bottom: holeBorder, left: holeBorder, right: holeBorder };

  const headers = ["ループ回数\n（a の値）", "For条件\n（a ≦ n？）", "Cells(a,1)\n（A列）", "Cells(a,1)*2\n（B列）", "実行", "Next a"];
  const rows = [
    ["1", "1 ≦ 5　✓", "1", "2", "A1=1, B1=2", "a=2"],
    ["2", "2 ≦ 5　✓", "2", "4", "A2=2, B2=4", "a=3"],
    ["3", "", "", "", "", ""],
    ["4", "", "", "", "", ""],
    ["5", "", "", "", "", ""],
    ["6", "6 ≦ 5　✗", "—", "—", "ループ終了", "—"],
  ];
  const filledRows = [0, 1]; // 1行目・2行目は見本として埋める

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      // ヘッダー
      new TableRow({
        children: headers.map((h, i) =>
          new TableCell({
            width: { size: cols[i], type: WidthType.DXA },
            shading: { fill: C.HEADER_BG, type: ShadingType.CLEAR },
            borders: { top: border, bottom: border, left: border, right: border },
            margins: { top: 60, bottom: 60, left: 80, right: 80 },
            children: h.split("\n").map(line =>
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: line, font: FONT, size: 16, bold: true, color: C.HEADER_FG })],
              })
            ),
          })
        ),
      }),
      // データ行
      ...rows.map((row, ri) =>
        new TableRow({
          height: { value: 480, rule: "exact" },
          children: row.map((val, ci) => {
            const isFilled = filledRows.includes(ri);
            const isLastRow = ri === rows.length - 1;
            const bg = isLastRow ? C.SECTION_BG : isFilled ? C.WHITE : C.HOLE_BG;
            return new TableCell({
              width: { size: cols[ci], type: WidthType.DXA },
              shading: { fill: bg, type: ShadingType.CLEAR },
              borders: isFilled || isLastRow ? { top: border, bottom: border, left: border, right: border } : { top: holeBorder, bottom: holeBorder, left: holeBorder, right: holeBorder },
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: val, font: FONT, size: 17, color: isLastRow ? "666666" : "000000" })],
              })],
            });
          }),
        })
      ),
    ],
  });
}

// ── 穴埋めコード表 ────────────────────────────────
// バグなし版（穴埋め形式で書かせる）
const CODE_LINES = [
  { no: "1",  code: "Sub DisplayMultiColumn()" },
  { no: "2",  code: "    Dim n As Integer" },
  { no: "3",  code: "    Dim a As Integer" },
  { no: "4",  code: '    n = InputBox("いくつまで表示しますか？")' },
  { no: "5",  code: "    For a = 1 To n" },
  { no: "6",  code: "        Cells(a, 1).Value = [     ]", hole: true },
  { no: "7",  code: "        Cells(a, 2).Value = [     ]", hole: true },
  { no: "8",  code: "    Next a" },
  { no: "9",  code: "End Sub" },
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
      new TableRow({ children: [hcell("行", colNO), hcell("コード（[ ] に入るものを書こう）", colCode)] }),
      ...CODE_LINES.map(({ no, code, hole }) => {
        const bg = hole ? C.HOLE_BG : C.WHITE;
        return new TableRow({
          height: { value: 420, rule: "exact" },
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
              children: [new Paragraph({ children: [new TextRun({ text: code, font: "Courier New", size: 18 })] })],
            }),
          ],
        });
      }),
    ],
  });
}

// ── プラス課題：仕様例表（n=3, ×1〜×9）────────────
function plusSpecTable() {
  const cols = Array(10).fill(Math.floor(CONTENT_W / 10));
  cols[9] = CONTENT_W - cols[0] * 9; // 端数調整
  const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const headers = ["行", "A(×1)", "B(×2)", "C(×3)", "D(×4)", "E(×5)", "F(×6)", "G(×7)", "H(×8)", "I(×9)"];
  const data = [1, 2, 3].map(a => [String(a), ...Array.from({length:9}, (_,i) => String(a*(i+1)))]);
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      new TableRow({
        children: headers.map((h, i) => new TableCell({
          width: { size: cols[i], type: WidthType.DXA },
          shading: { fill: C.HEADER_BG, type: ShadingType.CLEAR },
          borders: { top: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"}, bottom: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"}, left: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"}, right: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"} },
          margins: { top: 60, bottom: 60, left: 60, right: 60 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: FONT, size: 16, bold: true, color: C.HEADER_FG })] })],
        })),
      }),
      ...data.map(row =>
        new TableRow({
          height: { value: 380, rule: "exact" },
          children: row.map((val, i) => new TableCell({
            width: { size: cols[i], type: WidthType.DXA },
            shading: { fill: C.WHITE, type: ShadingType.CLEAR },
            borders: { top: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"}, bottom: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"}, left: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"}, right: {style:BorderStyle.SINGLE,size:1,color:"BBBBBB"} },
            margins: { top: 60, bottom: 60, left: 60, right: 60 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: val, font: FONT, size: 17 })] })],
          })),
        })
      ),
    ],
  });
}

// ── プラス課題：穴埋めコード表（×3〜×9 を穴にする）──
const PLUS_CODE_LINES = [
  { no: "1",  code: "Sub TimesTable()" },
  { no: "2",  code: "    Dim n As Integer" },
  { no: "3",  code: "    Dim a As Integer" },
  { no: "4",  code: '    n = InputBox("いくつまで表示しますか？")' },
  { no: "5",  code: "    For a = 1 To n" },
  { no: "6",  code: "        Cells(a, 1).Value = a          ' A列：連番（1倍）" },
  { no: "7",  code: "        Cells(a, 2).Value = a * 2      ' B列：2倍" },
  { no: "8",  code: "        Cells(a, [  ]).Value = a * [  ]  ' C列：3倍", hole: true },
  { no: "9",  code: "        Cells(a, [  ]).Value = a * [  ]  ' D列：4倍", hole: true },
  { no: "10", code: "        Cells(a, [  ]).Value = a * [  ]  ' E列：5倍", hole: true },
  { no: "11", code: "        Cells(a, [  ]).Value = a * [  ]  ' F列：6倍", hole: true },
  { no: "12", code: "        Cells(a, [  ]).Value = a * [  ]  ' G列：7倍", hole: true },
  { no: "13", code: "        Cells(a, [  ]).Value = a * [  ]  ' H列：8倍", hole: true },
  { no: "14", code: "        Cells(a, [  ]).Value = a * [  ]  ' I列：9倍", hole: true },
  { no: "15", code: "    Next a" },
  { no: "16", code: "End Sub" },
];

function plusCodeTable() {
  const colNO   = 600;
  const colCode = CONTENT_W - colNO;
  const border  = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [colNO, colCode],
    rows: [
      new TableRow({ children: [hcell("行", colNO), hcell("コード（[ ] に列番号と倍率を書こう）", colCode)] }),
      ...PLUS_CODE_LINES.map(({ no, code, hole }) => {
        const bg = hole ? C.HOLE_BG : C.WHITE;
        return new TableRow({
          height: { value: 420, rule: "exact" },
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
              children: [new Paragraph({ children: [new TextRun({ text: code, font: "Courier New", size: 17 })] })],
            }),
          ],
        });
      }),
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
      // タイトル
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [Math.floor(CONTENT_W * 0.6), Math.floor(CONTENT_W * 0.4)],
        rows: [new TableRow({ children: [
          new TableCell({
            shading: { fill: C.HEADER_BG, type: ShadingType.CLEAR },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "26-06-15 応用課題①", font: FONT, size: 28, bold: true, color: C.HEADER_FG })] }),
              new Paragraph({ children: [new TextRun({ text: "複数列出力（A列＋B列）", font: FONT, size: 22, color: "CCDDFF" })] }),
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
      para("太郎さんが作ったプログラムをさらに発展させます。", { indent: 200 }),
      para("今回は「A列に連番」「B列にその2倍の値」を同時に表示するプログラムを完成させましょう。", { indent: 200 }),
      para("コードの穴埋めと、処理の流れを日本語で書く練習もします。", { indent: 200, bold: true }),

      ...emptyLines(1),

      sectionHeader("プログラムの仕様"),
      para("①  ユーザーに「n（何行まで出力するか）」を入力してもらう", { indent: 200 }),
      para("②  A列（1列目）には 1, 2, 3, … n を順番に表示する", { indent: 200 }),
      para("③  B列（2列目）には A列の値の 2倍 を表示する", { indent: 200 }),
      para("④  例）n = 5 の場合 →", { indent: 200 }),
      ...emptyLines(1),
      specTable(),
      ...emptyLines(1),

      sectionHeader("課題 1　日本語で処理の流れを書こう"),
      para("ループの中で何をしているか、1ステップずつ日本語で書いてみましょう。", { indent: 200 }),
      para("（例：「変数 a が 1 のとき、A1に____を、B1に____を入れる」のように書く）", { indent: 200, color: "666666" }),
      ...emptyLines(1),
      answerBox(5),
      ...emptyLines(1),

      sectionHeader("課題 2　トレース表を完成させよう"),
      para("1行目と2行目は例として埋めてあります。3〜5行目を自分で埋めましょう。", { indent: 200 }),
      ...emptyLines(1),
      traceTable(),
      ...emptyLines(1),

      sectionHeader("課題 3　コードの穴埋めをしよう"),
      para("薄黄の行 [ ] に当てはまるコードを書いて、プログラムを完成させましょう。", { indent: 200 }),
      para("完成したら Excel に入力して動作を確認しましょう。", { indent: 200 }),
      ...emptyLines(1),
      codeTable(),

      ...emptyLines(2),

      // ── プラス課題 ──────────────────────────────────
      new Paragraph({
        spacing: { before: 200, after: 100 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "e67e22" } },
        shading: { fill: "FFF3E0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "⭐ プラス課題　2倍〜9倍を全列に出力しよう", font: FONT, size: 22, bold: true, color: "e67e22" })],
      }),
      para("基本課題では A列に連番・B列に2倍を出力しました。", { indent: 200 }),
      para("プラス課題では C列に3倍・D列に4倍・…・I列に9倍 まで一気に出力しましょう！", { indent: 200, bold: true }),
      ...emptyLines(1),
      para("例）n = 3 の場合（A〜I列）", { indent: 200 }),
      ...emptyLines(1),
      plusSpecTable(),
      ...emptyLines(1),
      para("【課題】薄黄の行の [ ] を埋めて完成させましょう。（ヒント：列番号と倍率は同じ数字になります）", { indent: 200, bold: true }),
      ...emptyLines(1),
      plusCodeTable(),
      ...emptyLines(1),
      para("◎ 発展：パターンに気づいたら、For〜Next を入れ子にしてまとめて書くことも挑戦してみよう！", { indent: 200, color: "666666" }),
    ],
  }],
});

const outPaths = [
  path.join(__dirname, "..", "worksheets", "step1", "task1_multi-column.docx"),
  path.join(__dirname, "..", "docs", "downloads", "task1_multi-column.docx"),
];
Packer.toBuffer(doc).then(buf => {
  outPaths.forEach(p => { fs.writeFileSync(p, buf); console.log("生成完了:", p); });
}).catch(err => { console.error(err); process.exit(1); });
