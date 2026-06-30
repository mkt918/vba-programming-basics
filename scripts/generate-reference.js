const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require("docx");
const fs   = require("fs");
const path = require("path");

// ── カラー ───────────────────────────────────────────────
const HEADER_BG = "4a6fa5";
const CAT_COLORS = {
  "変数・宣言": { bg: "e3f2fd", fg: "1565c0" },
  "入力":       { bg: "f3e5f5", fg: "6a1b9a" },
  "セル出力":   { bg: "e8f5e9", fg: "2e7d32" },
  "ループ":     { bg: "fff3e0", fg: "e65100" },
  "条件分岐":   { bg: "fce4ec", fg: "c62828" },
  "演算・計算": { bg: "f1f8e9", fg: "558b2f" },
  "文字列":     { bg: "e0f7fa", fg: "00695c" },
  "その他":     { bg: "f5f5f5", fg: "555555" },
};

const FONT  = "BIZ UDGothic";
const MONO  = "Consolas";
const PW    = 11906;
const MAR   = 1000;
const CW    = PW - MAR * 2;   // 9906
const LABEL = 1500;
const VALUE = CW - LABEL;     // 8406

// ── ヘルパー ─────────────────────────────────────────────
function noBorder() {
  const n = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: n, bottom: n, left: n, right: n, insideH: n, insideV: n };
}
function gridBorder(color = "cccccc") {
  const b = { style: BorderStyle.SINGLE, size: 1, color };
  return { top: b, bottom: b, left: b, right: b, insideH: b, insideV: b };
}

function mkCell(children, opts = {}) {
  return new TableCell({
    borders: opts.borders ?? noBorder(),
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    shading: opts.bg ? { fill: opts.bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: opts.ml ?? 120, right: 120 },
    verticalAlign: "top",
    children,
  });
}

function labelCell(text, bg, fg) {
  return mkCell(
    [new Paragraph({ children: [new TextRun({ text, font: FONT, size: 17, bold: true, color: fg })] })],
    { bg, w: LABEL, borders: noBorder() }
  );
}

// 複数行コードブロック（配列）
function codeBlock(lines, bg, fg) {
  return lines.map((line, i) =>
    new Paragraph({
      spacing: { before: i === 0 ? 40 : 0, after: i === lines.length - 1 ? 40 : 0 },
      shading: { fill: bg, type: ShadingType.CLEAR },
      indent: { left: 80, right: 80 },
      children: [new TextRun({ text: line, font: MONO, size: 16, color: fg })],
    })
  );
}

function notePara(text) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    children: [new TextRun({ text, font: FONT, size: 17, color: "555555" })],
  });
}

// ── リファレンスデータ ────────────────────────────────────
const REFS = [
  {
    cat: "変数・宣言",
    items: [
      {
        want: "変数を宣言したい",
        syntax: ["Dim 変数名 As 型"],
        note: "よく使う型：Integer（整数）/ Double（小数）/ String（文字列）/ Boolean（True/False）",
        example: ["Dim n As Integer", "Dim name As String"],
      },
      {
        want: "変数に値を代入したい",
        syntax: ["変数名 = 値"],
        note: "= の左辺に変数名、右辺に入れたい値を書く。",
        example: ["n = 10", 'name = "太郎"'],
      },
    ],
  },
  {
    cat: "入力",
    items: [
      {
        want: "ユーザーに数値・文字を入力させたい",
        syntax: ['変数名 = InputBox("メッセージ")'],
        note: "ダイアログボックスが表示され、入力した値を変数に格納できる。",
        example: ['n = InputBox("いくつまで表示しますか？")'],
      },
    ],
  },
  {
    cat: "セル出力",
    items: [
      {
        want: "セルに値を書き込みたい",
        syntax: ["Cells(行, 列).Value = 値"],
        note: "行・列は数値で指定。列は A=1, B=2, C=3 …",
        example: [
          "Cells(1, 1).Value = 100    ' A1 に 100",
          "Cells(a, 2).Value = a * 2  ' B列に 2倍",
        ],
      },
      {
        want: '"A1" など列名でセルを指定したい',
        syntax: ['Range("セル番地").Value = 値'],
        note: 'Cells は数値指定、Range は "A1" など文字列で指定。どちらも使える。',
        example: ['Range("A1").Value = "合計"', 'Range("B2").Value = 500'],
      },
      {
        want: "セルの値を読み取りたい",
        syntax: ["変数名 = Cells(行, 列).Value"],
        note: "セルの値を変数に取り込む。",
        example: ["x = Cells(1, 1).Value  ' A1 の値を x に"],
      },
      {
        want: "セルの中身を消したい",
        syntax: ["Cells(行, 列).ClearContents"],
        note: "値だけ消す。書式を含めて消すには .Clear。",
        example: ["Cells(1, 1).ClearContents", 'Range("A1:B10").ClearContents'],
      },
    ],
  },
  {
    cat: "ループ",
    items: [
      {
        want: "決まった回数だけ繰り返したい",
        syntax: ["For 変数 = 開始 To 終了", "    ' 繰り返す処理", "Next 変数"],
        note: "変数が開始から終了まで 1 ずつ増えながら繰り返す。",
        example: ["For a = 1 To 5", "    Cells(a, 1).Value = a", "Next a"],
      },
      {
        want: "2つ飛ばしなど刻み幅を変えて繰り返したい",
        syntax: ["For 変数 = 開始 To 終了 Step 刻み", "    ' 繰り返す処理", "Next 変数"],
        note: "Step に負の数を使うと逆順にも繰り返せる。",
        example: [
          "For a = 2 To 10 Step 2   ' 2,4,6,8,10",
          "    Cells(a / 2, 1).Value = a",
          "Next a",
        ],
      },
      {
        want: "出力行をループ変数と別に管理したい",
        syntax: [
          "Dim row As Integer",
          "row = 1",
          "For a = 開始 To 終了",
          "    Cells(row, 1).Value = a",
          "    row = row + 1",
          "Next a",
        ],
        note: "偶数・奇数だけ表示するときなど、出力行を別カウンタで管理する。",
        example: [
          "Dim row As Integer : row = 1",
          "For a = 1 To n",
          "    If a Mod 2 = 0 Then",
          "        Cells(row, 1).Value = a",
          "        row = row + 1",
          "    End If",
          "Next a",
        ],
      },
    ],
  },
  {
    cat: "条件分岐",
    items: [
      {
        want: "条件によって処理を変えたい（If）",
        syntax: ["If 条件 Then", "    ' 条件が True のとき", "End If"],
        note: "条件が True のときだけ処理を実行する。",
        example: ["If a > 10 Then", '    Cells(1, 1).Value = "大きい"', "End If"],
      },
      {
        want: "条件を満たすときとそうでないときで処理を分けたい（If〜Else）",
        syntax: [
          "If 条件 Then",
          "    ' True のとき",
          "Else",
          "    ' False のとき",
          "End If",
        ],
        note: "Else 以降は条件が False のときに実行される。",
        example: [
          "If a Mod 2 = 0 Then",
          "    Cells(rowA, 1).Value = a  ' 偶数→A列",
          "Else",
          "    Cells(rowB, 2).Value = a  ' 奇数→B列",
          "End If",
        ],
      },
      {
        want: "3つ以上の条件で分けたい（ElseIf）",
        syntax: [
          "If 条件1 Then",
          "    ' 条件1 が True",
          "ElseIf 条件2 Then",
          "    ' 条件2 が True",
          "Else",
          "    ' どれも False",
          "End If",
        ],
        note: "",
        example: [
          "If a Mod 15 = 0 Then",
          '    Cells(a,1).Value = "FizzBuzz"',
          "ElseIf a Mod 3 = 0 Then",
          '    Cells(a,1).Value = "Fizz"',
          "ElseIf a Mod 5 = 0 Then",
          '    Cells(a,1).Value = "Buzz"',
          "Else",
          "    Cells(a,1).Value = a",
          "End If",
        ],
      },
    ],
  },
  {
    cat: "演算・計算",
    items: [
      {
        want: "割り算の余りを求めたい（Mod）",
        syntax: ["結果 = 数値 Mod 割る数"],
        note: "余りが 0 → 割り切れる（偶数、3の倍数など）  余りが 1 → 割り切れない",
        example: [
          "10 Mod 3   ' → 1",
          " 9 Mod 3   ' → 0 （3の倍数）",
          " a Mod 2 = 0   ' a が偶数なら True",
        ],
      },
      {
        want: "足す・引く・かける・割るを計算したい",
        syntax: [
          "+   足し算",
          "-   引き算",
          "*   掛け算",
          "/   割り算（小数あり）",
          "\\  整数除算（小数切捨て）",
        ],
        note: "",
        example: [
          "Cells(a,2).Value = a * 2   ' 2倍",
          "Cells(a,3).Value = a + 10  ' 10足す",
          "7 / 2   ' → 3.5",
          "7 \\ 2  ' → 3 （整数部分のみ）",
        ],
      },
      {
        want: "大小・等しいかを比べたい（比較演算子）",
        syntax: [
          "=    等しい",
          "<>   等しくない",
          ">    より大きい",
          "<    より小さい",
          ">=   以上",
          "<=   以下",
        ],
        note: "",
        example: [
          "If a = 5 Then      ' a が 5 のとき",
          "If a >= 10 Then    ' a が 10 以上のとき",
          "If a <> 0 Then     ' a が 0 でないとき",
        ],
      },
    ],
  },
  {
    cat: "文字列",
    items: [
      {
        want: "文字列をつなげたい（&）",
        syntax: ["文字列1 & 文字列2"],
        note: "変数と文字を組み合わせて表示するときに使う。",
        example: [
          'Cells(1,1).Value = "合計：" & goukei',
          'Cells(a,1).Value = a & "番"',
        ],
      },
      {
        want: "数値と文字列を変換したい",
        syntax: [
          "CStr(数値)    数値 → 文字列",
          "CInt(文字列)  文字列 → 整数",
          "CDbl(文字列)  文字列 → 小数",
        ],
        note: "",
        example: [
          's = CStr(100)   \' s = "100"',
          "n = CInt(\"42\")  ' n = 42",
        ],
      },
    ],
  },
  {
    cat: "その他",
    items: [
      {
        want: "メッセージボックスを表示したい",
        syntax: ['MsgBox "メッセージ"'],
        note: "プログラムの途中・終了後に結果を知らせるときに便利。",
        example: ['MsgBox "完了しました！"', 'MsgBox "合計：" & goukei'],
      },
      {
        want: "コードにメモ・説明を書きたい（コメント）",
        syntax: ["' ここに説明を書く"],
        note: "行頭または行末に ' を書くと、その後ろは実行されない。",
        example: [
          "' ループで 1〜n を出力する",
          "For a = 1 To n  ' n 回繰り返す",
        ],
      },
      {
        want: "マクロ（プロシージャ）を作りたい",
        syntax: ["Sub マクロ名()", "    ' ここに処理を書く", "End Sub"],
        note: "ひとまとまりの処理に名前をつけて定義する。Excel から実行できる。",
        example: [
          "Sub DisplayNumbers()",
          "    Dim a As Integer",
          "    For a = 1 To 5",
          "        Cells(a, 1).Value = a",
          "    Next a",
          "End Sub",
        ],
      },
    ],
  },
];

// ── ドキュメント構築 ──────────────────────────────────────
const children = [];

// タイトルブロック
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    shading: { fill: HEADER_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "VBA 逆引きリファレンス", font: FONT, size: 38, bold: true, color: "FFFFFF" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    shading: { fill: HEADER_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "「やりたいこと」から VBA 構文を調べるリファレンス", font: FONT, size: 19, color: "dce8ff" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 360 },
    shading: { fill: HEADER_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "ビジネス系高等学校 プログラミング実務コース", font: FONT, size: 16, color: "aac4e8" })],
  })
);

// カテゴリ目次
children.push(
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "カテゴリ一覧", font: FONT, size: 22, bold: true, color: "333333" })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: HEADER_BG } },
  }),
  new Paragraph({
    spacing: { before: 80, after: 320 },
    children: [new TextRun({
      text: REFS.map(r => r.cat).join("　／　"),
      font: FONT, size: 18, color: "555555",
    })],
  })
);

// 各カテゴリ
REFS.forEach((section, si) => {
  const { bg, fg } = CAT_COLORS[section.cat] || CAT_COLORS["その他"];

  children.push(
    new Paragraph({
      spacing: { before: si === 0 ? 0 : 360, after: 120 },
      shading: { fill: bg, type: ShadingType.CLEAR },
      indent: { left: 120 },
      border: { left: { style: BorderStyle.SINGLE, size: 16, color: fg } },
      children: [
        new TextRun({ text: "■ " + section.cat, font: FONT, size: 26, bold: true, color: fg }),
      ],
    })
  );

  section.items.forEach((item, ii) => {
    // やりたいこと
    children.push(
      new Paragraph({
        spacing: { before: ii === 0 ? 80 : 200, after: 60 },
        children: [
          new TextRun({ text: "▶ ", font: FONT, size: 20, color: fg }),
          new TextRun({ text: item.want, font: FONT, size: 20, bold: true, color: "111111" }),
        ],
      })
    );

    // ラベル＋内容テーブル
    const rows = [];

    rows.push(new TableRow({
      children: [
        labelCell("構文", bg, fg),
        mkCell(codeBlock(item.syntax, "1e2433", "cdd6f4"), { w: VALUE }),
      ],
    }));

    if (item.note) {
      rows.push(new TableRow({
        children: [
          labelCell("説明", "f5f5f5", "777777"),
          mkCell([notePara(item.note)], { w: VALUE }),
        ],
      }));
    }

    rows.push(new TableRow({
      children: [
        labelCell("使用例", "edf5e8", "3a6020"),
        mkCell(codeBlock(item.example, "2d3548", "a6e3a1"), { w: VALUE }),
      ],
    }));

    children.push(
      new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [LABEL, VALUE],
        borders: gridBorder("cccccc"),
        rows,
      })
    );
  });
});

// フッター
children.push(
  new Paragraph({ spacing: { before: 400 } }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "VBA プログラミング基礎　学習サイト", font: FONT, size: 16, color: "aaaaaa" })],
  })
);

// ── 生成 ─────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: PW, height: 16838 },
        margin: { top: MAR, bottom: MAR, left: MAR, right: MAR },
      },
    },
    children,
  }],
});

const outDir = path.join(__dirname, "..", "docs", "downloads");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, "reference.docx");

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log("Generated:", outPath);
});
