const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, BorderStyle, ShadingType, AlignmentType, WidthType,
} = require("docx");
const fs   = require("fs");
const path = require("path");

const C = { HEADER_BG:"4a6fa5", HEADER_FG:"FFFFFF", BUG_BG:"FFEEEE", SECTION_BG:"EEF2F8", WHITE:"FFFFFF" };
const FONT = "BIZ UDGothic", PAGE_W = 11906, MARGIN = 1000, CONTENT_W = PAGE_W - MARGIN * 2;

function border() { return { style:BorderStyle.SINGLE, size:1, color:"BBBBBB" }; }
function borders() { const b=border(); return { top:b, bottom:b, left:b, right:b }; }
function cell(text, opts={}) {
  const { bg=C.WHITE, bold=false, size=20, color="000000", w=null, align=AlignmentType.LEFT } = opts;
  return new TableCell({ width:w?{size:w,type:WidthType.DXA}:undefined, shading:{fill:bg,type:ShadingType.CLEAR}, borders:borders(), margins:{top:80,bottom:80,left:120,right:120},
    children:[new Paragraph({ alignment:align, children:[new TextRun({text,font:FONT,size,bold,color})] })] });
}
function hcell(text, w=null) { return cell(text,{bg:C.HEADER_BG,color:C.HEADER_FG,bold:true,size:20,w}); }
function para(text, opts={}) {
  const { bold=false,size=20,color="000000",spacing=120,indent=0 } = opts;
  return new Paragraph({ indent:indent?{left:indent}:undefined, spacing:{before:spacing,after:spacing}, children:[new TextRun({text,font:FONT,size,bold,color})] });
}
function sectionHeader(title) {
  return new Paragraph({ spacing:{before:200,after:100}, border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.HEADER_BG}}, shading:{fill:C.SECTION_BG,type:ShadingType.CLEAR},
    children:[new TextRun({text:"■ "+title,font:FONT,size:22,bold:true,color:C.HEADER_BG})] });
}
function emptyLines(n=1) { return Array.from({length:n},()=>new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:""})]})); }
function answerBox(lines=4) {
  const b=border(), bs={top:b,bottom:b,left:b,right:b};
  return new Table({ width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[CONTENT_W],
    rows:Array.from({length:lines},()=>new TableRow({ height:{value:400,rule:"exact"},
      children:[new TableCell({width:{size:CONTENT_W,type:WidthType.DXA},shading:{fill:C.WHITE,type:ShadingType.CLEAR},borders:bs,margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:""})]})] })] })) });
}
function titleBlock(t1,t2) {
  return new Table({ width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[Math.floor(CONTENT_W*0.6),Math.floor(CONTENT_W*0.4)],
    rows:[new TableRow({ children:[
      new TableCell({ shading:{fill:C.HEADER_BG,type:ShadingType.CLEAR}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, margins:{top:100,bottom:100,left:160,right:160},
        children:[new Paragraph({children:[new TextRun({text:t1,font:FONT,size:28,bold:true,color:C.HEADER_FG})]}),new Paragraph({children:[new TextRun({text:t2,font:FONT,size:22,color:"CCDDFF"})]})] }),
      new TableCell({ shading:{fill:C.HEADER_BG,type:ShadingType.CLEAR}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, margins:{top:100,bottom:100,left:160,right:160},
        children:[
          new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"クラス：________  番号：____",font:FONT,size:18,color:C.HEADER_FG})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"氏名：__________________",font:FONT,size:18,color:C.HEADER_FG})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"日付：____年____月____日",font:FONT,size:18,color:C.HEADER_FG})]}),
        ] }),
    ]})] });
}

// 2B: バグ行 = 行8 (If a Mod 2 = 0 → 1)
const CODE_LINES = [
  { no:"1",  code:"Sub DisplayOdd()" },
  { no:"2",  code:"    Dim n As Integer" },
  { no:"3",  code:"    Dim a As Integer" },
  { no:"4",  code:"    Dim row As Integer" },
  { no:"5",  code:'    n = InputBox("いくつまで調べますか？")' },
  { no:"6",  code:"    row = 1" },
  { no:"7",  code:"    For a = 1 To n" },
  { no:"8",  code:"        If a Mod 2 = 0 Then", bug:true },
  { no:"9",  code:"            Cells(row, 1).Value = a" },
  { no:"10", code:"            row = row + 1" },
  { no:"11", code:"        End If" },
  { no:"12", code:"    Next a" },
  { no:"13", code:"End Sub" },
];

function codeTable() {
  const colNO=600, colCode=CONTENT_W-colNO;
  return new Table({ width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[colNO,colCode],
    rows:[
      new TableRow({children:[hcell("行",colNO),hcell("コード",colCode)]}),
      ...CODE_LINES.map(({no,code,bug})=>new TableRow({children:[
        new TableCell({width:{size:colNO,type:WidthType.DXA},shading:{fill:bug?C.BUG_BG:C.WHITE,type:ShadingType.CLEAR},borders:borders(),margins:{top:60,bottom:60,left:100,right:100},
          children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:no,font:"Courier New",size:18})]})]}),
        new TableCell({width:{size:colCode,type:WidthType.DXA},shading:{fill:bug?C.BUG_BG:C.WHITE,type:ShadingType.CLEAR},borders:borders(),margins:{top:60,bottom:60,left:120,right:120},
          children:[new Paragraph({children:[new TextRun({text:code,font:"Courier New",size:18}),...(bug?[new TextRun({text:"  ← バグあり",font:FONT,size:18,color:"CC0000",bold:true})]:[])]})]})
      ]})),
    ] });
}

function specTable() {
  const col=Math.floor(CONTENT_W/2);
  return new Table({ width:{size:CONTENT_W,type:WidthType.DXA}, columnWidths:[col,col],
    rows:[
      new TableRow({children:[hcell("セル",col),hcell("表示値（奇数のみ）",col)]}),
      ...["A1:1","A2:3","A3:5","A4:7","A5:9"].map(s=>{ const [c,v]=s.split(":"); return new TableRow({children:[cell(c,{w:col}),cell(v,{w:col})]}); }),
    ] });
}

const doc = new Document({ sections:[{ properties:{ page:{size:{width:PAGE_W,height:16838},margin:{top:MARGIN,right:MARGIN,bottom:MARGIN,left:MARGIN}} },
  children:[
    titleBlock("26-06-15 条件判定 2B","奇数表示（If文・Mod演算子）"),
    ...emptyLines(1),
    sectionHeader("状況説明"),
    para("太郎さんは今度は「1 〜 n の中から奇数だけを A列に表示する」プログラムを作りました。", {indent:200}),
    para("しかし実行すると、奇数ではなく偶数が表示されてしまいます。", {indent:200}),
    para("2Aで学んだ Mod 演算子の知識を使って、バグを直してみましょう。", {indent:200,bold:true}),
    ...emptyLines(1),
    sectionHeader("プログラムの仕様"),
    para("①  ユーザーに「n（何まで調べるか）」を入力してもらう", {indent:200}),
    para("②  1 〜 n の中から奇数だけを A列に上から順に表示する", {indent:200}),
    para("③  例）n = 10 の場合 →", {indent:200}),
    ...emptyLines(1),
    specTable(),
    ...emptyLines(1),
    sectionHeader("バグのあるプログラム"),
    para("※ 薄赤の行にバグが仕込まれています。", {indent:200,color:"CC0000"}),
    ...emptyLines(1),
    codeTable(),
    ...emptyLines(1),
    sectionHeader("グループワーク"),
    para("【課題 1】　予想してみよう", {bold:true,size:22}),
    para("n = 10 で実行したとき、A列にはどんな値が入ると思いますか？", {indent:200}),
    ...emptyLines(1),
    answerBox(3),
    ...emptyLines(1),
    para("【課題 2】　バグを発見しよう", {bold:true,size:22}),
    para("2A のバグとどこが違いますか？　条件式の違いに注目して説明しましょう。", {indent:200}),
    ...emptyLines(1),
    answerBox(4),
    ...emptyLines(1),
    para("【課題 3】　修正しよう", {bold:true,size:22}),
    para("正しいコードに書き直して、Excel で動作を確認しましょう。", {indent:200}),
    ...emptyLines(1),
    answerBox(3),
  ],
}] });

const outPath = path.join(__dirname,"..","worksheets","step2","2b_odd-numbers.docx");
Packer.toBuffer(doc).then(buf=>{ fs.writeFileSync(outPath,buf); console.log("生成完了:", outPath); }).catch(e=>{ console.error(e); process.exit(1); });
