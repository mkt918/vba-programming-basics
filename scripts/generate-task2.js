const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, BorderStyle, ShadingType, AlignmentType, WidthType,
} = require("docx");
const fs   = require("fs");
const path = require("path");

const C = { HEADER_BG:"4a6fa5", HEADER_FG:"FFFFFF", HOLE_BG:"FFFFEE", SECTION_BG:"EEF2F8", WHITE:"FFFFFF" };
const FONT = "BIZ UDGothic", PAGE_W = 11906, MARGIN = 1000, CONTENT_W = PAGE_W - MARGIN * 2;

function border() { return { style:BorderStyle.SINGLE, size:1, color:"BBBBBB" }; }
function borders() { const b=border(); return {top:b,bottom:b,left:b,right:b}; }
function cell(text, opts={}) {
  const {bg=C.WHITE,bold=false,size=20,color="000000",w=null,align=AlignmentType.LEFT}=opts;
  return new TableCell({width:w?{size:w,type:WidthType.DXA}:undefined,shading:{fill:bg,type:ShadingType.CLEAR},borders:borders(),margins:{top:80,bottom:80,left:120,right:120},
    children:[new Paragraph({alignment:align,children:[new TextRun({text,font:FONT,size,bold,color})]})]});
}
function hcell(text,w=null){return cell(text,{bg:C.HEADER_BG,color:C.HEADER_FG,bold:true,size:20,w});}
function para(text,opts={}){
  const{bold=false,size=20,color="000000",spacing=120,indent=0}=opts;
  return new Paragraph({indent:indent?{left:indent}:undefined,spacing:{before:spacing,after:spacing},children:[new TextRun({text,font:FONT,size,bold,color})]});
}
function sectionHeader(title){
  return new Paragraph({spacing:{before:200,after:100},border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.HEADER_BG}},shading:{fill:C.SECTION_BG,type:ShadingType.CLEAR},
    children:[new TextRun({text:"■ "+title,font:FONT,size:22,bold:true,color:C.HEADER_BG})]});
}
function emptyLines(n=1){return Array.from({length:n},()=>new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:""})]}))}
function answerBox(lines=4){
  const b=border(),bs={top:b,bottom:b,left:b,right:b};
  return new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[CONTENT_W],
    rows:Array.from({length:lines},()=>new TableRow({height:{value:400,rule:"exact"},
      children:[new TableCell({width:{size:CONTENT_W,type:WidthType.DXA},shading:{fill:C.WHITE,type:ShadingType.CLEAR},borders:bs,margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:""})]})]})]}))});
}
function titleBlock(t1,t2){
  return new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[Math.floor(CONTENT_W*0.6),Math.floor(CONTENT_W*0.4)],
    rows:[new TableRow({children:[
      new TableCell({shading:{fill:C.HEADER_BG,type:ShadingType.CLEAR},borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:100,bottom:100,left:160,right:160},
        children:[new Paragraph({children:[new TextRun({text:t1,font:FONT,size:28,bold:true,color:C.HEADER_FG})]}),new Paragraph({children:[new TextRun({text:t2,font:FONT,size:22,color:"CCDDFF"})]})]},),
      new TableCell({shading:{fill:C.HEADER_BG,type:ShadingType.CLEAR},borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:100,bottom:100,left:160,right:160},
        children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"クラス：________  番号：____",font:FONT,size:18,color:C.HEADER_FG})]}),new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"氏名：__________________",font:FONT,size:18,color:C.HEADER_FG})]}),new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"日付：____年____月____日",font:FONT,size:18,color:C.HEADER_FG})]})]}),
    ]})]});
}

// 課題② 穴埋めコード (if~else版、バグなし)
const CODE_LINES=[
  {no:"1", code:"Sub ClassifyNumbers()"},
  {no:"2", code:"    Dim n As Integer"},
  {no:"3", code:"    Dim a As Integer"},
  {no:"4", code:"    Dim rowA As Integer"},
  {no:"5", code:"    Dim rowB As Integer"},
  {no:"6", code:'    n = InputBox("いくつまで調べますか？")'},
  {no:"7", code:"    rowA = 1 : rowB = 1"},
  {no:"8", code:"    For a = 1 To n"},
  {no:"9", code:"        If a Mod 2 = [   ] Then", hole:true},
  {no:"10",code:"            Cells(rowA, 1).Value = [   ]", hole:true},
  {no:"11",code:"            rowA = rowA + 1"},
  {no:"12",code:"        Else"},
  {no:"13",code:"            Cells(rowB, 2).Value = [   ]", hole:true},
  {no:"14",code:"            rowB = rowB + 1"},
  {no:"15",code:"        End If"},
  {no:"16",code:"    Next a"},
  {no:"17",code:"End Sub"},
];

function codeTable(){
  const colNO=600,colCode=CONTENT_W-colNO;
  return new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[colNO,colCode],
    rows:[
      new TableRow({children:[hcell("行",colNO),hcell("コード（[ ] に入るものを書こう）",colCode)]}),
      ...CODE_LINES.map(({no,code,hole})=>new TableRow({height:{value:420,rule:"exact"},children:[
        new TableCell({width:{size:colNO,type:WidthType.DXA},shading:{fill:hole?C.HOLE_BG:C.WHITE,type:ShadingType.CLEAR},borders:borders(),margins:{top:60,bottom:60,left:100,right:100},
          children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:no,font:"Courier New",size:18})]})]}),
        new TableCell({width:{size:colCode,type:WidthType.DXA},shading:{fill:hole?C.HOLE_BG:C.WHITE,type:ShadingType.CLEAR},borders:borders(),margins:{top:60,bottom:60,left:120,right:120},
          children:[new Paragraph({children:[new TextRun({text:code,font:"Courier New",size:18})]})]}),
      ]})),
    ]});
}

function specTable(){
  const col=Math.floor(CONTENT_W/2);
  return new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[col,col],
    rows:[
      new TableRow({children:[hcell("A列（偶数）",col),hcell("B列（奇数）",col)]}),
      ...[[2,1],[4,3],[6,5],[8,7],[10,9]].map(([a,b])=>new TableRow({children:[cell(String(a),{w:col}),cell(String(b),{w:col})]})),
    ]});
}

const doc=new Document({sections:[{properties:{page:{size:{width:PAGE_W,height:16838},margin:{top:MARGIN,right:MARGIN,bottom:MARGIN,left:MARGIN}}},
  children:[
    titleBlock("26-06-15 応用課題②","分類表示（If〜Else）"),
    ...emptyLines(1),
    sectionHeader("状況説明"),
    para("STEP 2 の仕上げとして、偶数と奇数を別々の列に振り分けるプログラムを完成させましょう。", {indent:200}),
    para("If〜Else を使って、条件によって処理を分岐させる方法を練習します。", {indent:200,bold:true}),
    ...emptyLines(1),
    sectionHeader("プログラムの仕様"),
    para("①  ユーザーに「n（何まで調べるか）」を入力してもらう", {indent:200}),
    para("②  偶数は A列（rowA 行目）に表示し、rowA を 1 増やす", {indent:200}),
    para("③  奇数は B列（rowB 行目）に表示し、rowB を 1 増やす", {indent:200}),
    para("④  例）n = 10 の場合 →", {indent:200}),
    ...emptyLines(1),
    specTable(),
    ...emptyLines(1),
    sectionHeader("課題 1　日本語で処理の流れを書こう"),
    para("a = 1, 2, 3 のときそれぞれどの列に入るか、日本語で説明してみましょう。", {indent:200}),
    ...emptyLines(1),
    answerBox(4),
    ...emptyLines(1),
    sectionHeader("課題 2　コードの穴埋めをしよう"),
    para("薄黄の行の [ ] に当てはまるコードを書いて完成させましょう。", {indent:200}),
    para("完成したら Excel に入力して動作を確認しましょう。", {indent:200}),
    ...emptyLines(1),
    codeTable(),
  ],
}]});

const outPath=path.join(__dirname,"..","worksheets","step2","task2_classification.docx");
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync(outPath,buf);console.log("生成完了:",outPath);}).catch(e=>{console.error(e);process.exit(1);});
