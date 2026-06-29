// VBA ワークシート生成テンプレート
// 使用ライブラリ: docx (npm install docx)

const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, BorderStyle, ShadingType, AlignmentType,
  WidthType,
} = require("docx");
const fs = require("fs");

// 色定数
const COLOR = {
  HEADER_BG: "4a6fa5",
  HEADER_FG: "FFFFFF",
  BUG_BG: "FFEEEE",
  HOLE_BG: "FFFFEE",
  WHITE: "FFFFFF",
};

/**
 * バグ修正型ワークシートを生成する
 * @param {Object} config
 * @param {string} config.title          - ワークシートのタイトル
 * @param {string} config.scenario       - 状況説明（太郎さんが〜）
 * @param {string[]} config.specs        - プログラムの仕様（箇条書き）
 * @param {string} config.sampleDesc     - 例（n=?の場合の説明）
 * @param {string[][]} config.sampleTable - 例の表データ [["A1","7"],...]
 * @param {string[]} config.codeLines    - コード行（全行）
 * @param {number} config.bugLine        - バグ行番号（1始まり）
 * @param {string} config.question1      - 課題1の設問文
 * @param {string} config.question2      - 課題2の設問文
 * @param {string} config.question3      - 課題3の設問文
 * @param {string} config.outputPath     - 出力ファイルパス
 */
async function createBugFixWorksheet(config) {
  // 実装はここに追加
  // このテンプレートを参考に各ワークシートの生成スクリプトを作成する
  console.log(`生成: ${config.outputPath}`);
}

module.exports = { createBugFixWorksheet, COLOR };
