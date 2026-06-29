# Claude Code 移転マニュアル

VBA プログラミング基礎の全教材を Claude Code での管理・保守に移転するための手順書

---

## 📁 ディレクトリ構成案

```
vba-programming-basics/
├── README.md                    # プロジェクト概要
├── MANIFEST.md                  # 全ファイル一覧（このドキュメント）
├── design/                      # 計画・設計ドキュメント
│   ├── learning-flow.md         # 学習設計書（MD版）
│   ├── assessment.md            # 学習目標・評価規準
│   └── bugs-catalog.md          # バグ内容の一覧表
├── worksheets/                  # 学習ワークシート（Word）
│   ├── step1/
│   │   ├── 1a_same-value-repeat.docx
│   │   ├── 1b_sequence-display.docx
│   │   └── task1_multi-column.docx
│   ├── step2/
│   │   ├── 2a_even-numbers.docx
│   │   ├── 2b_odd-numbers.docx
│   │   └── task2_classification.docx
│   └── step3-4/
│       ├── 3_fizzbuzz-prep.docx
│       └── 4_fizzbuzz-complete.docx
├── templates/                   # テンプレート・スクリプト
│   ├── worksheet-template.js    # Docx生成テンプレート
│   ├── bugfix-template.vba      # バグ修正版プログラムテンプレート
│   └── correct-template.vba     # 正解プログラムテンプレート
├── solutions/                   # 解答・解説（教員用）
│   ├── step1-solutions.md
│   ├── step2-solutions.md
│   └── trace-tables/
│       ├── 1a-trace.xlsx
│       ├── 1b-trace.xlsx
│       ├── 2a-trace.xlsx
│       └── 2b-trace.xlsx
└── docs/                        # その他ドキュメント
    ├── classroom-usage.md       # 授業活用ガイド
    ├── troubleshooting.md       # トラブルシューティング
    └── version-history.md       # バージョン履歴
```

---

## 🔄 ファイル移転チェックリスト

### ドキュメント（MD形式で移転）

- [ ] `VBA1重ループ学習設計書.md` → `design/learning-flow.md`
- [ ] `学習目標・評価規準シート` → `design/assessment.md`
- [ ] `全教材ファイル一覧.md` → `MANIFEST.md`（編集）
- [ ] バグ内容一覧表（新規作成） → `design/bugs-catalog.md`

### ワークシート（DOCX のまま保持）

**STEP 1**
- [ ] `ループ基礎1A_同値繰返し.docx` → `worksheets/step1/1a_same-value-repeat.docx`
- [ ] `ループ基礎1B_連番表示.docx` → `worksheets/step1/1b_sequence-display.docx`
- [ ] `ループ応用課題①_複数列出力.docx` → `worksheets/step1/task1_multi-column.docx`

**STEP 2**
- [ ] `条件判定2A_偶数表示.docx` → `worksheets/step2/2a_even-numbers.docx`
- [ ] `条件判定2B_奇数表示.docx` → `worksheets/step2/2b_odd-numbers.docx`
- [ ] `条件判定応用課題②_分類表示.docx` → `worksheets/step2/task2_classification.docx`（次回作成）

### スクリプト（新規作成）

- [ ] `worksheet-template.js` → Node.js docx生成用テンプレート
- [ ] `bugfix-template.vba` → バグ修正版プログラム
- [ ] `correct-template.vba` → 正解プログラム

### 解答・解説（新規作成）

- [ ] `step1-solutions.md` → 各STEPの解答説明
- [ ] `step2-solutions.md` → 条件判定の解答説明
- [ ] トレース表（XLSX）5枚

---

## 📋 各ファイルの役割と更新ポリシー

### `MANIFEST.md` (全体一覧)

**役割**：プロジェクトの完全なインデックス

```markdown
# VBA Programming Basics - File Manifest
- Last Updated: 2026-06-15
- Version: 1.0.0
- Status: In Development

## Quick Links
- [Learning Flow Design](design/learning-flow.md)
- [Assessment Criteria](design/assessment.md)
- [Bug Catalog](design/bugs-catalog.md)

## File Structure
...
```

**更新ポリシー**
- 新しいファイル追加時に即座に更新
- 月1回の整合性チェック
- Git のコミットメッセージに参照記載

---

### `design/learning-flow.md` (学習設計書)

**役割**：教材開発の基本方針・各STEPの目標

```markdown
# VBA 1重ループ学習設計書

## 全体目標
...

## STEP 1-A: 同値繰返し
- 仕様：...
- バグ：...（→design/bugs-catalog.md#step1aを参照）
- 教材：worksheets/step1/1a_same-value-repeat.docx
- 解答：solutions/step1-solutions.md#step1a
```

**更新ポリシー**
- 新STEPの設計時に更新
- 教材ファイル名の変更があれば即座に同期

---

### `design/bugs-catalog.md` (バグ一覧)

**役割**：全バグの集約管理・参考資料

```markdown
# Bug Catalog for VBA Learning Materials

## STEP 1-A: 同値繰返し
- **Bug Line**: 6
- **Bug Code**: `Cells(a, 1).Value = a`
- **Correct Code**: `Cells(a, 1).Value = n`
- **Concept Tested**: ループ変数 vs 固定値の違い
- **Worksheet**: worksheets/step1/1a_same-value-repeat.docx
- **Solution**: solutions/step1-solutions.md#step1a

## STEP 1-B: 連番表示
...

## STEP 2-A: 偶数表示
...
```

**更新ポリシー**
- 新バグを仕込むたびに追加
- 学習効果の検証結果を「Note」に記載

---

### `templates/worksheet-template.js`

**役割**：Word ワークシート生成の再利用可能テンプレート

```javascript
// 使用例
const createBugFixWorksheet = ({
  title: "26-06-15 ループ基礎 1A：同値繰返し",
  bugLine: "6",
  bugContent: "Cells(a, 1).Value = a",
  correctContent: "Cells(a, 1).Value = n",
  sampleInput: "「数値」と「回数」をユーザーに入力してもらう",
  sampleOutput: "入力した数値を、指定した回数分だけ A列 に表示する",
  question1: "数値＝7、回数＝5 を入力した場合、A1〜A5 にはどんな値が入ると思いますか？",
  question2: "下の表を見て、現在のコードで実行したら何が起こるか説明してください。"
});
```

**更新ポリシー**
- コード改善時に即座にテンプレートに反映
- Git で変更履歴を管理

---

### `solutions/step1-solutions.md`

**役割**：解答・解説（教員用）

```markdown
# STEP 1 Solutions

## STEP 1-A: 同値繰返し

### 【課題1】予想
**正解例**：
- A1: 7
- A2: 7
- A3: 7
- A4: 7
- A5: 7

### 【課題2】バグ発見
**正解**：
- 間違っている行：6番目
- 理由：`Cells(a, 1) = a` とすると、ループ変数aの値（1,2,3,4,5）が入力される。
  しかし仕様では「入力した数値nを繰り返す」なので、`Cells(a, 1) = n` にすべき。

### 【課題3】修正
**正解コード**：
```vba
Cells(a, 1).Value = n
```

### 教育的ポイント
...
```

**更新ポリシー**
- 生徒の回答傾向に基づいて「よくある間違い」セクションを追加
- 半年ごとに難易度調整の必要性を検討

---

## 🔗 ファイル間の相互参照ルール

```
README.md (入口)
  ↓
  ├→ MANIFEST.md (全体索引)
  │   ├→ design/learning-flow.md
  │   ├→ design/assessment.md
  │   ├→ design/bugs-catalog.md
  │   └→ worksheets/step1/...
  │
  ├→ docs/classroom-usage.md (授業活用)
  │   ├→ worksheets/... (教材ファイル)
  │   └→ solutions/... (解答)
  │
  └→ docs/troubleshooting.md (よくある質問)
```

**参照記法の統一**
```markdown
- 同一フォルダ内：`[ファイル名](file.md)`
- 異なるフォルダ：`[説明](../folder/file.md)`
- 外部リンク：`[説明](https://...)`
- ファイル内アンカー：`[セクション](#section-name)`
```

---

## 🔐 バージョン管理戦略

### セマンティック バージョニング

```
X.Y.Z
│ │ └─ パッチ（バグ修正・軽微な改善）
│ └─── マイナー（新STEP追加・新教材）
└───── メジャー（カリキュラム大改訂）

例：
- 1.0.0 → 初版リリース（STEP 1-A/B + 課題① + STEP 2-A/B）
- 1.1.0 → 課題② 追加、FizzBuzz STEP 3-4 追加
- 1.1.1 → STEP 1-A のバグ記述を修正
- 2.0.0 → 2重ループ（99表）導入、カリキュラム大刷新
```

### Changelog 記載項目

```markdown
## [1.1.0] - 2026-07-15

### Added
- 【課題②】分類表示ワークシート
- STEP 3: FizzBuzz 準備版
- STEP 4: FizzBuzz 完全版
- 解答・解説版（教員用）

### Changed
- STEP 1-A の難易度判定を ★☆☆ → ★★☆ に改訂
- design/bugs-catalog.md に「学習効果検証結果」セクション追加

### Fixed
- STEP 2-A ワークシートの例図が誤っていた問題を修正
```

---

## 📦 Claude Code での保管・保守

### フォルダ構成

```
vba-prog-basics/
  ├─ README.md                （プロジェクト入口）
  ├─ MANIFEST.md              （ファイル一覧・リンク集）
  ├─ .gitignore               （生成ファイル除外）
  ├─ design/
  ├─ worksheets/
  ├─ templates/
  ├─ solutions/
  ├─ docs/
  └─ scripts/                 （生成・検証用スクリプト）
      ├─ generate-worksheets.js
      ├─ validate-all.py
      └─ sync-manifest.js
```

### 必須ファイル

| ファイル | 用途 | 更新頻度 |
|---|---|---|
| `README.md` | プロジェクト説明 | 半年ごと |
| `MANIFEST.md` | 全ファイル索引 | 毎回（ファイル追加時） |
| `CHANGELOG.md` | 更新履歴 | 毎回（更新時） |
| `.gitignore` | Git 除外設定 | 初回のみ |
| `scripts/sync-manifest.js` | マニフェスト自動生成 | 自動実行 |

---

## 🚀 Claude Code での作業フロー

### 新しいSTEPを追加する場合

1. **設計フェーズ**
   - `design/learning-flow.md` に仕様を追記
   - `design/assessment.md` に学習目標を追記
   - `design/bugs-catalog.md` にバグを登録

2. **教材作成フェーズ**
   - `templates/worksheet-template.js` を参考に生成スクリプト作成
   - Word ワークシート（DOCX）を `worksheets/stepX/` に配置

3. **解答作成フェーズ**
   - `solutions/stepX-solutions.md` に解答・解説を記述
   - トレース表（XLSX）を `solutions/trace-tables/` に配置

4. **同期フェーズ**
   - `MANIFEST.md` を自動生成・更新
   - `CHANGELOG.md` に追記
   - Git コミット

---

## 💾 移転作業のチェックリスト（実行手順）

### Phase 1: 情報整理

- [ ] 全ファイルの内容を `.md` で文書化
- [ ] バグ一覧を `bugs-catalog.md` に統約
- [ ] 相互参照マッピング作成

### Phase 2: ディレクトリ構成構築

- [ ] Claude Code でプロジェクトを新規作成
- [ ] ディレクトリ構造を作成
- [ ] `.gitignore` を配置

### Phase 3: ファイル移転

- [ ] `design/` 内の MD ファイルを配置
- [ ] `worksheets/` 内の DOCX ファイルを配置（相対パス）
- [ ] `templates/` 内のスクリプトを配置
- [ ] `solutions/` 内の解答・解説を配置

### Phase 4: 同期・検証

- [ ] `MANIFEST.md` を生成・検証
- [ ] `README.md` を作成
- [ ] 全ファイルのリンク整合性をチェック
- [ ] Git で初期コミット

### Phase 5: 運用開始

- [ ] メンテナンス担当者にアクセス権付与
- [ ] `CHANGELOG.md` を初期化
- [ ] 最初のバージョン（1.0.0）をタグ付け

---

## 📞 運用・保守に関する注意事項

### 定期メンテナンス（月1回）

```bash
# MANIFEST の整合性チェック
node scripts/sync-manifest.js --verify

# 全ワークシートの検証
python scripts/validate-all.py

# リンク切れチェック
node scripts/check-links.js
```

### バージョン更新の決定基準

| 変更内容 | バージョン | 例 |
|---|---|---|
| バグ修正・誤字修正 | パッチ（Z） | 1.0.0 → 1.0.1 |
| 新STEP・新教材追加 | マイナー（Y） | 1.0.0 → 1.1.0 |
| カリキュラム大改訂 | メジャー（X） | 1.0.0 → 2.0.0 |

### 緊急修正対応

- バグ発見時の報告書テンプレート
- ホットフィックスブランチの運用ルール
- 緊急時の連絡フロー

---

## ✅ 移転完了のサイン

チェック項目がすべて ✓ になったら移転完了です：

- [ ] 全ファイルが `vba-prog-basics/` 下に配置されている
- [ ] `README.md` が完成し、プロジェクト概要が明確
- [ ] `MANIFEST.md` が自動生成され、全ファイルがリンク済み
- [ ] `.gitignore` が設定され、生成ファイルが除外されている
- [ ] `CHANGELOG.md` にバージョン履歴が記載されている
- [ ] すべての内部リンク（相対パス）が動作確認済み
- [ ] `scripts/` 内のメンテナンス用スクリプトが動作確認済み
- [ ] Git で初期コミット完了、タグ `v1.0.0` が付与されている

---

