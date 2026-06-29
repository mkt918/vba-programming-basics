# VBA 1重ループ 学習設計書

## 全体目標

For〜Next ループと If 文を組み合わせ、セルへの出力・条件判定ができる。

## 難易度マップ

```
STEP 1-A（同値繰返し）★☆☆
  └ ループの仕組み（変数aを使わない）

STEP 1-B（連番表示）★★☆
  └ ループ変数aを値として使う

【課題①】複数列出力 ★★☆
  └ 複数処理の組み合わせ、ループ内での計算

STEP 2-A（偶数表示）★★★
  └ if文 + Mod演算子

STEP 2-B（奇数表示）★★★
  └ 条件式反転

【課題②】分類表示 ★★★
  └ if〜else + 複数列出力
```

**原則**：各ステップで「新しい概念は1つだけ」追加

---

## STEP 1-A: 同値繰返し

- **仕様**：固定値 `n` を指定回数分だけ A列に表示
- **バグ**：[bugs-catalog.md#step1a](bugs-catalog.md)
- **教材**：[worksheets/step1/1a_same-value-repeat.docx](../worksheets/step1/1a_same-value-repeat.docx)
- **解答**：[solutions/step1-solutions.md](../solutions/step1-solutions.md)

---

## STEP 1-B: 連番表示

- **仕様**：1〜n を順に A列に表示
- **バグ**：[bugs-catalog.md#step1b](bugs-catalog.md)
- **教材**：[worksheets/step1/1b_sequence-display.docx](../worksheets/step1/1b_sequence-display.docx)
- **解答**：[solutions/step1-solutions.md](../solutions/step1-solutions.md)

---

## 応用課題①: 複数列出力

- **仕様**：A列に1〜n、B列に各行の2倍を表示
- **形式**：日本語手順書き版（バグなし）
- **教材**：[worksheets/step1/task1_multi-column.docx](../worksheets/step1/task1_multi-column.docx)

---

## STEP 2-A: 偶数表示

- **仕様**：1〜n の偶数だけ A列に表示
- **バグ**：[bugs-catalog.md#step2a](bugs-catalog.md)
- **教材**：[worksheets/step2/2a_even-numbers.docx](../worksheets/step2/2a_even-numbers.docx)
- **解答**：[solutions/step2-solutions.md](../solutions/step2-solutions.md)

---

## STEP 2-B: 奇数表示

- **仕様**：1〜n の奇数だけ A列に表示
- **バグ**：[bugs-catalog.md#step2b](bugs-catalog.md)
- **教材**：[worksheets/step2/2b_odd-numbers.docx](../worksheets/step2/2b_odd-numbers.docx)
- **解答**：[solutions/step2-solutions.md](../solutions/step2-solutions.md)

---

## 応用課題②: 分類表示（次作成）

- **仕様**：偶数は A列、奇数は B列に表示（if〜else）
- **教材**：[worksheets/step2/task2_classification.docx](../worksheets/step2/task2_classification.docx)
