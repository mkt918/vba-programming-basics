# Bug Catalog - VBA Learning Materials

## STEP 1-A: 同値繰返し

- **Bug Line**: 6
- **Bug Code**: `Cells(a, 1).Value = a`
- **Correct Code**: `Cells(a, 1).Value = n`
- **Concept Tested**: ループ変数 vs 固定値の違い
- **Worksheet**: [worksheets/step1/1a_same-value-repeat.docx](../worksheets/step1/1a_same-value-repeat.docx)
- **Solution**: [solutions/step1-solutions.md#step1a](../solutions/step1-solutions.md)

---

## STEP 1-B: 連番表示

- **Bug Line**: 6
- **Bug Code**: `Cells(1, 1).Value = a`
- **Correct Code**: `Cells(a, 1).Value = a`
- **Concept Tested**: ループ変数を行番号として使う
- **Worksheet**: [worksheets/step1/1b_sequence-display.docx](../worksheets/step1/1b_sequence-display.docx)
- **Solution**: [solutions/step1-solutions.md#step1b](../solutions/step1-solutions.md)

---

## STEP 2-A: 偶数表示

- **Bug Line**: 7
- **Bug Code**: `If a Mod 2 = 1 Then`
- **Correct Code**: `If a Mod 2 = 0 Then`
- **Concept Tested**: 偶数判定（Mod演算子の余り = 0）
- **Worksheet**: [worksheets/step2/2a_even-numbers.docx](../worksheets/step2/2a_even-numbers.docx)
- **Solution**: [solutions/step2-solutions.md#step2a](../solutions/step2-solutions.md)

---

## STEP 2-B: 奇数表示

- **Bug Line**: 7
- **Bug Code**: `If a Mod 2 = 0 Then`
- **Correct Code**: `If a Mod 2 = 1 Then`
- **Concept Tested**: 条件式を反転させる（偶数↔奇数）
- **Worksheet**: [worksheets/step2/2b_odd-numbers.docx](../worksheets/step2/2b_odd-numbers.docx)
- **Solution**: [solutions/step2-solutions.md#step2b](../solutions/step2-solutions.md)
