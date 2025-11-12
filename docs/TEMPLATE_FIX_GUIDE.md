# Carbone 模板修复指南 - month2carbone

## 当前问题

检查模板发现以下问题：

### 模板当前结构（5行）
- 第1行：标题行（行业、当月新增放款）
- 第2行：月份标题（1月、2月...）
- 第3行：`{d.rows[i].industry}` + `{d.rows[i].m1}` ... `{d.rows[i].m12}`
- 第4行：`{d.rows[i+1].industry}` + **空数据列**（问题所在）
- 第5行：合计行

### 问题分析

第4行使用了 `{d.rows[i+1].industry}`，但后续的数据列都是空的，导致只有2个行业能被正确渲染。

## 解决方案

使用 Carbone 的循环语法来动态生成所有行业数据。

### 推荐的模板结构

```
第1行：行业 | 当月新增放款（万元）
第2行：    | 1月 | 2月 | 3月 | ... | 12月
第3行：{d.rows[i].industry :hideIf(i>0)} | {d.rows[i].m1} | {d.rows[i].m2} | ... | {d.rows[i].m12}
第4行：合计 | {d.total.m1} | {d.total.m2} | ... | {d.total.m12}
```

### Carbone 循环语法说明

在第3行使用：
- `{d.rows[i].industry :hideIf(i>0)}` - 循环遍历 rows 数组
- `hideIf(i>0)` 确保只有第一个元素时显示，后续行会隐藏第一列（可选）
- 或者直接使用 `{d.rows[i].industry}` 让每行都显示行业名

### 修改步骤

1. 在 Excel 中打开 `public/reportTemplates/month2carbone.xlsx`
2. 删除当前的第4行（有问题的行）
3. 在第3行使用正确的循环语法：
   - A3: `{d.rows[i].industry}`
   - B3: `{d.rows[i].m1}`
   - C3: `{d.rows[i].m2}`
   - ... 依此类推到 M3: `{d.rows[i].m12}`
4. 第4行改为合计行：
   - A4: `合计`
   - B4: `{d.total.m1}`
   - C4: `{d.total.m2}`
   - ... 依此类推到 M4: `{d.total.m12}`

### 修改后的结构（4行）

- 第1行：标题行
- 第2行：月份标题
- 第3行：行业数据（Carbone 会自动循环生成3行）
- 第4行：合计行

## 验证

修改后，Carbone 渲染结果应该有 **6行**：
- 第1行：标题
- 第2行：月份
- 第3行：基建工程（数据）
- 第4行：医药医疗（数据）
- 第5行：再保理（数据）
- 第6行：合计

## PostProcess 代码已适配

`month2Post.ts` 中的代码已修改为动态适配行数：
```typescript
const lastDataRow = sheet.rowCount - 1 // 最后一行是合计行
for (let rowNum = 3; rowNum <= lastDataRow; rowNum++) {
  // 应用数据行样式
}
const totalRow = sheet.getRow(sheet.rowCount) // 最后一行
// 应用合计行样式
```

## 参考资料

- [Carbone Documentation - Arrays and Loops](https://carbone.io/documentation.html#arrays-iteration)
- [Carbone Filters](https://carbone.io/documentation.html#formatters)
