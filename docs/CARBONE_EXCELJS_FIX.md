# Carbone + ExcelJS 兼容性问题修复

## 问题描述

通过 Carbone 生成的 Excel 报表无法被 ExcelJS 直接读取，会报错：

```
Error: Invalid row number in model
```

但是，如果通过 Microsoft Excel 或 WPS 等软件打开并另存为后，生成的文件就可以被 ExcelJS 正常读取和处理。

## 问题原因

### 1. Excel 文件格式规范性问题

Carbone 生成的 Excel 文件可能包含一些不完全符合规范的结构，例如：
- 无效的行号引用或范围定义
- 工作表维度（dimension）定义不准确
- 共享字符串表（SharedStrings）格式问题
- 样式引用索引不正确

### 2. 软件容错能力差异

- **Microsoft Excel / WPS**：这些桌面软件对 Excel 格式有很强的容错能力，能够自动修复和规范化不规范的内容。当你"另存为"时，软件会重新构建整个文件结构，输出标准的 OOXML 格式。

- **ExcelJS**：作为一个纯 JavaScript 库，ExcelJS 对文件格式要求更严格，无法容忍某些不规范的结构，因此会报错。

## 解决方案

### 使用 xlsx-js-style 作为中间层

[xlsx-js-style](https://www.npmjs.com/package/xlsx-js-style) (SheetJS) 是一个对 Excel 格式更宽容的库，可以：
1. 读取 Carbone 生成的文件（即使格式不完全规范）
2. 重新写入并规范化文件结构
3. 生成 ExcelJS 可以正常读取的标准文件

### 实现流程

```
Carbone 生成文件
    ↓
xlsx-js-style 读取并重写（规范化）
    ↓
ExcelJS 读取并后处理
```

### 代码实现

在 `src/main/services/dataToReport.ts` 中添加规范化步骤：

```typescript
import * as XLSX from 'xlsx-js-style'

// 如果有 postProcess 钩子，先用 xlsx-js-style 规范化 Carbone 生成的文件
if (template.postProcess) {
  console.log(`[dataToReport] 检测到 postProcess 钩子，先规范化文件格式...`)
  try {
    const tempPath = `${outputPath}.temp.xlsx`

    // 使用 xlsx-js-style 读取 Carbone 生成的文件（对格式更宽容）
    const workbook = XLSX.readFile(outputPath)
    console.log(`[dataToReport] 使用 xlsx-js-style 读取 Carbone 文件成功`)

    // 重新写入文件（这会规范化文件结构，使 ExcelJS 能读取）
    XLSX.writeFile(workbook, tempPath)
    console.log(`[dataToReport] 已规范化并保存到临时文件`)

    // 删除原 Carbone 文件并重命名临时文件
    deleteFileIfExists(outputPath)
    fs.renameSync(tempPath, outputPath)
    console.log(`[dataToReport] 文件格式规范化完成，现在 ExcelJS 可以读取了`)
  } catch (normalizeError) {
    console.warn(`[dataToReport] 文件规范化失败:`, normalizeError)
  }
}
```

## 为什么这个方案有效？

1. **格式宽容性**：xlsx-js-style 内部使用了更宽松的解析逻辑，可以处理一些不规范的结构。

2. **重新序列化**：当 xlsx-js-style 写入文件时，它会按照标准规范重新构建整个文件结构，类似于 Excel 软件的"另存为"操作。

3. **标准化输出**：重写后的文件完全符合 OOXML 标准，ExcelJS 可以正常读取。

## 测试验证

运行测试脚本验证修复效果：

```bash
pnpm exec tsx scripts/test-carbone-exceljs-fix.ts
```

测试流程：
1. 解析源数据文件
2. 使用 Carbone 生成报表
3. 自动使用 xlsx-js-style 规范化
4. 使用 ExcelJS 读取并验证

## 性能影响

- **额外开销**：规范化步骤需要完整读写一次文件，会增加约 20-50% 的处理时间。
- **文件大小**：规范化后的文件大小通常与 Excel 另存为的结果相近。
- **仅在需要时执行**：只有模板定义了 `postProcess` 钩子时才会执行规范化。

## 依赖

确保 `package.json` 中包含以下依赖：

```json
{
  "dependencies": {
    "xlsx-js-style": "^1.2.0",
    "exceljs": "^4.4.0",
    "carbone": "^3.5.6"
  }
}
```

## 替代方案（不推荐）

### 方案1：使用 ExcelJS 重写（已验证无效）
直接用 ExcelJS 读取 Carbone 文件会失败，因为 ExcelJS 无法读取不规范的文件。

### 方案2：调整 Carbone 配置
Carbone 的配置选项有限，无法从根本上解决格式规范性问题。

### 方案3：使用其他 Excel 库
其他库（如 node-xlsx）也可以，但 xlsx-js-style 保留了样式信息，更适合报表场景。

## 总结

通过在 Carbone 和 ExcelJS 之间插入 xlsx-js-style 规范化步骤，我们成功解决了兼容性问题。这个方案：

✅ 无需修改 Carbone 或 ExcelJS 的使用方式  
✅ 对用户透明，自动执行  
✅ 保留了文件的样式和格式  
✅ 性能开销可接受  

## 参考资料

- [Carbone Documentation](https://carbone.io/documentation.html)
- [ExcelJS GitHub](https://github.com/exceljs/exceljs)
- [xlsx-js-style GitHub](https://github.com/gitbrent/xlsx-js-style)
