# 测试指南

## 命令行测试报表生成

### 准备测试数据

创建一个测试用的 Excel 文件（例如 `test-data.xlsx`），包含以下结构：

**Sheet1 (默认表)**
| 姓名 | 部门 | 销售额 | 日期 |
|------|------|--------|------------|
| 张三 | 销售部 | 10000 | 2024-01-01 |
| 李四 | 技术部 | 15000 | 2024-01-02 |
| 王五 | 销售部 | 12000 | 2024-01-03 |

或者任何包含表头和数据行的标准表格结构。

### 运行测试

1. **准备测试文件**
   - 将测试 Excel 文件放在项目根目录（或任意位置）
   - 创建输出目录（如 `./output`）

2. **执行测试命令**

```bash
# 基本用法
pnpm test:report --templateId=month1carbone --source=./test-data.xlsx --outDir=./output

# 指定自定义报表名称
pnpm test:report --templateId=month1carbone --source=./test-data.xlsx --outDir=./output --reportName=我的报表.xlsx
```

3. **查看结果**
   - 脚本会显示详细的执行日志
   - 成功后会输出报表文件路径和大小
   - 可以使用提示的命令打开输出文件夹

### 示例输出

```
=== 报表生成测试 ===

📋 模板 ID: month1carbone
📂 源文件: ./test-data.xlsx
📁 输出目录: ./output

🔧 初始化模板系统...
[Templates] 开始初始化模板系统...
[Registry] 模板已注册: month1carbone - 月度报表模板
[Templates] 模板系统初始化完成

📊 解析 Excel 数据...
[excelToData] 开始处理: D:\...\test-data.xlsx with template: month1carbone
[excelToData] 模板已加载: 月度报表模板
[excelToData] Workbook 已加载，共 1 个 sheet
[excelToData] 解析完成
✅ 解析完成，耗时: 156ms
   - 数据源: D:\...\test-data.xlsx
   - 文件大小: 8.45 KB
   - 工作表: Sheet1

📄 生成报表...
[dataToReport] 开始生成报表: month1carbone
[dataToReport] 模板已加载: 月度报表模板
[dataToReport] 报表数据已构建
[dataToReport] 模板文件路径: D:\...\public\reportTemplates\month1carbone.xlsx
[dataToReport] Carbone 渲染完成，大小: 10540 字节
[dataToReport] 输出路径: D:\...\output\month1carbone-20241111-150325.xlsx
[dataToReport] 报表已写入
✅ 报表生成完成，耗时: 234ms
   - 输出路径: D:\...\output\month1carbone-20241111-150325.xlsx
   - 文件大小: 10.29 KB
   - 生成时间: 2024-11-11T07:03:25.123Z

🎉 测试完成！
总耗时: 390ms

💡 提示: 可以使用以下命令打开输出文件夹:
   explorer "D:\...\output"
```

### 错误处理示例

如果遇到错误，会显示详细的错误信息：

```
❌ 生成失败:
   错误: 模板未找到: month1carbone
   错误码: TEMPLATE_NOT_FOUND
   详情: {
     "templateId": "month1carbone"
   }
```

常见错误码：
- `TEMPLATE_NOT_FOUND` - 模板不存在或未注册
- `EXCEL_UNSUPPORTED_FILE` - 文件不支持或不存在
- `EXCEL_FILE_TOO_LARGE` - 文件超过 100MB 限制
- `EXCEL_PARSE_ERROR` - Excel 解析失败
- `REPORT_RENDER_ERROR` - Carbone 渲染失败
- `OUTPUT_DIR_NOT_SELECTED` - 输出目录未指定
- `OUTPUT_WRITE_ERROR` - 写入文件失败

### 多表聚合测试

如果源 Excel 包含多个 Sheet，可以通过模板的 parseOptions 指定：

```typescript
// 当前示例模板默认读取首表
// 要支持多表聚合，需要在模板解析器中实现自定义逻辑
// 参考 src/main/services/templates/month1carbone.ts 中的 parseOptions 接口
```

### 自定义模板测试

1. 在 `public/reportTemplates/` 添加新模板文件
2. 在 `src/main/services/templates/` 创建对应的模板模块
3. 在 `src/main/services/templates/index.ts` 中注册
4. 使用新的 templateId 运行测试

## 单元测试（待实现）

后续可添加：
- 模板解析器单元测试
- 路径解析工具测试
- 错误包装测试

## 集成测试（待实现）

后续可添加：
- 完整报表生成流程测试
- 并发生成测试
- 边界条件测试
