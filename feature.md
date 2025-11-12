# Feature: 基于模板的 Excel 报表链路

## 背景

- Carbone 模板保存在 `public/reportTemplates`
- 模板解析与渲染逻辑完全由开发者维护
- 当前目标：主进程中完成“源 Excel → 模板 → 报表”闭环，不涉及渲染器

## 总体目标

- 读取用户上传的 Excel，按模板定义解析成结构化数据
- 基于模板数据构建 Carbone 渲染上下文并输出报表文件
- 提供可扩展的模板注册体系，方便后续新增或维护模板

## 体系设计

### 模板注册中心

- 位置：`src/main/services/templates/registry.ts`
- 包含 `TemplateId`、`TemplateDefinition`、`TemplateMeta`、`TemplateParser`、`TemplateReportBuilder` 等类型
- 负责解析 dev/prod 模板物理路径、暴露 `getTemplate(id)` / `listTemplates()` API
- 每个模板一个模块（`src/main/services/templates/<templateId>.ts`）导出：
  - `parseWorkbook(workbook, parseOptions)`：用 exceljs 提取所需数据
  - `buildReportData(parsedData)`：生成 Carbone JSON 数据模型
  - 注册时指定模板文件名、支持的源文件扩展、Carbone 选项、输出命名器

### Excel → 数据服务 (`src/main/services/excelToData.ts`)

- `async excelToData({ sourcePath, templateId, parseOptions })`
- 校验模板存在、源文件合法、大小限制
- 使用 exceljs 读取 workbook，调用模板解析器
- 返回 `{ templateId, data, warnings, sourceMeta }`
- 抛出统一错误类型：`TemplateNotFoundError`、`UnsupportedFileError`、`ExcelParseError`

### 数据 → 报表服务 (`src/main/services/dataToReport.ts`)

- `async dataToReport({ templateId, parsedData, outputDir, renderOptions })`
- 获取模板定义，调用 `buildReportData` 得到 Carbone 数据
- `carbone.render(templatePath, reportData, carboneOptions)` 生成 Buffer
- 将结果写入 `outputDir`（默认 `app.getPath('documents')/reports`），返回 `{ outputPath, size, generatedAt }`
- 失败时清理临时文件并抛出 `ReportRenderError`

### 公共工具

- `src/main/services/errors.ts`：集中定义错误类，携带 `code/message/details`
- `src/main/services/utils/filePaths.ts`：模板路径解析、输出目录创建
- `src/main/services/utils/output.ts`：统一的文件命名策略（如 `<templateId>-YYYYMMDD-HHmmss.xlsx`）

## 开发步骤

1. 定义模板类型与错误类型，搭建模板注册中心骨架
2. 实现示例模板（如 `salesSummary`）：模板文件、解析器、报表数据构建
3. 编写 `excelToData`：文件校验、exceljs 加载、模板解析调度、错误包装
4. 编写 `dataToReport`：模板数据生成、Carbone 渲染、文件落盘、结果返回
5. 提供 CLI/脚本示例，方便本地验证；后续可挂 IPC 给渲染器

## 验证计划

- 使用示例 Excel 跑通 `excelToData` → `dataToReport`，确认生成报表正确
- 若条件允许，增加 Node 层测试（jest/vitest）覆盖模板解析和数据结构
- 打包后手动验证模板路径解析与 Carbone 渲染在生产环境可用
