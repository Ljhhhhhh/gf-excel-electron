# 报表生成功能使用指南

## 概述

本项目实现了基于模板的 Excel 报表生成功能，支持：
- 📊 从源 Excel 文件解析数据
- 🎨 使用 Carbone 模板引擎渲染报表
- 🔄 支持多表聚合
- ⚙️ 可扩展的模板体系

## 快速开始

### 1. 准备测试数据

创建一个 Excel 文件（例如 `test-data.xlsx`），包含以下结构：

| 姓名 | 部门 | 销售额 | 日期 |
|------|------|--------|------------|
| 张三 | 销售部 | 10000 | 2024-01-01 |
| 李四 | 技术部 | 15000 | 2024-01-02 |
| 王五 | 销售部 | 12000 | 2024-01-03 |

### 2. 运行测试命令

```bash
# 基本用法
pnpm test:report --templateId=month1carbone --source=./test-data.xlsx --outDir=./output

# 自定义报表名称
pnpm test:report --templateId=month1carbone --source=./test-data.xlsx --outDir=./output --reportName=月度报表.xlsx
```

### 3. 查看生成结果

生成的报表将保存在指定的输出目录中，文件名格式：
- 未指定名称：`<templateId>-YYYYMMDD-HHmmss.xlsx`
- 指定名称：使用你提供的名称

## 架构说明

### 核心流程

```
源 Excel → excelToData → 结构化数据 → dataToReport → 报表文件
             (解析)                        (Carbone渲染)
```

### 目录结构

```
src/main/services/
├── templates/              # 模板系统
│   ├── types.ts           # 类型定义
│   ├── registry.ts        # 注册中心
│   ├── index.ts           # 初始化入口
│   └── month1carbone.ts   # 示例模板
├── utils/                  # 工具函数
│   ├── filePaths.ts       # 路径解析
│   ├── naming.ts          # 命名策略
│   └── fileOps.ts         # 文件操作
├── errors.ts              # 错误定义
├── excelToData.ts         # Excel解析服务
└── dataToReport.ts        # 报表生成服务
```

### 关键组件

#### 1. 模板注册中心 (`templates/registry.ts`)
- 管理所有可用模板
- 提供模板注册、获取、列举功能
- 校验模板文件存在性

#### 2. 模板定义
每个模板包含：
- **元信息**：ID、名称、文件名、支持的源文件类型
- **解析器**：`parseWorkbook(workbook, parseOptions)` - 从 Excel 提取数据
- **构建器**：`buildReportData(parsedData)` - 构建 Carbone 渲染数据

#### 3. 服务层
- **excelToData**：读取源文件 → 校验 → 解析 → 返回结构化数据
- **dataToReport**：构建数据 → Carbone 渲染 → 写入文件

## 添加新模板

### 步骤 1：创建模板文件
将 `.xlsx` 模板文件放入 `public/reportTemplates/`

### 步骤 2：实现模板模块
在 `src/main/services/templates/` 创建新文件（如 `myTemplate.ts`）：

```typescript
import type { Workbook } from 'exceljs'
import type { TemplateDefinition } from './types'

// 定义解析选项接口（可选）
interface MyParseOptions {
  sheets?: Array<string | number>
  // ...其他选项
}

// 实现解析器
export function parseWorkbook(workbook: Workbook, options?: MyParseOptions) {
  // 从 workbook 提取数据
  // 返回结构化数据
}

// 实现构建器
export function buildReportData(parsedData: unknown) {
  // 将解析数据转换为 Carbone 需要的格式
  return {
    // Carbone 数据模型
  }
}

// 导出模板定义
export const myTemplate: TemplateDefinition = {
  meta: {
    id: 'myTemplate',
    name: '我的模板',
    filename: 'myTemplate.xlsx',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx', 'xls'],
    description: '模板描述'
  },
  parser: parseWorkbook,
  builder: buildReportData,
  carboneOptions: {
    lang: 'zh-cn',
    timezone: 'Asia/Shanghai'
  }
}
```

### 步骤 3：注册模板
在 `src/main/services/templates/index.ts` 中注册：

```typescript
import { myTemplate } from './myTemplate'

export function initTemplates(): void {
  console.log('[Templates] 开始初始化模板系统...')
  
  registerTemplate(month1carboneTemplate)
  registerTemplate(myTemplate)  // 添加这行
  
  console.log('[Templates] 模板系统初始化完成')
}
```

### 步骤 4：测试
```bash
pnpm test:report --templateId=myTemplate --source=./test.xlsx --outDir=./output
```

## 配置说明

### Carbone 默认选项
- **语言**：`zh-cn`
- **时区**：`Asia/Shanghai`
- 可在模板定义或运行时通过 `renderOptions` 覆盖

### 文件限制
- 最大文件大小：100MB
- 支持扩展名：`.xlsx`, `.xls`（由模板定义）

### 路径解析
- 开发环境：`<项目根>/public/reportTemplates`
- 生产环境：`<app>/resources/reportTemplates`
- 自动根据 `app.isPackaged` 切换

## 错误处理

### 常见错误码
| 错误码 | 说明 | 解决方法 |
|--------|------|----------|
| `TEMPLATE_NOT_FOUND` | 模板不存在 | 检查 templateId 是否正确注册 |
| `EXCEL_UNSUPPORTED_FILE` | 文件不支持 | 检查文件格式与路径 |
| `EXCEL_FILE_TOO_LARGE` | 文件过大 | 文件需小于 100MB |
| `EXCEL_PARSE_ERROR` | 解析失败 | 检查文件完整性与格式 |
| `REPORT_RENDER_ERROR` | 渲染失败 | 检查模板与数据结构匹配 |
| `OUTPUT_WRITE_ERROR` | 写入失败 | 检查输出目录权限 |

## 性能考虑

- ExcelJS 解析：通常 < 200ms（对于标准大小文件）
- Carbone 渲染：通常 < 300ms
- 总耗时：通常 < 500ms

对于大文件（>10MB 或 >10000 行），考虑：
- 使用 `parseOptions.maxRows` 限制行数
- 分批处理

## 打包部署

### electron-builder 配置
已在 `electron-builder.yml` 中配置：

```yaml
extraResources:
  - from: public/reportTemplates
    to: reportTemplates
    filter:
      - '**/*'
```

### 验证打包
```bash
# 打包但不生成安装包
pnpm run build:unpack

# 检查 dist 目录中的 resources/reportTemplates
```

## 下一步开发

### 待实现功能
- [ ] tRPC 接口暴露给渲染进程
- [ ] 前端 UI 界面（模板选择、文件上传、进度显示）
- [ ] 任务队列与并发控制
- [ ] 批量生成
- [ ] 更多模板示例

### 扩展建议
- 支持 CSV 数据源
- 字段映射自动推荐
- PDF 导出（需 LibreOffice）
- 模板编辑器

## 参考资料

- [Carbone 官方文档](https://carbone.io/documentation.html)
- [ExcelJS 文档](https://github.com/exceljs/exceljs)
- [项目 AGENTS.md](./AGENTS.md) - 架构设计
- [项目 feature.md](./feature.md) - 功能需求
- [TESTING.md](./TESTING.md) - 详细测试指南

## 技术支持

遇到问题请参考：
1. 错误信息中的 `code` 和 `details`
2. 控制台日志（包含详细执行步骤）
3. TESTING.md 中的常见问题

---

**版本**：v1.0.0  
**最后更新**：2024-11-11
