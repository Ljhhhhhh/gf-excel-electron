# 开发计划：基于模板的 Excel 报表链路

## 项目目标
构建"源 Excel → 模板 → 报表"闭环，基于 Carbone + ExcelJS，完全在主进程中实现报表生成能力。

---

## 里程碑与步骤

### ✅ 阶段 0：需求确认（已完成）
- [x] 确认模板类型：仅 xlsx
- [x] 确认发布方式：extraResources 随应用打包
- [x] 确认输出方式：用户选择目录 + 自定义报表名
- [x] 确认语言/时区：zh-cn / Asia/Shanghai
- [x] 确认 Carbone 默认选项
- [x] 确认多表聚合策略：默认首表，按模板自定义解析
- [x] 确认错误码与返回结构

---

### ✅ 阶段 1：基础设施与配置（已完成）
- [x] 1.1 修改 electron-builder.yml：增加 extraResources 映射
- [x] 1.2 创建目录结构：
  - `src/main/services/templates/` - 模板注册与解析
  - `src/main/services/errors.ts` - 统一错误定义
  - `src/main/services/utils/` - 工具函数（路径解析、命名、文件操作）
- [x] 1.3 实现工具函数：
  - `getTemplatePath(templateId)` - dev/prod 路径解析
  - `generateReportName(templateId)` - 兜底命名策略
  - `ensureOutputDir(dir)` - 输出目录验证与创建
  - `openFolder(path)` - 打开所在文件夹（基于 shell）

---

### ✅ 阶段 2：类型定义与错误体系（已完成）
- [x] 2.1 定义核心类型（`src/main/services/templates/types.ts`）
- [x] 2.2 定义错误类（`src/main/services/errors.ts`）
- [x] 2.3 定义统一 Warning 结构

---

### ✅ 阶段 3：模板注册中心（已完成）
- [x] 3.1 实现 `src/main/services/templates/registry.ts`
- [x] 3.2 创建初始化入口 `src/main/services/templates/index.ts`

---

### ✅ 阶段 4：示例模板实现（已完成）
- [x] 4.1 验证模板文件 `public/reportTemplates/month1carbone.xlsx` 存在
- [x] 4.2 实现 `src/main/services/templates/month1carbone.ts`
- [x] 4.3 编写解析逻辑（表头识别、数据提取、多表聚合）

---

### ✅ 阶段 5：Excel → 数据服务（已完成）
- [x] 5.1 实现 `src/main/services/excelToData.ts`
- [x] 5.2 错误处理与校验
- [x] 5.3 添加日志输出

---

### ✅ 阶段 6：数据 → 报表服务（已完成）
- [x] 6.1 实现 `src/main/services/dataToReport.ts`
- [x] 6.2 错误处理与清理
- [x] 6.3 添加日志与耗时统计

---

### ✅ 阶段 7：命令行验证脚本（已完成）
- [x] 7.1 创建 `scripts/test-report-generation.ts`
- [x] 7.2 添加 npm script：`"test:report": "tsx scripts/test-report-generation.ts"`
- [x] 7.3 安装 tsx 依赖
- [x] 7.4 创建测试指南文档 `TESTING.md`
- [ ] 7.5 本地验证：运行脚本，确认生成报表正确（需要你准备测试 Excel 后执行）

---

### 📦 阶段 8：打包与生产验证（待执行）
- [ ] 8.1 运行 `npm run build:unpack`，检查 resources/reportTemplates 是否正确打包
- [ ] 8.2 手动验证生产环境路径解析：
  - 在 dist 目录中启动应用
  - 确认 getTemplatePath 正确解析到 resources/reportTemplates
- [ ] 8.3 冒烟测试：打包后运行命令行脚本，确认报表生成

---

### 🔗 阶段 9：tRPC 接口暴露（可选，后续与渲染器对接）
- [ ] 9.1 定义 tRPC router：`src/main/trpc/routers/report.ts`
  - `report.generate(input) -> { jobId }` 或直接返回结果
- [ ] 9.2 在 preload 暴露 tRPC 客户端
- [ ] 9.3 前端调用验证（待 UI 完成后）

---

### 📚 阶段 10：文档与测试
- [ ] 10.1 编写 README：
  - 模板开发指南
  - 添加新模板的步骤
  - 错误码参考表
- [ ] 10.2 编写单元测试（可选，使用 vitest）：
  - 模板解析逻辑
  - 路径解析工具
  - 错误包装
- [ ] 10.3 更新 AGENTS.md 与 feature.md（如有变化）

---

## 📊 当前进度总结

### ✅ 已完成（阶段 1-7）
- 基础设施：配置、目录结构、工具函数
- 类型系统：完整的类型定义与错误体系
- 模板系统：注册中心 + 示例模板 month1carbone
- 服务层：excelToData + dataToReport 完整实现
- 测试工具：命令行验证脚本 + 测试指南文档

### 🚧 待完成
- **阶段 7.5**：准备测试 Excel 并验证完整链路
- **阶段 8**：打包与生产环境验证

### 🎯 下一步行动
1. 准备一个测试用的 Excel 文件（参考 TESTING.md）
2. 运行 `pnpm test:report` 验证报表生成
3. 执行 `pnpm run build:unpack` 验证打包配置
4. 在生产环境中测试模板路径解析

---

## 📂 已创建文件清单

### 核心服务
- `src/main/services/errors.ts` - 统一错误定义
- `src/main/services/excelToData.ts` - Excel 解析服务
- `src/main/services/dataToReport.ts` - 报表生成服务

### 模板系统
- `src/main/services/templates/types.ts` - 类型定义
- `src/main/services/templates/registry.ts` - 注册中心
- `src/main/services/templates/index.ts` - 初始化入口
- `src/main/services/templates/month1carbone.ts` - 示例模板

### 工具函数
- `src/main/services/utils/filePaths.ts` - 路径解析
- `src/main/services/utils/naming.ts` - 命名策略
- `src/main/services/utils/fileOps.ts` - 文件操作

### 测试与文档
- `scripts/test-report-generation.ts` - 命令行测试脚本
- `TESTING.md` - 测试指南
- `TODO.md` - 开发计划（本文件）

### 配置
- `electron-builder.yml` - 已添加 extraResources 配置
- `package.json` - 已添加 test:report 脚本与 date-fns、tsx 依赖

---

## 注意事项
- 每完成一个子步骤，更新本文档状态为 [x]
- 遇到问题或需要调整时，先更新 TODO.md 再继续
- 保持代码风格一致，遵循项目 eslint 与 prettier 配置
- 提交前运行 `npm run lint` 与 `npm run typecheck`

---

## 附录：关键决策记录
- 模板类型：仅 xlsx
- 发布方式：extraResources
- 输出目录：用户每次选择，不持久化
- 报表名称：用户指定，兜底使用 `<templateId>-YYYYMMDD-HHmmss.xlsx`
- 语言/时区：zh-cn / Asia/Shanghai
- 多表聚合：默认首表，按模板自定义 parseOptions
- LibreOffice：xlsx→xlsx 不需要，仅格式转换时需要
