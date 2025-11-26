# AGENTS 设计说明（gf-excel-electron）

本文件阐述项目在 Electron 主进程与 Vue 渲染进程之间的“Agent 化”职责分解与通信约定，聚焦“源数据文件 + Carbone 模板 + ExcelJS 渲染”的报表生成能力，强调优雅架构、可维护性与可测试性。

---

## 1. 目标与非目标

- 目标
  - 基于 Carbone 与 ExcelJS 的报表输出（xlsx）。
  - 支持“模板库 + 源数据文件（xlsx）+ 映射配置”的一键生成。
  - 通过 electron-trpc 提供类型安全的后端能力给前端使用。
  - 保持清晰的进程边界与最小权限原则；确保可观测、可诊断、可扩展。
- 非目标（当前版本不做）
  - 在线协同编辑模板。
  - 复杂工作流编排与分布式执行。

---

## 2. 技术栈映射与边界

- 后端（主进程 / Node）
  - electron-vite：构建与开发。
  - electron-trpc：IPC 层上的类型安全 RPC。
  - carbone：基于 Office 模板的数据填充与生成。
  - exceljs：直接操作 xlsx 工作簿进行渲染。
  - execa：受控执行外部程序（如可选的转换工具）。
- 前端（渲染进程 / Vue）
  - vue, vue-router：UI 框架与路由。
  - tailwindcss, Element Plus：样式与组件库。
- 边界原则
  - 渲染进程不得直接访问文件系统；仅通过 tRPC 使用受控能力。
  - Preload 仅暴露 tRPC 客户端与必要的只读环境信息。

---

## 3. 进程与模块拓扑

- Main（主进程）
  - TRPC Router 层：对外暴露 API（report/template/dataset/job/export...）。
  - Agents（业务协调器）：编排多服务与策略选择。
  - Services（原子能力）：文件、模板、渲染、转换、队列、配置等。
- Preload（隔离层）
  - 仅实例化并暴露 tRPC 客户端与基础事件桥接（如需要）。
- Renderer（渲染进程）
  - 视图与交互逻辑，通过 tRPC 调用主进程能力。

---

## 4. Agent 划分与职责

- TemplateAgent（模板管理）
  - 职责：模板库 CRUD、校验、预览元信息（文件名、类型、字段占位抽取-可选）。
  - 依赖：FsService。
  - 备注：Carbone 模板与 ExcelJS 模板可共存；统一登记与类型标注。
- DatasetAgent（源数据管理）
  - 职责：导入/登记源数据文件（xlsx）、抽样预览、字段统计。
  - 依赖：FsService、ParsingService（基于 exceljs 解析）。
- ReportAgent（报表生成协调）
  - 职责：根据策略选择“Carbone 模式”或“ExcelJS 模式”生成报表；支持批量、进度、取消。
  - 依赖：CarboneService、ExcelService、JobService、FsService。
- ExportAgent（导出与后处理）
  - 职责：整理输出目录、命名、打开所在文件夹。
  - 依赖：FsService、ShellService（execa）。
- JobAgent（任务队列/状态机）
  - 职责：维护 Job 生命周期（pending/running/success/failed/canceled）、并发与取消。
  - 依赖：JobService（内部实现队列与事件）。
- SettingsAgent（配置）
  - 职责：读取/写入应用配置（并发度、默认引擎等）。不保存输出目录。
  - 依赖：SettingsService（基于 userData 路径的 json 存储）。

---

## 5. 服务层（Services）

- FsService：安全文件读写、路径白名单、临时目录管理、清理策略。
- ParsingService：数据文件解析（xlsx），返回统一表格/记录集结构。
- CarboneService：封装 carbone.render，提供模板检查与输出参数规范化。
- ExcelService：封装 exceljs，支持基于模板的填充、样式与公式处理。
- ShellService：封装 execa，执行受限外部命令（仅白名单）。
- JobService：任务队列、进度事件、取消（基于 AbortController）、持久化（可选）。
- SettingsService：应用配置的持久化读取/写入。

---

## 6. 目录建议

- 主进程
  - src/main/agents/\*
  - src/main/services/\*
  - src/main/trpc/routers/\*
  - 运行时数据目录（自动创建）
    - <userData>/templates
    - <userData>/datasets
    - <userData>/temp
- 渲染进程
  - src/renderer/pages/\*（templates/datasets/reports/jobs/settings）
  - src/renderer/composables/\*（useReport/useTemplates/...）
  - src/renderer/components/\*
- Preload
  - src/preload/trpc.ts（暴露类型安全客户端）

---

## 7. tRPC 路由草案（对外接口）

仅为接口原型，具体命名与细节以实现为准。

```ts
// 命名空间 report
interface GenerateInput {
  engine: 'carbone' | 'exceljs'
  templateId: string
  datasetId: string
  mapping?: Record<string, string> // 字段映射：模板字段 -> 数据字段
  output: {
    format: 'xlsx'
    filename?: string
    outDir?: string // 每次由用户选择，不在软件内持久化
  }
  options?: Record<string, any> // 引擎特定选项
}
interface GenerateResult {
  jobId: string
}

// 命名空间 template
interface TemplateMeta {
  id: string
  name: string
  type: 'carbone' | 'exceljs'
  ext: 'xlsx' | 'docx' | 'odt'
  path: string
}

// 命名空间 dataset
interface DatasetMeta {
  id: string
  name: string
  type: 'xlsx'
  path: string
  preview?: { headers: string[]; rows: any[][] }
}

// 命名空间 job
type JobStatus = 'pending' | 'running' | 'success' | 'failed' | 'canceled'
interface JobInfo {
  id: string
  status: JobStatus
  progress: number
  resultPath?: string
  error?: { code: string; message: string }
}
```

- report.generate(input: GenerateInput) -> { jobId }
- job.get(jobId) -> JobInfo
- job.cancel(jobId) -> void
- template.list() -> TemplateMeta[]
- template.add(file: FileRef, meta?: Partial<TemplateMeta>) -> TemplateMeta
- template.remove(id: string) -> void
- template.validate(id: string) -> { ok: boolean; issues?: string[] }
- dataset.list() -> DatasetMeta[]
- dataset.add(file: FileRef, options?: { sheet?: string }) -> DatasetMeta
- dataset.remove(id: string) -> void
- dataset.preview(id: string, limit?: number) -> DatasetMeta['preview']
- export.openInFolder(path: string) -> void

> 说明：FileRef 是由渲染进程选择文件后传入主进程的受控引用（如路径 + 摘要），由主进程完成复制到白名单目录并登记。

---

## 8. 数据与模板规范

- 模板
  - 支持：.xlsx（ExcelJS/Carbone）；docx/odt 可扩展（Carbone，待确认）。
  - 统一元数据：id、name、type、ext、path。
  - 命名建议：<业务域>/<场景>/<版本>/<名称>.<ext>（例如 sales/monthly/v1/report.xlsx）。
- 源数据
  - 支持：.xlsx；导入后复制到 <userData>/datasets 并生成 DatasetMeta。
  - 解析策略：优先第一张表/指定 sheet。
  - 大小限制：单个源文件不超过 100MB，超出直接拒绝。
- 输出
  - 每次生成前通过系统目录选择对话框确定输出目录，不在应用内持久化；文件命名：<模板名>-<数据集名>-<yyyyMMddHHmmss>.<ext>。

---

## 9. 生成策略与流程

- Carbone 模式（适合复杂版式、跨格式转换）
  1. 校验模板与数据 -> 规范化 options。
  2. carbone.render(template, data) -> Buffer/File。
  3. 写入用户选择的输出目录 -> 返回 job 完成与路径。

- ExcelJS 模式（适合纯 xlsx、细粒度单元格控制）
  1. 打开模板工作簿 -> 按映射填充数据/写入公式/样式。
  2. 写入用户选择的输出目录。

- 批量/并发
  - JobService 控制并发度（默认 2，可在 Settings 中配置）。
  - 支持取消：基于 AbortController，Carbone/ExcelJS 任务需轮询检查。

---

## 10. 错误处理与可观测性

- 统一错误结构：{ code, message, details? }，按域划分错误码（TEMPLATE*\*, DATASET*\_, REPORT\_\_, JOB*\*, EXPORT*\*）。
- 日志：main 侧输出结构化日志；失败任务保留上下文（输入、栈、采样数据路径）。
- 进度：JobInfo.progress 0~100；长任务每步推进并持久化（可选）。

---

## 11. 安全与权限

- 渲染进程无文件系统直达；所有文件 I/O 走 FsService 白名单路径。
- ShellService 仅允许配置指定命令与参数模板（仅白名单）。
- 路径净化与扩展名校验；拒绝路径穿越与未知扩展名。
- contextIsolation 与 sandbox 保持开启；Preload 暴露最小接口。

---

## 12. 性能建议

- ExcelJS：优先使用流式写入（writeFile/stream），避免内存峰值。
- Carbone：模板复用与缓存；大数据分片合并（如需要）。
- I/O：避免在渲染进程传递大 Buffer，改为路径引用。

---

## 13. 前端信息架构（路由与页面）

- /templates 模板库
  - 列表、上传、删除、校验、元信息预览。
- /datasets 数据源
  - 列表、导入、预览、删除。
- /reports/new 生成向导
  - 选择模板 + 数据 -> 字段映射 -> 生成 -> 查看任务。
- /jobs 任务与历史
  - 进度、取消、打开所在文件夹、失败重试。
- /settings 设置
  - 并发度、默认引擎、外部工具路径（可选）。不保存输出目录。

---

## 14. 验收与测试

- 单元测试：Services 与 Agents 的纯函数/边界行为。
- 集成测试：tRPC 路由端到端（主进程内存级）。
- 冒烟用例：
  - 导入 xlsx 数据集并预览。
  - 导入 carbone xlsx 模板并生成 xlsx 报表。
  - 导入 exceljs 模板并生成 xlsx 报表。
  - 并发 2 个任务与一个取消场景。

---

## 15. 里程碑（建议）

1. tRPC 基础与 Preload 暴露，FsService 与 SettingsService 落地。
2. TemplateAgent/DatasetAgent MVP（列表/新增/删除/预览）。
3. ReportAgent（Carbone 模式）MVP：单任务生成 xlsx。
4. ReportAgent（ExcelJS 模式）MVP：单任务生成 xlsx。
5. JobAgent/ExportAgent：队列、进度、取消、打开文件夹。
6. 错误码体系与日志；端到端冒烟测试。

---

## 16. 待确认清单（需与你确认后落地）

- 模板类型范围：是否同时支持 .docx/.odt？默认先以 .xlsx 为主？
- 源数据类型：当前仅支持 .xlsx（如需扩展 csv/json 可后续评估）。
- PDF 导出：当前不需要（如需再行评估与设计）。
- 输出目录：每次由用户选择且不持久化（已确定）。
- 并发上限与默认值（建议默认 2，最大 4）。
- 是否需要字段占位自动抽取与自动映射建议？

---

## 17. 术语

- 模板：报表版式定义文件（Carbone 或 ExcelJS）。
- 数据集：原始数据文件的受控副本与其元信息。
- 任务（Job）：一次报表生成流程的调度单位。
- 引擎：Carbone 或 ExcelJS。

注意：所有的 Excel 文件（包括主数据源和额外数据源）加载逻辑统一修改为 使用 fs.createReadStream 创建文件流，并通过 workbook.xlsx.read(stream) 进行加载。
