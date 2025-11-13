## 项目现状概览

- 当前无前端路由，`App.vue` 直接渲染 `ReportTest.vue`；UI 使用 Element Plus，动态表单使用 form-create；tRPC 通过 IPC 调用主进程服务。

## 设计目标

- 以“模板中心 → 报表生成”两大页面为核心，形成清晰的两步业务闭环。
- 统一 UI 与交互，减少自定义样式，提升易用性与可维护性。

## 技术选型

- Vue 3 + Element Plus + form-create（保留现有栈）。
- 引入 `vue-router` 管理两大页面路由（依赖已存在）。

## 信息架构

- 顶层 AppShell：Header（标题、当前模板状态、快捷入口）+ Sidebar（导航）+ Content（`<router-view />`）。
- 导航与页面：
  - 模板中心（`/templates`）：浏览/搜索/验证/选择模板。
  - 报表生成（`/report`）：基于所选模板的向导式生成流程。

## 路由与导航

- 新增 `src/renderer/src/router/index.ts`：
  - 定义路由：`/templates`、`/report`。
  - 路由守卫：如果未选择模板，访问 `/report` 时自动跳转至 `/templates` 并提示选择模板。
- `App.vue` 改造：引入侧边导航与头部，内容区域渲染 `RouterView`。

## 关键页面

- 模板中心（TemplatesView）：
  - 加载与展示模板列表（卡片/表格视图），支持搜索与筛选（按名称、文件名）。
  - 查看模板详情（元信息、文件位置），提供“校验模板”与“设为当前模板”。
  - 选中模板后，在 Header 显示当前模板信息与“前往生成”按钮。
- 报表生成（ReportWizardView）：
  - `el-steps` 分步：1.选择源文件 2.选择输出目录 3.填写参数 4.生成与结果。
  - 复用 `TemplateInputForm` 动态参数；外层增加整体校验与状态提示。
  - 生成过程：Loading、错误通知；成功展示结果卡片（输出路径、大小、时间、耗时、打开目录）。

## 组件拆分

- 布局组件：`AppHeader.vue`（当前模板与快捷操作）、`SidebarNav.vue`（仅两项导航）、`ContentContainer.vue`。
- 生成流程组件：`ResultCard.vue`（统一成功/失败展示）、`FilePathField.vue`（封装 tRPC 文件/目录选择与只读展示）。
- 复用组件：`TemplateInputForm.vue`。

## 状态与数据流

- 组合式状态管理：
  - `useTemplates()`：模板列表、当前选中模板、校验与元信息获取；提供 `setCurrentTemplate()`。
  - `useReport()`：源文件、输出目录、参数、生成结果；提供生成与校验方法。
- 两页面共享状态，避免重复请求与数据分散。

## 交互与校验

- 使用 Element Plus 表单校验与 `ElMessage`/`ElNotification` 提示统一成功与错误。
- 在未选择模板时禁用生成相关操作并明确引导；生成中禁用会导致状态不一致的按钮。

## 样式与主题

- 以 Element Plus 组件为主，减少手写 CSS；必要样式集中在 `assets/main.css` 微调。
- 保留现有暗色变量与基础样式，确保与 Element Plus 协同。

## 性能与健壮性

- 模板与规则缓存，避免重复请求；tRPC 统一错误处理并显示可读信息。

## Electron 集成

- 保持现有单窗口架构与 tRPC 接口使用；文件/目录选择与结果打开沿用 `fileRouter`。

## 实施步骤

1. 路由与布局：新增 `router/index.ts`，改造 `App.vue` 为 AppShell（Header+Sidebar+RouterView）。
2. 页面搭建：实现 `TemplatesView` 与 `ReportWizardView`，迁移与拆分 `ReportTest.vue` 逻辑。
3. 组件封装：复用 `TemplateInputForm`；新增 `ResultCard` 与 `FilePathField`；统一消息与 Loading。
4. 状态抽取：实现 `useTemplates` 与 `useReport` 并在两页面使用，减少重复 `ref` 与手动校验。
5. 样式统一：用 Element Plus 组件替换自定义 UI；仅保留必要的局部样式。
6. 验证打磨：跑通“选择模板 → 生成报表 → 查看结果”闭环，优化提示与边界处理。

## 验收与验证

- 导航清晰，两页面功能分明；模板选择与生成流程顺畅。
- 结果展示准确，错误提示清晰；与现有主进程服务完全兼容。

说明：已按要求移除“任务记录”和“设置”页面；不引入国际化，仅保留中文界面。
