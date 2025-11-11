# tRPC API 使用指南

## 概述

本项目使用 tRPC + electron-trpc 实现类型安全的主进程与渲染进程通信。所有 API 调用都具有完整的 TypeScript 类型推断和自动补全。

## 架构

```
渲染进程 (Vue)
    ↓ window.api.trpc (类型安全调用)
Preload (tRPC Client Proxy)
    ↓ IPC 通信
主进程 (tRPC Router → Services)
```

## 在前端使用 tRPC

### 基本用法

```typescript
// 在 Vue 组件中
const trpc = window.api.trpc

// Query 操作（查询数据）
const templates = await trpc.template.list.query()

// Mutation 操作（修改数据/执行操作）
const result = await trpc.report.generate.mutate({
  templateId: 'month1carbone',
  sourcePath: '/path/to/source.xlsx',
  outputDir: '/path/to/output',
  reportName: '月度报表.xlsx'
})
```

### 错误处理

```typescript
try {
  const result = await trpc.report.generate.mutate({ ... })
  console.log('成功:', result)
} catch (error) {
  // error 是 TRPCClientError，包含详细错误信息
  console.error('失败:', error.message)
}
```

---

## API 接口文档

### 1. Template Router (模板管理)

#### `template.list`

列出所有已注册的模板。

```typescript
const templates = await trpc.template.list.query()
// 返回: TemplateMeta[]
```

**返回类型**：
```typescript
interface TemplateMeta {
  id: string              // 模板 ID
  name: string            // 模板显示名称
  filename: string        // 模板文件名
  ext: 'xlsx'             // 文件扩展名
  supportedSourceExts: string[]  // 支持的源文件扩展名
  description?: string    // 描述
}
```

#### `template.getMeta`

获取指定模板的元信息。

```typescript
const meta = await trpc.template.getMeta.query({ id: 'month1carbone' })
// 返回: TemplateMeta
```

**参数**：
- `id: string` - 模板 ID

**错误**：
- `NOT_FOUND` - 模板不存在

#### `template.validate`

校验模板文件是否存在。

```typescript
const result = await trpc.template.validate.query({ id: 'month1carbone' })
// 返回: { ok: boolean, issues?: string[] }
```

**参数**：
- `id: string` - 模板 ID

**返回**：
```typescript
{
  ok: boolean           // 是否通过校验
  issues?: string[]     // 问题列表（校验失败时）
}
```

---

### 2. File Router (文件操作)

#### `file.selectSourceFile`

打开文件选择对话框，选择源 Excel 文件。

```typescript
const result = await trpc.file.selectSourceFile.query()
// 返回: { canceled: boolean, filePath: string | null }
```

**返回**：
```typescript
{
  canceled: boolean      // 是否取消选择
  filePath: string | null  // 选择的文件路径（取消时为 null）
}
```

#### `file.selectOutputDir`

打开目录选择对话框，选择输出目录。

```typescript
const result = await trpc.file.selectOutputDir.query()
// 返回: { canceled: boolean, dirPath: string | null }
```

**返回**：
```typescript
{
  canceled: boolean      // 是否取消选择
  dirPath: string | null   // 选择的目录路径（取消时为 null）
}
```

#### `file.openInFolder`

在文件管理器中打开指定路径（文件或文件夹）。

```typescript
await trpc.file.openInFolder.mutate({ path: '/path/to/file.xlsx' })
// 返回: { success: true }
```

**参数**：
- `path: string` - 文件或文件夹路径

**错误**：
- `NOT_FOUND` - 路径不存在

---

### 3. Report Router (报表生成)

#### `report.generate`

生成报表（同步返回结果）。

```typescript
const result = await trpc.report.generate.mutate({
  templateId: 'month1carbone',
  sourcePath: '/path/to/source.xlsx',
  outputDir: '/path/to/output',
  reportName: '月度报表-2024.xlsx',  // 可选
  parseOptions: {},                  // 可选
  renderOptions: {}                  // 可选
})
```

**参数**：
```typescript
{
  templateId: string           // 模板 ID（必填）
  sourcePath: string           // 源文件路径（必填）
  outputDir: string            // 输出目录（必填）
  reportName?: string          // 报表名称（可选，留空使用默认）
  parseOptions?: Record<string, unknown>   // 模板解析选项（可选）
  renderOptions?: Record<string, unknown>  // Carbone 渲染选项（可选）
}
```

**返回**：
```typescript
{
  success: true                // 成功标志
  outputPath: string          // 输出文件路径
  size: number                // 文件大小（字节）
  generatedAt: Date           // 生成时间
  warnings: Warning[]         // 警告信息
  duration: number            // 耗时（毫秒）
}
```

**错误**：
- `NOT_FOUND` - 模板不存在
- `BAD_REQUEST` - 参数错误、文件不支持、文件过大、解析失败等
- `INTERNAL_SERVER_ERROR` - 渲染失败、写入失败等

---

## 完整示例：报表生成流程

```vue
<script setup lang="ts">
import { ref } from 'vue'

const trpc = window.api.trpc

const templateId = ref('month1carbone')
const sourcePath = ref<string | null>(null)
const outputDir = ref<string | null>(null)
const reportName = ref('')
const generating = ref(false)
const result = ref<any>(null)

// 选择源文件
async function selectSource() {
  const res = await trpc.file.selectSourceFile.query()
  if (!res.canceled && res.filePath) {
    sourcePath.value = res.filePath
  }
}

// 选择输出目录
async function selectOutput() {
  const res = await trpc.file.selectOutputDir.query()
  if (!res.canceled && res.dirPath) {
    outputDir.value = res.dirPath
  }
}

// 生成报表
async function generate() {
  if (!sourcePath.value || !outputDir.value) return

  generating.value = true
  try {
    result.value = await trpc.report.generate.mutate({
      templateId: templateId.value,
      sourcePath: sourcePath.value,
      outputDir: outputDir.value,
      reportName: reportName.value || undefined
    })
    console.log('报表生成成功:', result.value)
  } catch (error) {
    console.error('报表生成失败:', error)
    alert('生成失败: ' + (error as any).message)
  } finally {
    generating.value = false
  }
}

// 打开结果文件夹
async function openResult() {
  if (result.value?.outputPath) {
    await trpc.file.openInFolder.mutate({ path: result.value.outputPath })
  }
}
</script>

<template>
  <div>
    <button @click="selectSource">选择源文件</button>
    <button @click="selectOutput">选择输出目录</button>
    <button @click="generate" :disabled="generating">
      {{ generating ? '生成中...' : '生成报表' }}
    </button>
    <button v-if="result" @click="openResult">打开文件夹</button>
  </div>
</template>
```

---

## 类型推断与自动补全

得益于 tRPC，你可以享受到完整的类型推断：

```typescript
// ✅ TypeScript 自动推断返回类型
const templates = await trpc.template.list.query()
// templates 类型: TemplateMeta[]

// ✅ 参数自动校验
await trpc.template.getMeta.query({ id: 'xxx' })  // ✓ 正确
await trpc.template.getMeta.query({ id: 123 })    // ✗ 类型错误

// ✅ 返回值类型推断
const result = await trpc.report.generate.mutate({ ... })
result.outputPath  // ✓ 类型: string
result.size       // ✓ 类型: number
```

---

## 错误码参考

| 错误码 | 说明 | 触发场景 |
|-------|------|---------|
| `NOT_FOUND` | 资源不存在 | 模板不存在、文件/文件夹不存在 |
| `BAD_REQUEST` | 请求参数错误 | 文件不支持、文件过大、解析失败 |
| `INTERNAL_SERVER_ERROR` | 服务器内部错误 | 渲染失败、写入失败 |

---

## 调试技巧

### 1. 查看主进程日志

主进程会输出详细的日志，包括：
- tRPC handler 挂载状态
- 每个 API 调用的开始和结束
- 错误堆栈

### 2. 使用 DevTools

按 `F12` 打开开发者工具，在 Console 中可以：
- 查看 tRPC 调用日志
- 查看错误详情
- 测试 API 调用：`window.api.trpc.template.list.query()`

### 3. 类型检查

运行 `pnpm typecheck` 确保类型正确：
```bash
pnpm typecheck
```

---

## 扩展 API

如需添加新的 API 接口：

1. **创建/修改 Router**（如 `src/main/trpc/routers/xxx.ts`）
2. **在根 Router 中注册**（`src/main/trpc/router.ts`）
3. **无需修改前端代码**，类型会自动推断

示例：
```typescript
// src/main/trpc/routers/settings.ts
export const settingsRouter = router({
  get: publicProcedure.query(() => {
    return { theme: 'dark', lang: 'zh-cn' }
  }),
  
  set: publicProcedure
    .input(z.object({ key: z.string(), value: z.any() }))
    .mutation(({ input }) => {
      // 保存设置
      return { success: true }
    })
})

// src/main/trpc/router.ts
export const appRouter = router({
  template: templateRouter,
  file: fileRouter,
  report: reportRouter,
  settings: settingsRouter  // 新增
})

// 前端立即可用，且有类型推断：
await trpc.settings.get.query()
await trpc.settings.set.mutate({ key: 'theme', value: 'light' })
```

---

## 注意事项

1. **Query vs Mutation**
   - Query: 查询操作，不改变状态（`query()`）
   - Mutation: 修改操作，改变状态（`mutate()`）

2. **大文件传输**
   - 避免在 IPC 中传递大 Buffer
   - 使用文件路径引用

3. **异步操作**
   - 所有 tRPC 调用都是异步的，返回 Promise
   - 使用 `async/await` 或 `.then()/.catch()`

4. **错误处理**
   - 始终使用 try/catch 包裹 tRPC 调用
   - 错误对象包含 `message`、`code`、`cause` 等信息

---

## 参考文档

- [tRPC 官方文档](https://trpc.io/)
- [electron-trpc](https://github.com/jsonnull/electron-trpc)
- [Zod 验证库](https://zod.dev/)
