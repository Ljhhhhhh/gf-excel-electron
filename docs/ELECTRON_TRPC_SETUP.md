# electron-trpc 配置说明

## 依赖安装

```bash
pnpm add electron-trpc @trpc/client@^10.45.2 @trpc/server@^10.45.2 superjson
```

**重要**：`electron-trpc` 需要 `@trpc/client` 和 `@trpc/server` 作为 peer dependencies，版本必须 >10.0.0。

## 配置结构

### 1. 服务端配置 (主进程)

**文件**: `src/main/trpc/trpc.ts`

```typescript
import { initTRPC } from '@trpc/server'
import superjson from 'superjson'
import type { Context } from './context'

const t = initTRPC.context<Context>().create({
  transformer: superjson,  // ← 必须配置
  errorFormatter({ shape }) {
    return shape
  }
})

export const router = t.router
export const publicProcedure = t.procedure
```

**文件**: `src/main/index.ts`

```typescript
import { createIPCHandler } from 'electron-trpc/main'
import { appRouter } from './trpc/router'

app.whenReady().then(() => {
  createWindow()
  
  createIPCHandler({
    router: appRouter,
    windows: [mainWindow],
    createContext: async () => createContext(mainWindow)
  })
})
```

### 2. Preload 配置

**文件**: `src/preload/index.ts`

```typescript
import { exposeElectronTRPC } from 'electron-trpc/main'

// 暴露 electron-trpc IPC 通道
exposeElectronTRPC()
```

**注意**：需要在 BrowserWindow 配置中指定 preload 文件路径。

### 3. 客户端配置 (渲染进程)

**文件**: `src/renderer/src/utils/trpc.ts`

```typescript
import { createTRPCProxyClient } from '@trpc/client'
import { ipcLink } from 'electron-trpc/renderer'
import superjson from 'superjson'
import type { AppRouter } from '@shared/trpc'

export const trpc = createTRPCProxyClient<AppRouter>({
  links: [ipcLink()],
  transformer: superjson  // ← 必须与服务端一致
})
```

## 关键要点

1. **transformer 必须在客户端和服务端都配置**，且必须一致
2. **不要**单独卸载 `@trpc/client` 和 `@trpc/server`，它们是 `electron-trpc` 的必需依赖
3. `exposeElectronTRPC()` 必须在 preload 中调用
4. `createIPCHandler()` 必须在窗口创建后调用

## 常见错误

### 错误 1: "Cannot read properties of undefined (reading 'serialize')"

**原因**：transformer 配置不一致或缺失

**解决**：确保客户端和服务端都配置了相同的 transformer

### 错误 2: peer dependency 警告

**原因**：缺少 `@trpc/client` 或 `@trpc/server`

**解决**：
```bash
pnpm add @trpc/client@^10.45.2 @trpc/server@^10.45.2
```

### 错误 3: IPC 通道未暴露

**原因**：preload 中未调用 `exposeElectronTRPC()`

**解决**：在 preload 文件中添加该调用

## 测试验证

启动开发服务器：
```bash
pnpm dev
```

在 Vue 组件中测试：
```typescript
import { trpc } from '@/utils/trpc'

// 查询
const templates = await trpc.template.list.query()

// 变更
const result = await trpc.report.generate.mutate({...})
```

## 参考文档

- [electron-trpc 官方文档](https://electron-trpc.dev)
- [tRPC 文档](https://trpc.io)
- [superjson 文档](https://github.com/blitz-js/superjson)
