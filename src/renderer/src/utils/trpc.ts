/**
 * tRPC 客户端实例
 * 用于渲染进程调用主进程的 tRPC API
 */

import { createTRPCProxyClient } from '@trpc/client'
import { ipcLink } from 'electron-trpc/renderer'
import superjson from 'superjson'
import type { AppRouter } from '@shared/trpc'

/**
 * 创建 tRPC 客户端
 * 注意：transformer 必须与服务端配置保持一致
 */
export const trpc = createTRPCProxyClient<AppRouter>({
  links: [ipcLink()],
  transformer: superjson
})
