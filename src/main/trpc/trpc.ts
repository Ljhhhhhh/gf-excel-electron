/**
 * tRPC 实例初始化
 * 定义 tRPC 基础配置、中间件和 procedure 构建器
 */

import { initTRPC, TRPCError } from '@trpc/server'
import superjson from 'superjson'
import type { Context } from './context'

/**
 * 初始化 tRPC 实例
 * 注意：transformer 必须在客户端和服务端同时配置
 */
const t = initTRPC.context<Context>().create({
  transformer: superjson,
  errorFormatter({ shape }) {
    return shape
  }
})

/**
 * 导出 tRPC 工具
 */
export const router = t.router
export const publicProcedure = t.procedure

/**
 * tRPC 错误类型导出
 */
export { TRPCError }
