/**
 * tRPC 根 Router
 * 合并所有子 router 并导出类型
 */

import { router } from './trpc'
import { templateRouter } from './routers/template'
import { fileRouter } from './routers/file'
import { reportRouter } from './routers/report'

/**
 * 应用根 router
 */
export const appRouter = router({
  template: templateRouter,
  file: fileRouter,
  report: reportRouter
})

/**
 * 导出 AppRouter 类型供前端使用
 */
export type AppRouter = typeof appRouter
