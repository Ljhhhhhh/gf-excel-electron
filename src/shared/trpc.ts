/**
 * tRPC 类型共享
 * 用于在渲染进程和主进程之间共享 tRPC 类型定义
 */

// 从主进程导出 AppRouter 类型
// 注意：这个文件只在构建时使用，运行时不会被打包到渲染进程
export type { AppRouter } from '../main/trpc/router'
