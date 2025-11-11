/**
 * tRPC Context 定义
 * 为所有 tRPC procedures 提供共享上下文
 */

import type { BrowserWindow } from 'electron'

/**
 * tRPC 上下文接口
 */
export interface Context {
  /** 主窗口实例（可选，用于对话框等） */
  mainWindow?: BrowserWindow
}

/**
 * 创建 tRPC 上下文
 * @param mainWindow 主窗口实例
 * @returns Context 对象
 */
export function createContext(mainWindow?: BrowserWindow): Context {
  return {
    mainWindow
  }
}
