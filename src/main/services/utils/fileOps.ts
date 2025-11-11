/**
 * 文件操作工具
 * 负责打开文件夹、文件大小检查等
 */

import { shell } from 'electron'
import fs from 'node:fs'
import path from 'node:path'

/**
 * 在文件管理器中打开指定路径所在的文件夹
 * @param filePath 文件或文件夹路径
 */
export async function openFolder(filePath: string): Promise<void> {
  // 如果是文件，则打开其所在文件夹并选中该文件
  // 如果是文件夹，则直接打开
  const stats = fs.statSync(filePath)
  if (stats.isDirectory()) {
    await shell.openPath(filePath)
  } else {
    shell.showItemInFolder(filePath)
  }
}

/**
 * 获取文件大小（字节）
 */
export function getFileSize(filePath: string): number {
  const stats = fs.statSync(filePath)
  return stats.size
}

/**
 * 检查文件大小是否超过限制
 * @param filePath 文件路径
 * @param limitMB 限制大小（MB）
 * @returns 是否超限
 */
export function isFileTooLarge(filePath: string, limitMB: number): boolean {
  const size = getFileSize(filePath)
  const limitBytes = limitMB * 1024 * 1024
  return size > limitBytes
}

/**
 * 获取文件扩展名（小写，不含点）
 */
export function getFileExtension(filePath: string): string {
  return path.extname(filePath).toLowerCase().replace('.', '')
}

/**
 * 删除文件（如果存在）
 */
export function deleteFileIfExists(filePath: string): void {
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath)
  }
}
