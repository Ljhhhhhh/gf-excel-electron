/**
 * 文件路径解析工具
 * 负责开发/生产环境下的模板路径解析
 */

import path from 'node:path'
import fs from 'node:fs'

// 动态导入 electron app，支持命令行脚本环境
let electronApp: { isPackaged: boolean } | undefined
try {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  electronApp = require('electron').app
} catch {
  // 命令行环境，检查 global.app（测试脚本注入的 mock）
  electronApp = (global as { app?: { isPackaged: boolean } }).app
}

/**
 * 判断是否为开发环境
 */
export function isDev(): boolean {
  // 如果没有 app 对象，默认为开发环境
  if (!electronApp) {
    return true
  }
  return !electronApp.isPackaged
}

/**
 * 获取模板根目录
 * - 开发环境: <repo>/public/reportTemplates
 * - 生产环境: <app>/resources/reportTemplates
 */
export function getTemplateRootDir(): string {
  if (isDev()) {
    // 开发环境：从项目根目录的 public/reportTemplates
    return path.join(process.cwd(), 'public', 'reportTemplates')
  } else {
    // 生产环境：从 process.resourcesPath/reportTemplates
    return path.join(process.resourcesPath, 'reportTemplates')
  }
}

/**
 * 获取指定模板的完整路径
 * @param filename 模板文件名（如 'month1carbone.xlsx'）
 */
export function getTemplatePath(filename: string): string {
  const rootDir = getTemplateRootDir()
  // const rootDir = '../../../../resources/reportTemplates'
  const templatePath = path.join(rootDir, filename)
  return templatePath
}

/**
 * 检查模板文件是否存在
 */
export function templateExists(filename: string): boolean {
  const templatePath = getTemplatePath(filename)
  return fs.existsSync(templatePath)
}

/**
 * 确保输出目录存在，不存在则创建
 */
export function ensureOutputDir(dir: string): void {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true })
  }
}

/**
 * 验证路径是否安全（防止路径穿越）
 */
export function isSafePath(targetPath: string, baseDir: string): boolean {
  const normalizedTarget = path.normalize(targetPath)
  const normalizedBase = path.normalize(baseDir)
  return normalizedTarget.startsWith(normalizedBase)
}
