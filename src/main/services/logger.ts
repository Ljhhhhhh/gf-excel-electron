/**
 * 简易日志服务
 * - 将结构化日志写入 <userData>/logs/app.log（或开发环境的 ./logs）
 * - 同步输出到控制台，便于开发调试
 */

import fs from 'node:fs'
import path from 'node:path'

type LogLevel = 'INFO' | 'WARN' | 'ERROR'
type LogContext = Record<string, unknown> | undefined

// 尝试获取 Electron app（在命令行脚本中可能不存在）
let electronApp: { getPath?(name: string): string; isReady?(): boolean } | undefined
try {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  electronApp = require('electron').app
} catch {
  electronApp = (global as { app?: typeof electronApp }).app
}

const LOG_DIR_NAME = 'logs'
const LOG_FILE_NAME = 'app.log'
let cachedLogFilePath: string | null = null

function resolveLogFile(): string {
  if (cachedLogFilePath) return cachedLogFilePath

  let baseDir: string | undefined
  if (electronApp?.getPath) {
    try {
      baseDir = electronApp.getPath('userData')
    } catch {
      /* 在 app 未 ready 时可能失败，改用 cwd */
    }
  }
  if (!baseDir) {
    baseDir = process.cwd()
  }

  const logDir = path.join(baseDir, LOG_DIR_NAME)
  if (!fs.existsSync(logDir)) {
    fs.mkdirSync(logDir, { recursive: true })
  }

  cachedLogFilePath = path.join(logDir, LOG_FILE_NAME)
  console.log('Log file path:', cachedLogFilePath)
  return cachedLogFilePath
}

function formatContext(context?: LogContext): string {
  if (!context) return ''
  try {
    return ` | ${JSON.stringify(context)}`
  } catch {
    return ` | ${String(context)}`
  }
}

function writeLog(level: LogLevel, scope: string, message: string, context?: LogContext): void {
  const timestamp = new Date().toISOString()
  const line = `${timestamp} [${level}] [${scope}] ${message}${formatContext(context)}`

  // 控制台输出（保留开发时体验）
  switch (level) {
    case 'WARN':
      console.warn(line)
      break
    case 'ERROR':
      console.error(line)
      break
    default:
      console.log(line)
  }

  // 附加写入日志文件
  try {
    const logPath = resolveLogFile()
    fs.appendFileSync(logPath, `${line}\n`, 'utf-8')
  } catch (error) {
    // 如果日志写入失败，不中断主流程，只输出告警
    console.error('写入日志失败:', error)
  }
}

export function createLogger(scope: string) {
  return {
    info(message: string, context?: LogContext) {
      writeLog('INFO', scope, message, context)
    },
    warn(message: string, context?: LogContext) {
      writeLog('WARN', scope, message, context)
    },
    error(message: string, context?: LogContext & { error?: unknown }) {
      const errorContext = context?.error
        ? {
            ...context,
            error:
              context.error instanceof Error
                ? { message: context.error.message, stack: context.error.stack }
                : context.error
          }
        : context
      writeLog('ERROR', scope, message, errorContext)
    }
  }
}

// 默认导出一个基础 logger，方便快速引用
export const logger = createLogger('app')

export function getLogFilePath(): string {
  return resolveLogFile()
}
