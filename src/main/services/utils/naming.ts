/**
 * 命名策略工具
 * 负责生成报表文件名（兜底策略）
 */

import { format } from 'date-fns'

/**
 * 生成报表文件名（兜底策略）
 * 格式: <templateId>-YYYYMMDD-HHmmss.xlsx
 */
export function generateReportName(templateId: string, ext = 'xlsx'): string {
  const timestamp = format(new Date(), 'yyyyMMdd-HHmmss')
  return `${templateId}-${timestamp}.${ext}`
}

/**
 * 清理文件名中的非法字符
 * Windows 不允许: < > : " / \ | ? *
 */
export function sanitizeFilename(filename: string): string {
  return filename.replace(/[<>:"/\\|?*]/g, '_')
}
