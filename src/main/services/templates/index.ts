/**
 * 模板系统入口
 * 负责初始化与注册所有模板
 */

import { registerTemplate } from './registry'
import { month1carboneTemplate } from './month1carbone'
import { month2exceljsTemplate } from './month2exceljs'
import { basicTemplate } from './basic'
import { month4excelTemplate } from './month4excel'

/**
 * 初始化模板系统
 * 在应用启动时调用，注册所有可用模板
 */
export function initTemplates(): void {
  console.log('[Templates] 开始初始化模板系统...')

  // 注册所有模板
  registerTemplate(month1carboneTemplate)
  registerTemplate(month2exceljsTemplate)
  registerTemplate(month4excelTemplate)
  registerTemplate(basicTemplate)

  console.log('[Templates] 模板系统初始化完成')
}

// 导出所有公共 API
export { registerTemplate, getTemplate, listTemplates, validateTemplate } from './registry'
export * from './types'
