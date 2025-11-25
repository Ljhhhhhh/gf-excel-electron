/**
 * 模板系统入口
 * 负责初始化与注册所有模板
 */

import { registerTemplate } from './registry'
// import { month1carboneTemplate } from './month1carbone'
import { month2exceljsTemplate } from './month2exceljs'
import { month4excelTemplate } from './month4excel'
import { top10CustomersTemplate } from './top10customers'
import { bankCommonTemplate } from './bankCommon'
import { f103FactoringDetailTemplate } from './f103FactoringDetail'
import { localStatisticsTemplate } from './localStatistics'
import { month1excelTemplate } from './month1excel'
import { ledgerDailyTemplate } from './ledgerDaily'
import { emailNotifyTemplate } from './emailNotify'

/**
 * 初始化模板系统
 * 在应用启动时调用，注册所有可用模板
 */
export function initTemplates(): void {
  console.log('[Templates] 开始初始化模板系统...')

  // 注册所有模板
  // registerTemplate(month1carboneTemplate) // * 改用 month1excelTemplate 解决打包后多数据源+carbone模板生成报错的问题
  registerTemplate(month1excelTemplate)
  registerTemplate(month2exceljsTemplate)
  registerTemplate(month4excelTemplate)
  registerTemplate(top10CustomersTemplate)
  registerTemplate(bankCommonTemplate)
  registerTemplate(f103FactoringDetailTemplate)
  registerTemplate(localStatisticsTemplate)
  registerTemplate(ledgerDailyTemplate)
  registerTemplate(emailNotifyTemplate)

  console.log('[Templates] 模板系统初始化完成')
}

// 导出所有公共 API
export { registerTemplate, getTemplate, listTemplates, validateTemplate } from './registry'
export * from './types'
