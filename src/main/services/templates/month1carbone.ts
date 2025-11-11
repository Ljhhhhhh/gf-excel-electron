/**
 * 示例模板：month1carbone
 * 月度报表模板，支持多表聚合
 */

import type { Workbook } from 'exceljs'
import type { TemplateDefinition, ParseOptions, ParsedData } from './types'
import { parseSimpleTable } from '../utils/xlsx'

// ========== 解析选项接口 ==========

interface Month1ParseOptions extends ParseOptions {
  /** 指定要解析的 sheet 名称或索引列表，默认解析首表 */
  sheets?: Array<string | number>
  /** 表头行索引（1-based），默认为 1 */
  headerRow?: number
  /** 数据起始行索引（1-based），默认为表头行 + 1 */
  dataStartRow?: number
  /** 最大行数限制（防止读取过多数据），默认 10000 */
  maxRows?: number
}

// ========== 解析结果接口 ==========

interface Month1ParsedData {
  /** 表格数据（多表聚合后的统一结构） */
  rows: Record<string, unknown>[]
  /** 表头字段列表 */
  headers: string[]
  /** 汇总统计信息 */
  summary?: {
    totalRows: number
    totalSheets: number
  }
}

// ========== 解析器实现 ==========

/**
 * 解析 workbook，提取表格数据
 */
export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: Month1ParseOptions
): Month1ParsedData {
  const options: Required<Month1ParseOptions> = {
    sheets: parseOptions?.sheets ?? [0], // 默认首表
    headerRow: parseOptions?.headerRow ?? 1,
    dataStartRow: parseOptions?.dataStartRow ?? 2,
    maxRows: parseOptions?.maxRows ?? 10000
  }

  const table = parseSimpleTable({
    workbook,
    sheets: options.sheets,
    headerRow: options.headerRow,
    dataStartRow: options.dataStartRow,
    maxRows: options.maxRows
  })

  return {
    rows: table.rows,
    headers: table.headers,
    summary: table.summary
  }
}

// ========== 数据构建器输入接口 ==========

/**
 * 用户输入的查询参数
 */
export interface ReportInput {
  /** 查询年份，如 2024 */
  queryYear: number
  /** 查询月份，1-12 */
  queryMonth: number
}

// ========== 数据构建器实现 ==========

/**
 * 构建 Carbone 渲染数据
 * @param parsedData 解析后的 Excel 数据
 * @param userInput 用户输入的查询参数（年份、月份）
 */
export function buildReportData(parsedData: ParsedData, userInput?: unknown): unknown {
  // 类型检查和默认值
  if (!userInput || typeof userInput !== 'object') {
    throw new Error('缺少必需的用户输入参数 (queryYear, queryMonth)')
  }

  const input = userInput as ReportInput
  const { queryYear, queryMonth } = input
  const data = parsedData as Month1ParsedData

  // 过滤出指定月份的数据行
  const targetRows = data.rows.filter((row) => {
    const dateStr = row['实际放款日期'] as string | undefined
    if (!dateStr) return false

    // 解析日期字符串（格式：YYYY-MM-DD）
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/)
    if (!match) return false

    const year = parseInt(match[1], 10)
    const month = parseInt(match[2], 10)

    return year === queryYear && month === queryMonth
  })

  // 1. d.total - 放款金额总和（万元）
  const total =
    targetRows.reduce((sum, row) => {
      const amount = parseFloat(String(row['放款金额'] || 0))
      return sum + amount
    }, 0) / 10000

  // 2. d.infrastructurePercentage - 基建工程占比
  const infrastructureAmount =
    targetRows
      .filter((row) => row['所属行业'] === '基建工程')
      .reduce((sum, row) => {
        const amount = parseFloat(String(row['放款金额'] || 0))
        return sum + amount
      }, 0) / 10000
  const infrastructurePercentage = total > 0 ? (infrastructureAmount / total) * 100 : 0

  // 3. medicinePercentage - 医药医疗占比
  const medicineAmount =
    targetRows
      .filter((row) => row['所属行业'] === '医药医疗')
      .reduce((sum, row) => {
        const amount = parseFloat(String(row['放款金额'] || 0))
        return sum + amount
      }, 0) / 10000
  const medicinePercentage = total > 0 ? (medicineAmount / total) * 100 : 0

  // 4. d.factoringPercentage - 大宗商品占比
  const factoringAmount =
    targetRows
      .filter((row) => row['所属行业'] === '大宗商品')
      .reduce((sum, row) => {
        const amount = parseFloat(String(row['放款金额'] || 0))
        return sum + amount
      }, 0) / 10000
  const factoringPercentage = total > 0 ? (factoringAmount / total) * 100 : 0

  // 5. d.accountsReceivable - 转让总金额（万元）
  const accountsReceivable =
    targetRows.reduce((sum, row) => {
      const amount = parseFloat(String(row['转让总金额'] || 0))
      return sum + amount
    }, 0) / 10000

  // 6. d.customerTotal - 保理/再保理申请人名称去重后的个数
  const uniqueCustomers = new Set<string>()
  targetRows.forEach((row) => {
    const customer = row['保理/再保理申请人名称']
    if (customer && String(customer).trim()) {
      uniqueCustomers.add(String(customer).trim())
    }
  })
  const customerTotal = uniqueCustomers.size

  // 7. d.businessCount - 保理融资申请书合同编号去重后的个数
  const uniqueContracts = new Set<string>()
  targetRows.forEach((row) => {
    const contract = row['保理融资申请书合同编号']
    if (contract && String(contract).trim()) {
      uniqueContracts.add(String(contract).trim())
    }
  })
  const businessCount = uniqueContracts.size

  // 8. 计算平均保理利率（如果有资金费报价字段）
  // 注：当前数据源暂无此字段，返回占位符
  const factoringRate = '待定'

  // 返回 Carbone 渲染数据
  // 注意：Carbone 对嵌套对象支持有限，直接使用扁平结构
  return {
    queryYear: queryYear.toString(),
    queryMonth: queryMonth.toString(),
    total: total.toFixed(2),
    infrastructurePercentage: infrastructurePercentage.toFixed(2),
    medicinePercentage: medicinePercentage.toFixed(2),
    factoringPercentage: factoringPercentage.toFixed(2),
    accountsReceivable: accountsReceivable.toFixed(2),
    customerTotal: customerTotal.toString(),
    businessCount: businessCount.toString(),
    factoringRate
  }
}

// ========== 模板定义与导出 ==========

export const month1carboneTemplate: TemplateDefinition = {
  meta: {
    id: 'month1carbone',
    name: '月度报表模板',
    filename: 'month1carbone.xlsx',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx', 'xls'],
    description: '标准月度报表，支持多表聚合'
  },
  parser: parseWorkbook,
  builder: buildReportData,
  carboneOptions: {
    lang: 'zh-cn',
    timezone: 'Asia/Shanghai'
  }
}
