/**
 * 月报2模板：month2carbone
 * 统计本年1月到指定月的三个行业（基建工程、医药医疗、再保理）新增放款数据
 * 渐进式表格：第1组展示1月，第2组展示1-2月，...，第N组展示1-N月
 */

import type { Workbook } from 'exceljs'
import type { TemplateDefinition, ParseOptions, ParsedData, FormCreateRule } from './types'
import { parseSimpleTable } from '../utils/xlsx'

// ========== 解析选项接口 ==========

interface Month2ParseOptions extends ParseOptions {
  /** 指定要解析的 sheet 名称或索引列表，默认解析首表 */
  sheets?: Array<string | number>
  /** 表头行索引（1-based），默认为 1 */
  headerRow?: number
  /** 数据起始行索引（1-based），默认为表头行 + 1 */
  dataStartRow?: number
  /** 最大行数限制，默认 10000 */
  maxRows?: number
}

// ========== 解析结果接口 ==========

interface Month2ParsedData {
  /** 表格数据 */
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
  parseOptions?: Month2ParseOptions
): Month2ParsedData {
  const options: Required<Month2ParseOptions> = {
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
export interface Month2CarboneInput {
  /** 查询年份，如 2025 */
  queryYear: number
  /** 查询截止月份，1-12（表示统计1月到该月的数据） */
  queryMonth: number
}

// ========== formCreate 规则定义 ==========

/**
 * 月报2模板的 formCreate 规则
 */
const inputRules: FormCreateRule[] = [
  {
    type: 'InputNumber',
    field: 'queryYear',
    title: '查询年份',
    value: new Date().getFullYear(),
    props: {
      placeholder: '请输入年份',
      min: 1900,
      max: 2100,
      step: 1,
      controlsPosition: 'right'
    },
    validate: [
      { required: true, message: '请输入查询年份', trigger: 'blur' },
      {
        type: 'number',
        min: 1900,
        max: 2100,
        message: '年份必须在1900-2100之间',
        trigger: 'blur'
      }
    ]
  },
  {
    type: 'Select',
    field: 'queryMonth',
    title: '查询截止月份',
    value: new Date().getMonth() + 1,
    options: [
      { label: '1月', value: 1 },
      { label: '2月', value: 2 },
      { label: '3月', value: 3 },
      { label: '4月', value: 4 },
      { label: '5月', value: 5 },
      { label: '6月', value: 6 },
      { label: '7月', value: 7 },
      { label: '8月', value: 8 },
      { label: '9月', value: 9 },
      { label: '10月', value: 10 },
      { label: '11月', value: 11 },
      { label: '12月', value: 12 }
    ],
    props: {
      placeholder: '请选择截止月份',
      clearable: true
    },
    validate: [
      { required: true, message: '请选择查询截止月份', trigger: 'change' },
      { type: 'number', min: 1, max: 12, message: '月份必须在1-12之间', trigger: 'change' }
    ]
  }
]

// ========== 数据构建器实现 ==========

/**
 * 构建 Carbone 渲染数据
 * @param parsedData 解析后的 Excel 数据
 * @param userInput 用户输入的查询参数（年份、截止月份）
 */
export function buildReportData(parsedData: ParsedData, userInput?: unknown): unknown {
  // 类型检查和默认值
  if (!userInput || typeof userInput !== 'object') {
    throw new Error('缺少必需的用户输入参数 (queryYear, queryMonth)')
  }

  const input = userInput as Month2CarboneInput
  const { queryYear, queryMonth } = input
  const data = parsedData as Month2ParsedData

  console.log(`[month2carbone] 开始构建报表数据: ${queryYear}年1月-${queryMonth}月`)

  // 定义三个行业
  const industries = ['基建工程', '医药医疗', '再保理']

  // 1. 过滤本年1月到指定月的数据
  const filteredRows = data.rows.filter((row) => {
    const dateStr = row['实际放款日期'] as string | undefined
    if (!dateStr) return false

    // 解析日期字符串（格式：YYYY-MM-DD）
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/)
    if (!match) return false

    const year = parseInt(match[1], 10)
    const month = parseInt(match[2], 10)

    return year === queryYear && month >= 1 && month <= queryMonth
  })

  console.log(`[month2carbone] 过滤后数据行数: ${filteredRows.length}`)

  // 2. 按月份+行业分组聚合放款金额（万元）
  // monthlyData[month][industry] = amount
  const monthlyData = new Map<number, Map<string, number>>()

  // 初始化所有月份和行业的数据为 0
  for (let month = 1; month <= queryMonth; month++) {
    const industryMap = new Map<string, number>()
    industries.forEach((industry) => industryMap.set(industry, 0))
    monthlyData.set(month, industryMap)
  }

  // 聚合数据
  filteredRows.forEach((row) => {
    const dateStr = row['实际放款日期'] as string
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/)
    if (!match) return

    const month = parseInt(match[2], 10)
    const industry = String(row['所属行业'] || '').trim()
    const amount = parseFloat(String(row['放款金额'] || 0))

    // 只统计三个指定行业
    if (!industries.includes(industry)) return

    const currentAmount = monthlyData.get(month)?.get(industry) || 0
    monthlyData.get(month)?.set(industry, currentAmount + amount)
  })

  // 转换为万元
  for (let month = 1; month <= queryMonth; month++) {
    const industryMap = monthlyData.get(month)!
    industries.forEach((industry) => {
      const amount = industryMap.get(industry)!
      industryMap.set(industry, amount / 10000)
    })
  }

  // 3. 构建固定12列的数据结构（m1到m12）
  // Carbone 无法处理嵌套数组索引，所以使用扁平化结构
  const rows = industries.map((industry) => {
    const row: Record<string, string> = { industry }
    // 生成 m1 到 m12 的字段
    for (let m = 1; m <= 12; m++) {
      if (m <= queryMonth) {
        const amount = monthlyData.get(m)?.get(industry) || 0
        row[`m${m}`] = amount.toFixed(2)
      } else {
        row[`m${m}`] = '' // 超出查询月份的设为空字符串，避免显示多余的0.00
      }
    }
    return row
  })

  // 构建合计行数据（固定12列）
  const total: Record<string, string> = {}
  for (let m = 1; m <= 12; m++) {
    if (m <= queryMonth) {
      let monthTotal = 0
      industries.forEach((industry) => {
        monthTotal += monthlyData.get(m)?.get(industry) || 0
      })
      total[`m${m}`] = monthTotal.toFixed(2)
    } else {
      total[`m${m}`] = '' // 空字符串
    }
  }

  console.log(`[month2carbone] 构建完成，共 ${queryMonth} 个月，${rows.length} 个行业`)

  // 返回 Carbone 渲染数据（扁平化结构）
  const result = {
    rows,
    total
  }

  // 输出详细的数据结构用于调试
  console.log('[month2carbone] 返回的数据结构（扁平化）:')
  console.log('[month2carbone] 完整数据:', JSON.stringify(result, null, 2))

  return result
}

// ========== 模板定义与导出 ==========

export const month2carboneTemplate: TemplateDefinition<Month2CarboneInput> = {
  meta: {
    id: 'month2carbone',
    name: '月报2 - 行业月度放款统计',
    filename: 'month2carbone.xlsx',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx', 'xls'],
    description: '统计本年1月到指定月的三个行业（基建工程、医药医疗、再保理）新增放款数据'
  },
  inputRule: {
    rules: inputRules,
    options: {
      labelWidth: '120px',
      labelPosition: 'right',
      submitBtn: false,
      resetBtn: false
    },
    example: {
      queryYear: 2025,
      queryMonth: 10
    },
    description: `
### 参数说明
- **查询年份**: 筛选指定年份的放款数据
- **查询截止月份**: 统计1月到该月份的数据（1-12）

### 使用示例
选择 "2025年 10月" 将生成一个表格，包含10列（1月到10月）

### 统计行业
- 基建工程
- 医药医疗
- 再保理
    `.trim()
  },
  parser: parseWorkbook,
  builder: buildReportData,
  carboneOptions: {
    lang: 'zh-cn',
    timezone: 'Asia/Shanghai'
  }
}
