import ExcelJS from 'exceljs'
import type { Workbook, Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule } from './types'

interface LocalStatisticsParseOptions extends ParseOptions {
  sheet?: string | number
  dataStartRow?: number
  maxRows?: number
}

interface LocalStatisticsParsedRow {
  actualLoanDate: unknown
  loanAmount: unknown
  enterpriseScale: unknown
  customerName: unknown
  contractId: unknown
}

interface LocalStatisticsParsedData {
  rows: LocalStatisticsParsedRow[]
}

interface LocalStatisticsInput {
  period: string
}

type SpanLabel = '月' | '季' | '年'

interface ParsedPeriod {
  period: string
  span: SpanLabel
  queryFormat: string
  year: number
  months?: number[]
}

const COLUMN_INDEX = {
  customerName: 3, // C 列
  enterpriseScale: 6, // F 列
  contractId: 14, // N 列
  actualLoanDate: 16, // P 列
  loanAmount: 49 // AW 列
} as const

const DEFAULT_PARSE_OPTIONS = {
  sheet: 0,
  dataStartRow: 2,
  maxRows: 200000
}
const ROW_YIELD_INTERVAL = 2000
const STREAM_WORKBOOK_OPTIONS = {
  sharedStrings: 'cache' as const,
  hyperlinks: 'ignore' as const,
  styles: 'ignore' as const,
  worksheets: 'emit' as const
}

export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: LocalStatisticsParseOptions
): LocalStatisticsParsedData {
  const options = {
    sheet: parseOptions?.sheet ?? DEFAULT_PARSE_OPTIONS.sheet,
    dataStartRow: parseOptions?.dataStartRow ?? DEFAULT_PARSE_OPTIONS.dataStartRow,
    maxRows: parseOptions?.maxRows ?? DEFAULT_PARSE_OPTIONS.maxRows
  }

  const worksheet =
    typeof options.sheet === 'number'
      ? workbook.worksheets[options.sheet]
      : workbook.getWorksheet(options.sheet)

  if (!worksheet) {
    throw new Error(`[localStatistics] 无法找到工作表: ${options.sheet}`)
  }

  const rows: LocalStatisticsParsedRow[] = []
  const endRow = Math.min(worksheet.rowCount, options.dataStartRow + options.maxRows - 1)

  for (let rowIndex = options.dataStartRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    const parsed = extractRow(row)
    if (parsed) {
      rows.push(parsed)
    }
  }

  console.log(`[localStatistics] 解析完成，共解析 ${rows.length} 行`)
  return { rows }
}

export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: LocalStatisticsParseOptions
): Promise<LocalStatisticsParsedData> {
  const options = {
    sheet: parseOptions?.sheet ?? DEFAULT_PARSE_OPTIONS.sheet,
    dataStartRow: parseOptions?.dataStartRow ?? DEFAULT_PARSE_OPTIONS.dataStartRow,
    maxRows: parseOptions?.maxRows ?? DEFAULT_PARSE_OPTIONS.maxRows
  }

  console.log('[localStatistics] 开始流式解析', {
    filePath,
    sheet: options.sheet,
    dataStartRow: options.dataStartRow,
    maxRows: options.maxRows
  })

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_WORKBOOK_OPTIONS)
  const rows: LocalStatisticsParsedRow[] = []

  const targetSheetName = typeof options.sheet === 'string' ? options.sheet : null
  const targetSheetIndex = typeof options.sheet === 'number' ? options.sheet : null
  const maxRowIndex = options.dataStartRow + options.maxRows - 1
  let sheetIndex = 0
  let processedSheet = false

  for await (const worksheetReader of workbookReader) {
    const sheetName = (worksheetReader as any).name
    const isTarget =
      (targetSheetName && sheetName === targetSheetName) ||
      (targetSheetIndex !== null && sheetIndex === targetSheetIndex)

    if (!isTarget) {
      sheetIndex++
      continue
    }

    let rowIndex = 0
    for await (const row of worksheetReader) {
      rowIndex++
      if (rowIndex < options.dataStartRow) {
        continue
      }
      if (rowIndex > maxRowIndex) {
        break
      }

      const parsed = extractRow(row as Row)
      if (parsed) {
        rows.push(parsed)
      }

      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
    }

    processedSheet = true
    console.log(
      `[localStatistics] 流式解析完成，共解析 ${rows.length} 行 (sheet=${sheetName ?? `#${sheetIndex}`})`
    )
    break
  }

  if (!processedSheet) {
    throw new Error(`[localStatistics] 无法找到工作表: ${options.sheet}`)
  }

  return { rows }
}

function extractRow(row: Row): LocalStatisticsParsedRow | null {
  if (!row.hasValues) {
    return null
  }

  return {
    customerName: row.getCell(COLUMN_INDEX.customerName).value,
    enterpriseScale: row.getCell(COLUMN_INDEX.enterpriseScale).value,
    contractId: row.getCell(COLUMN_INDEX.contractId).value,
    actualLoanDate: row.getCell(COLUMN_INDEX.actualLoanDate).value,
    loanAmount: row.getCell(COLUMN_INDEX.loanAmount).value
  }
}

function parseExcelDate(value: unknown): Date | null {
  if (!value) return null
  if (value instanceof Date) return value
  if (typeof value === 'number') {
    const excelEpoch = new Date(1899, 11, 30)
    return new Date(excelEpoch.getTime() + value * 86400000)
  }
  if (typeof value === 'string') {
    const parsed = new Date(value)
    return Number.isNaN(parsed.getTime()) ? null : parsed
  }
  if (typeof value === 'object' && value && 'result' in (value as Record<string, unknown>)) {
    return parseExcelDate((value as Record<string, unknown>).result)
  }
  return null
}

function toNumber(value: unknown): number {
  if (typeof value === 'number') return value
  if (typeof value === 'string') {
    const num = Number(value.replace(/,/g, '').trim())
    return Number.isFinite(num) ? num : 0
  }
  if (typeof value === 'object' && value && 'result' in (value as Record<string, unknown>)) {
    return toNumber((value as Record<string, unknown>).result)
  }
  return 0
}

function normalizeString(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'number') return value.toString().trim()
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object' && value && 'text' in (value as Record<string, unknown>)) {
    return normalizeString((value as Record<string, unknown>).text)
  }
  if (typeof value === 'object' && value && 'result' in (value as Record<string, unknown>)) {
    return normalizeString((value as Record<string, unknown>).result)
  }
  return String(value).trim()
}

function parsePeriodInput(input: string | undefined): ParsedPeriod {
  const raw = (input ?? '').trim()
  if (!raw) {
    throw new Error('统计期间不能为空')
  }

  const upper = raw.toUpperCase()

  const monthMatch = upper.match(/^(\d{4})(0[1-9]|1[0-2])$/)
  if (monthMatch) {
    const year = Number(monthMatch[1])
    const month = Number(monthMatch[2])
    return {
      period: raw,
      span: '月',
      queryFormat: `${year}年${month}月`,
      year,
      months: [month]
    }
  }

  const quarterMatch = upper.match(/^(\d{4})Q([1-4])$/)
  if (quarterMatch) {
    const year = Number(quarterMatch[1])
    const quarter = Number(quarterMatch[2]) as 1 | 2 | 3 | 4
    const quarterLabels: Record<1 | 2 | 3 | 4, string> = {
      1: '第一季度',
      2: '第二季度',
      3: '第三季度',
      4: '第四季度'
    }
    const quarterMonths: Record<1 | 2 | 3 | 4, number[]> = {
      1: [1, 2, 3],
      2: [4, 5, 6],
      3: [7, 8, 9],
      4: [10, 11, 12]
    }
    return {
      period: upper,
      span: '季',
      queryFormat: `${year}年${quarterLabels[quarter]}`,
      year,
      months: quarterMonths[quarter]
    }
  }

  const yearMatch = upper.match(/^\d{4}$/)
  if (yearMatch) {
    const year = Number(yearMatch[0])
    return {
      period: upper,
      span: '年',
      queryFormat: `${year}年`,
      year
    }
  }

  throw new Error('统计期间格式不正确，请输入 YYYYMM / YYYYQn / YYYY')
}

function isWithinPeriod(dateValue: unknown, period: ParsedPeriod): boolean {
  const date = parseExcelDate(dateValue)
  if (!date) return false
  if (date.getFullYear() !== period.year) return false
  if (!period.months) return true
  return period.months.includes(date.getMonth() + 1)
}

function sumLoanAmount(rows: LocalStatisticsParsedRow[]): number {
  return rows.reduce((total, row) => total + toNumber(row.loanAmount), 0)
}

function convertToWan(amount: number): number {
  return Number((amount / 10000).toFixed(2))
}

function isSmallEnterprise(value: unknown): boolean {
  const normalized = normalizeString(value)
  return normalized === '小型' || normalized === '微型'
}

function uniqueCount(
  rows: LocalStatisticsParsedRow[],
  selector: (row: LocalStatisticsParsedRow) => string
): number {
  const uniqueValues = new Set<string>()
  rows.forEach((row) => {
    const value = selector(row)
    if (value) {
      uniqueValues.add(value)
    }
  })
  return uniqueValues.size
}

function buildReportData(parsedData: unknown, userInput?: LocalStatisticsInput) {
  if (!userInput) {
    throw new Error('缺少统计期间参数 period')
  }

  const { period } = userInput
  const parsedPeriod = parsePeriodInput(period)
  const data = parsedData as LocalStatisticsParsedData
  if (!data || !Array.isArray(data.rows)) {
    throw new Error('[localStatistics] 解析结果格式不正确')
  }

  const scopedRows = data.rows.filter((row) => isWithinPeriod(row.actualLoanDate, parsedPeriod))

  const totalBusinessAmount = convertToWan(sumLoanAmount(scopedRows))
  const agriSmallRows = scopedRows.filter((row) => isSmallEnterprise(row.enterpriseScale))
  const agriSmallAmount = convertToWan(sumLoanAmount(agriSmallRows))
  const servedCustomerCount = uniqueCount(scopedRows, (row) => normalizeString(row.customerName))
  const factoringContractCount = uniqueCount(scopedRows, (row) => normalizeString(row.contractId))

  return {
    period: parsedPeriod.period,
    span: parsedPeriod.span,
    queryFormat: parsedPeriod.queryFormat,
    summary: {
      total_business_amount: totalBusinessAmount,
      agri_small_business_amount: agriSmallAmount,
      served_customer_count: servedCustomerCount
    },
    factoring: {
      financing_factoring_business_amount: totalBusinessAmount,
      agri_small_factoring_business_amount: agriSmallAmount,
      factoring_contract_count: factoringContractCount
    }
  }
}

const inputRules: FormCreateRule[] = [
  {
    type: 'Input',
    field: 'period',
    title: '统计期间',
    value: '',
    props: {
      placeholder: '例如 202501 / 2025Q1 / 2025'
    },
    validate: [
      { required: true, message: '请输入统计期间', trigger: 'blur' },
      {
        pattern: '^\\d{4}((0[1-9]|1[0-2])|([Qq][1-4]))?$',
        message: '格式需为 YYYYMM / YYYYQn / YYYY',
        trigger: 'blur'
      }
    ]
  }
]

export const localStatisticsTemplate: TemplateDefinition<LocalStatisticsInput> = {
  meta: {
    id: 'localStatistics',
    name: '地方金融组织融资借贷类（商业保理）统计表',
    filename: 'localStatistics.xlsx',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '根据 period 条件统计总表与保理表关键指标，支持月度 / 季度 / 年度的快速切换',
    sourceLabel: '放款明细表（xlsx）'
  },
  engine: 'carbone',
  inputRule: {
    rules: inputRules,
    options: {
      labelWidth: '100px',
      labelPosition: 'right',
      submitBtn: false,
      resetBtn: false
    },
    example: {
      period: '2025Q1'
    },
    description: `
### 统计期间填写说明
- **月度**：输入 \`YYYYMM\`，例如 \`202501\` 表示 2025 年 1 月。
- **季度**：输入 \`YYYYQn\`，例如 \`2025Q1\` 表示 2025 年第 1 季度。
- **年度**：输入 \`YYYY\`，例如 \`2025\` 表示 2025 年全年。

系统会根据 period 推导 \`span\`（月/季/年）与 \`queryFormat\`（中文展示），并统一按 P 列日期所属期间筛选。
    `.trim()
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  builder: buildReportData
}
