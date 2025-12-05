import ExcelJS from 'exceljs'
import type { Workbook, Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule, ExtraSourceContext } from './types'
import { streamWorksheetRows } from './streamUtils'

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

/** 额外数据源 ID */
const LEDGER_SOURCE_ID = 'ledger'

/** 台账-融资及还款明细列索引（1-based） */
const LEDGER_COLUMNS = {
  financingAppNo: 20, // T 列：融资申请书编号
  loanDate: 23 // W 列：放款日期
} as const

/** 台账数据起始行 */
const LEDGER_DATA_START_ROW = 3

interface ParsedPeriod {
  period: string
  queryFormat: string
  year: number
  months: number[]
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
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? value : null
  if (typeof value === 'number') {
    // Excel 序列号纪元：1899-12-30 (UTC)，避免本地时区偏移
    const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30)
    return new Date(EXCEL_EPOCH_MS + value * 86400000)
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

  const normalized = raw.toUpperCase()
  const match = normalized.match(/^(\d{4})(0[1-9]|1[0-2])$/)
  if (!match) {
    throw new Error('统计期间格式不正确，请输入 YYYYMM')
  }

  const year = Number(match[1])
  const month = Number(match[2])
  const months = Array.from({ length: month }, (_, idx) => idx + 1)
  return {
    period: raw,
    queryFormat: `${year}年${month}月`,
    year,
    months
  }
}

function isWithinPeriod(dateValue: unknown, period: ParsedPeriod): boolean {
  const date = parseExcelDate(dateValue)
  if (!date) return false
  // 使用 UTC 方法保持与 parseExcelDate 的一致性
  if (date.getUTCFullYear() !== period.year) return false
  return period.months.includes(date.getUTCMonth() + 1)
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

/**
 * 从台账【融资及还款明细】sheet 中统计融资申请书编号去重个数
 * 筛选条件：W 列（放款日期）在用户输入年份的 1 月 1 日至输入月份（含）之间
 */
async function collectFactoringContractCountFromLedger(
  extraSource: ExtraSourceContext,
  period: ParsedPeriod
): Promise<number> {
  const sheetRef = '融资及还款明细'
  const startRow = LEDGER_DATA_START_ROW
  const uniqueAppNos = new Set<string>()

  // 模式 1：workbook 已加载
  if (extraSource.workbook) {
    const worksheet = extraSource.workbook.getWorksheet(sheetRef)
    if (!worksheet) {
      console.warn(`[localStatistics] 台账工作簿中未找到工作表: ${sheetRef}`)
      return 0
    }

    const endRow = worksheet.rowCount
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const row = worksheet.getRow(rowNum)
      if (!row.hasValues) continue

      const loanDateValue = row.getCell(LEDGER_COLUMNS.loanDate).value
      if (!isWithinPeriod(loanDateValue, period)) continue

      const appNo = normalizeString(row.getCell(LEDGER_COLUMNS.financingAppNo).value)
      if (appNo) {
        uniqueAppNos.add(appNo)
      }
    }

    console.log(
      `[localStatistics] 台账融资申请书编号统计完成（workbook 模式），去重后共 ${uniqueAppNos.size} 个`
    )
    return uniqueAppNos.size
  }

  // 模式 2：流式读取
  if (extraSource.createReader) {
    await streamWorksheetRows(
      {
        readerFactory: extraSource.createReader,
        sheet: sheetRef,
        startRow,
        rowYieldInterval: ROW_YIELD_INTERVAL
      },
      (row) => {
        if (!row.hasValues) return

        const loanDateValue = row.getCell(LEDGER_COLUMNS.loanDate).value
        if (!isWithinPeriod(loanDateValue, period)) return

        const appNo = normalizeString(row.getCell(LEDGER_COLUMNS.financingAppNo).value)
        if (appNo) {
          uniqueAppNos.add(appNo)
        }
      }
    )

    console.log(
      `[localStatistics] 台账融资申请书编号统计完成（流式模式），去重后共 ${uniqueAppNos.size} 个`
    )
    return uniqueAppNos.size
  }

  throw new Error('[localStatistics] 台账数据源未提供可用的访问方式')
}

async function buildReportData(
  parsedData: unknown,
  userInput?: LocalStatisticsInput,
  extraSources?: Record<string, ExtraSourceContext>
) {
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

  // 从台账数据源统计保理合同数量
  let factoringContractCount = 0
  const ledgerSource = extraSources?.[LEDGER_SOURCE_ID]
  if (ledgerSource) {
    factoringContractCount = await collectFactoringContractCountFromLedger(
      ledgerSource,
      parsedPeriod
    )
  } else {
    console.warn('[localStatistics] 未提供台账数据源，保理合同数量将为 0')
  }

  return {
    period: parsedPeriod.period,
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
      placeholder: '例如 202509 表示统计 2025 年 1 月至 9 月'
    },
    validate: [
      { required: true, message: '请输入统计期间', trigger: 'blur' },
      {
        pattern: '^\\d{4}(0[1-9]|1[0-2])$',
        message: '格式需为 YYYYMM，例如 202509',
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
    description: '根据输入的年月统计当年 1 月至该月份的关键指标',
    sourceLabel: '放款明细表（xlsx）',
    extraSources: [
      {
        id: LEDGER_SOURCE_ID,
        label: '台账',
        description: '请选择台账 Excel 文件（需包含【融资及还款明细】工作表）',
        required: true,
        supportedExts: ['xlsx']
      }
    ]
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
      period: '202509'
    },
    description: `
### 统计期间填写说明
- 输入 \`YYYYMM\`，例如 \`202509\` 表示 2025 年 1 月至 9 月的年初累计。
- 系统会自动将统计范围扩展为“当年 1 月”到“输入月份”之间（含输入月份），并按 P 列日期筛选数据。
    `.trim()
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  builder: async (parsedData, userInput, extraSources) =>
    buildReportData(parsedData, userInput, extraSources)
}
