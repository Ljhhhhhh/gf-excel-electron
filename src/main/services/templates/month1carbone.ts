import type { Workbook } from 'exceljs'
import type { FormCreateRule, ParseOptions, TemplateDefinition } from './types'

const DEFAULT_LOAN_SHEET = 0
const DEFAULT_ASSET_SHEET = 0
const DEFAULT_LOAN_START_ROW = 2
const DEFAULT_ASSET_START_ROW = 2
const DEFAULT_MAX_ROWS = 300000
const ASSET_SOURCE_ID = 'assetDetail'

const LOAN_COLUMNS = {
  applicantName: 3, // C 列
  contractId: 14, // N 列
  actualLoanDate: 16, // P 列
  industry: 27, // AA 列
  arAssignedAmount: 44, // AR 列
  loanAmount: 49 // AW 列
} as const

const ASSET_COLUMNS = {
  transferAmount: 30 // AD 列
} as const

const INDUSTRY_LABELS = {
  infra: '基建工程',
  medical: '医药医疗',
  refactoring: '大宗商品'
} as const

interface Month1ParseOptions extends ParseOptions {
  loanSheet?: string | number
  loanDataStartRow?: number
  assetSheet?: string | number
  assetDataStartRow?: number
  maxRows?: number
}

interface LoanRowRecord {
  actualLoanDate: Date | null
  industry: string
  loanAmount: number
  arAssignedAmount: number
  applicantName: string
  contractId: string
}

interface Month1ParsedData {
  loanRows: LoanRowRecord[]
  assetTotalTransferAmount: number
}

interface Month1UserInput {
  statYear: number
  statMonth: number
  queryYear: number
  queryMonth: number
  queryDay: number
}

export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: Month1ParseOptions
): Month1ParsedData {
  const options = {
    loanSheet: parseOptions?.loanSheet ?? DEFAULT_LOAN_SHEET,
    loanDataStartRow: parseOptions?.loanDataStartRow ?? DEFAULT_LOAN_START_ROW,
    assetSheet: parseOptions?.assetSheet ?? DEFAULT_ASSET_SHEET,
    assetDataStartRow: parseOptions?.assetDataStartRow ?? DEFAULT_ASSET_START_ROW,
    maxRows: parseOptions?.maxRows ?? DEFAULT_MAX_ROWS
  }

  const loanWorksheet = resolveWorksheet(workbook, options.loanSheet, DEFAULT_LOAN_SHEET)
  if (!loanWorksheet) {
    throw new Error('[month1carbone] 无法找到放款明细工作表')
  }

  const endRow = Math.min(loanWorksheet.rowCount, options.loanDataStartRow + options.maxRows - 1)

  const loanRows: LoanRowRecord[] = []
  for (let rowIndex = options.loanDataStartRow; rowIndex <= endRow; rowIndex++) {
    const row = loanWorksheet.getRow(rowIndex)
    if (!row.hasValues) continue

    const actualLoanDate = parseExcelDate(row.getCell(LOAN_COLUMNS.actualLoanDate).value)
    const industry = normalizeString(row.getCell(LOAN_COLUMNS.industry).value)
    const loanAmount = toNumber(row.getCell(LOAN_COLUMNS.loanAmount).value)
    const arAssignedAmount = toNumber(row.getCell(LOAN_COLUMNS.arAssignedAmount).value)
    const applicantName = normalizeString(row.getCell(LOAN_COLUMNS.applicantName).value)
    const contractId = normalizeString(row.getCell(LOAN_COLUMNS.contractId).value)

    if (
      !actualLoanDate &&
      !industry &&
      !loanAmount &&
      !arAssignedAmount &&
      !applicantName &&
      !contractId
    ) {
      continue
    }

    loanRows.push({
      actualLoanDate,
      industry,
      loanAmount,
      arAssignedAmount,
      applicantName,
      contractId
    })
  }

  console.log(`[month1carbone] 已解析放款明细 ${loanRows.length} 行`)

  const assetSource = parseOptions?.extraSources?.[ASSET_SOURCE_ID]
  if (!assetSource?.workbook) {
    throw new Error('[month1carbone] 缺少资产明细数据源')
  }

  const assetWorksheet = resolveWorksheet(
    assetSource.workbook,
    options.assetSheet,
    DEFAULT_ASSET_SHEET
  )
  if (!assetWorksheet) {
    throw new Error('[month1carbone] 无法找到资产明细工作表')
  }

  let assetTotalTransferAmount = 0
  const assetEndRow = assetWorksheet.rowCount
  for (let rowIndex = options.assetDataStartRow; rowIndex <= assetEndRow; rowIndex++) {
    const row = assetWorksheet.getRow(rowIndex)
    if (!row.hasValues) continue
    assetTotalTransferAmount += toNumber(row.getCell(ASSET_COLUMNS.transferAmount).value)
  }

  console.log(`[month1carbone] 资产明细 AD 列汇总: ${assetTotalTransferAmount.toFixed(2)} 元`)

  return { loanRows, assetTotalTransferAmount }
}

function resolveWorksheet(
  workbook: Workbook,
  ref: string | number | undefined,
  fallbackIndex: number
) {
  if (typeof ref === 'number') {
    return workbook.worksheets[ref]
  }
  if (typeof ref === 'string') {
    const target = workbook.getWorksheet(ref)
    if (target) return target
  }
  return workbook.worksheets[fallbackIndex]
}

function buildReportData(parsedData: unknown, userInput?: Month1UserInput) {
  if (!userInput) {
    throw new Error('[month1carbone] 缺少报表参数')
  }

  const data = parsedData as Month1ParsedData
  if (!data || !Array.isArray(data.loanRows)) {
    throw new Error('[month1carbone] 解析数据格式不正确')
  }

  const statYear = toInteger(userInput.statYear, 'statYear')
  const statMonth = toInteger(userInput.statMonth, 'statMonth')
  const queryYear = toInteger(userInput.queryYear, 'queryYear')
  const queryMonth = toInteger(userInput.queryMonth, 'queryMonth')
  const queryDay = toInteger(userInput.queryDay, 'queryDay')

  validateYear(statYear, 'statYear')
  validateMonth(statMonth, 'statMonth')
  validateYear(queryYear, 'queryYear')
  validateMonth(queryMonth, 'queryMonth')
  validateDay(queryDay, 'queryDay')

  const queryDate = createDateOrThrow(queryYear, queryMonth, queryDay)

  const periodRows = data.loanRows.filter((row) =>
    isSameYearMonth(row.actualLoanDate, statYear, statMonth)
  )

  const periodLoanAmount = sumBy(periodRows, (row) => row.loanAmount)
  const newLoanAmount = formatAmount(periodLoanAmount, 10000, 2)
  const newArAssignedAmount = formatAmount(
    sumBy(periodRows, (row) => row.arAssignedAmount),
    10000,
    2
  )

  const newInfraRatio = calcRatio(
    sumIndustryAmount(periodRows, INDUSTRY_LABELS.infra),
    periodLoanAmount
  )
  const newMedicalRatio = calcRatio(
    sumIndustryAmount(periodRows, INDUSTRY_LABELS.medical),
    periodLoanAmount
  )
  const newRefactoringRatio = calcRatio(
    sumIndustryAmount(periodRows, INDUSTRY_LABELS.refactoring),
    periodLoanAmount
  )

  const newCoopCustomerCount = countUnique(periodRows, (row) => row.applicantName)
  const newAcceptedBusinessCount = countUnique(periodRows, (row) => row.contractId)

  const asOfRows = data.loanRows.filter((row) => isOnOrBefore(row.actualLoanDate, queryDate))
  const asOfLoanAmount = sumBy(asOfRows, (row) => row.loanAmount)

  const asOfDateSummary = {
    queryYear,
    queryMonth: formatTwoDigits(queryMonth),
    queryDay: formatTwoDigits(queryDay),
    cumLoanAmount: formatAmount(asOfLoanAmount, 10000, 2),
    infraRatio: calcRatio(sumIndustryAmount(asOfRows, INDUSTRY_LABELS.infra), asOfLoanAmount),
    medicalRatio: calcRatio(sumIndustryAmount(asOfRows, INDUSTRY_LABELS.medical), asOfLoanAmount),
    refactoringRatio: calcRatio(
      sumIndustryAmount(asOfRows, INDUSTRY_LABELS.refactoring),
      asOfLoanAmount
    )
  }

  const totalLoanAmount = sumBy(data.loanRows, (row) => row.loanAmount)
  const sinceInceptionSummary = {
    cumAcceptedBusinessCount: countUnique(data.loanRows, (row) => row.contractId),
    cumArAssignedAmount: formatAmount(data.assetTotalTransferAmount, 100000000, 2),
    cumLoanAmount: formatAmount(totalLoanAmount, 100000000, 2),
    infraRatio: calcRatio(sumIndustryAmount(data.loanRows, INDUSTRY_LABELS.infra), totalLoanAmount),
    medicalRatio: calcRatio(
      sumIndustryAmount(data.loanRows, INDUSTRY_LABELS.medical),
      totalLoanAmount
    ),
    refactoringRatio: calcRatio(
      sumIndustryAmount(data.loanRows, INDUSTRY_LABELS.refactoring),
      totalLoanAmount
    )
  }

  const monthlySummary = {
    statYear,
    statMonth: formatTwoDigits(statMonth),
    newLoanAmount,
    newInfraRatio,
    newMedicalRatio,
    newRefactoringRatio,
    newArAssignedAmount,
    newCoopCustomerCount,
    newAcceptedBusinessCount
  }

  return {
    monthlySummary,
    asOfDateSummary,
    sinceInceptionSummary
  }
}

function sumIndustryAmount(rows: LoanRowRecord[], industry: string): number {
  return sumBy(rows, (row) => (row.industry === industry ? row.loanAmount : 0))
}

function sumBy<T>(rows: T[], selector: (row: T) => number): number {
  return rows.reduce((total, row) => {
    const value = selector(row)
    return total + (Number.isFinite(value) ? value : 0)
  }, 0)
}

function countUnique<T>(rows: T[], selector: (row: T) => string): number {
  const unique = new Set<string>()
  rows.forEach((row) => {
    const value = selector(row).trim()
    if (value) {
      unique.add(value)
    }
  })
  return unique.size
}

function calcRatio(part: number, total: number): number {
  if (!total || !Number.isFinite(total) || total === 0) {
    return 0
  }
  return Number(((part / total) * 100).toFixed(2))
}

function formatAmount(value: number, divisor: number, digits: number): number {
  if (!value || !Number.isFinite(value)) {
    return 0
  }
  return Number((value / divisor).toFixed(digits))
}

function isSameYearMonth(date: Date | null, year: number, month: number): boolean {
  if (!date) return false
  return date.getFullYear() === year && date.getMonth() + 1 === month
}

function isOnOrBefore(date: Date | null, target: Date): boolean {
  if (!date) return false
  return date.getTime() <= target.getTime()
}

function createDateOrThrow(year: number, month: number, day: number): Date {
  const date = new Date(year, month - 1, day)
  if (date.getFullYear() !== year || date.getMonth() + 1 !== month || date.getDate() !== day) {
    throw new Error('[month1carbone] 截止日期不合法')
  }
  return date
}

function toInteger(value: unknown, field: string): number {
  const num = typeof value === 'string' ? Number(value) : (value as number)
  if (!Number.isFinite(num)) {
    throw new Error(`[month1carbone] 参数 ${field} 不是有效数字`)
  }
  return Math.trunc(num)
}

function validateYear(year: number, field: string): void {
  if (year < 2000 || year > 2100) {
    throw new Error(`[month1carbone] 参数 ${field} 超出允许范围`)
  }
}

function validateMonth(month: number, field: string): void {
  if (month < 1 || month > 12) {
    throw new Error(`[month1carbone] 参数 ${field} 必须在 1-12 之间`)
  }
}

function validateDay(day: number, field: string): void {
  if (day < 1 || day > 31) {
    throw new Error(`[month1carbone] 参数 ${field} 必须在 1-31 之间`)
  }
}

function formatTwoDigits(value: number): string {
  return String(value).padStart(2, '0')
}

function parseExcelDate(value: unknown): Date | null {
  if (!value) return null
  if (value instanceof Date) {
    return Number.isFinite(value.getTime()) ? value : null
  }
  if (typeof value === 'number') {
    const epoch = new Date(1899, 11, 30)
    const date = new Date(epoch.getTime() + value * 86400000)
    return Number.isFinite(date.getTime()) ? date : null
  }
  if (typeof value === 'string') {
    const parsed = new Date(value)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    if ('result' in record) {
      return parseExcelDate(record.result)
    }
    if ('text' in record) {
      return parseExcelDate(record.text)
    }
  }
  return null
}

function normalizeString(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'number') return Number.isFinite(value) ? `${value}`.trim() : ''
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    if ('text' in record) return normalizeString(record.text)
    if ('result' in record) return normalizeString(record.result)
  }
  return String(value).trim()
}

function toNumber(value: unknown): number {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  if (typeof value === 'string') {
    const cleaned = value.replace(/,/g, '').trim()
    const parsed = Number(cleaned)
    return Number.isFinite(parsed) ? parsed : 0
  }
  if (typeof value === 'object' && value) {
    const record = value as Record<string, unknown>
    if ('result' in record) return toNumber(record.result)
    if ('text' in record) return toNumber(record.text)
  }
  return 0
}

const inputRules: FormCreateRule[] = [
  {
    type: 'InputNumber',
    field: 'statYear',
    title: '月度统计年份',
    value: new Date().getFullYear(),
    props: { min: 2020, max: 2100, step: 1, placeholder: '例如 2025' },
    validate: [
      { required: true, message: '请输入年份', trigger: 'blur' },
      { type: 'number', min: 2020, max: 2100, message: '年份需在 2020-2100', trigger: 'blur' }
    ]
  },
  {
    type: 'InputNumber',
    field: 'statMonth',
    title: '月度统计月份',
    value: new Date().getMonth() + 1,
    props: { min: 1, max: 12, step: 1, placeholder: '1-12' },
    validate: [
      { required: true, message: '请输入月份', trigger: 'blur' },
      { type: 'number', min: 1, max: 12, message: '月份需在 1-12', trigger: 'blur' }
    ]
  },
  {
    type: 'InputNumber',
    field: 'queryYear',
    title: '截至日期-年',
    value: new Date().getFullYear(),
    props: { min: 2020, max: 2100, step: 1 },
    validate: [
      { required: true, message: '请输入年份', trigger: 'blur' },
      { type: 'number', min: 2020, max: 2100, message: '年份需在 2020-2100', trigger: 'blur' }
    ]
  },
  {
    type: 'InputNumber',
    field: 'queryMonth',
    title: '截至日期-月',
    value: new Date().getMonth() + 1,
    props: { min: 1, max: 12, step: 1 },
    validate: [
      { required: true, message: '请输入月份', trigger: 'blur' },
      { type: 'number', min: 1, max: 12, message: '月份需在 1-12', trigger: 'blur' }
    ]
  },
  {
    type: 'InputNumber',
    field: 'queryDay',
    title: '截至日期-日',
    value: new Date().getDate(),
    props: { min: 1, max: 31, step: 1 },
    validate: [
      { required: true, message: '请输入日期', trigger: 'blur' },
      { type: 'number', min: 1, max: 31, message: '日期需在 1-31', trigger: 'blur' }
    ]
  }
]

export const month1carboneTemplate: TemplateDefinition<Month1UserInput> = {
  meta: {
    id: 'month1carbone',
    name: '月报1模板',
    filename: 'month1carbone.xlsx',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '基于放款明细 + 资产明细生成月度、新增与累计统计概览',
    sourceLabel: '放款明细（Sheet1）',
    sourceDescription: '请选择包含放款明细 Sheet1 的工作簿',
    extraSources: [
      {
        id: ASSET_SOURCE_ID,
        label: '资产明细（Sheet1）',
        description: '请选择资产明细工作簿（Sheet1）',
        required: true,
        supportedExts: ['xlsx']
      }
    ]
  },
  engine: 'carbone',
  inputRule: {
    rules: inputRules,
    options: {
      labelWidth: '140px',
      labelPosition: 'right',
      submitBtn: false,
      resetBtn: false
    },
    example: {
      statYear: 2025,
      statMonth: 4,
      queryYear: 2025,
      queryMonth: 4,
      queryDay: 30
    },
    description: `
### 参数说明
- **月度统计年份/月份**：用于第一部分的「本月新增」指标，按 P 列日期匹配。
- **截至日期**：用于第二部分的累计统计，按 P 列小于等于该日期过滤。

### 数据来源
- 主数据源：放款明细 sheet1，需包含 C/N/P/AA/AR/AW 列。
- 额外数据源：资产明细 sheet1，需包含 AD 列（转让总金额）。

### 输出摘要
- monthlySummary：本月新增放款、行业占比、受让金额与客户/业务数。
- asOfDateSummary：截至指定日期的累计放款与行业占比。
- sinceInceptionSummary：全量历史受理笔数、受让/放款金额（亿元）及行业占比。
    `.trim()
  },
  parser: parseWorkbook,
  builder: buildReportData
}
