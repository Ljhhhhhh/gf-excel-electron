import ExcelJS from 'exceljs'
import type { Workbook, Row, CellValue, Worksheet } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule } from './types'

interface BankCommonParseOptions extends ParseOptions {
  ledgerSheet?: string | number
  loanSheet?: string | number
  ledgerDataStartRow?: number
  loanDataStartRow?: number
}

interface LedgerRowRecord {
  financeId: string
  loanDate: Date | null
  balance: number | null
  repaymentDate: Date | null
  repaymentAmount: number | null
  loanSerial: string
  repaymentSerial: string
  counterpartyName: string
  applicantName: string
}

interface LoanRowRecord {
  loanDate: Date | null
  amount: number | null
}

interface BankCommonParsedData {
  ledgerRows: LedgerRowRecord[]
  loanRows: LoanRowRecord[]
}

interface BankCommonUserInput {
  period: string
}

const DEFAULT_LEDGER_SHEET = '融资及还款明细'
const DEFAULT_LOAN_SHEET = 'sheet1'
const LEDGER_START_ROW = 3
const LOAN_START_ROW = 2
const ROW_YIELD_INTERVAL = 3000
const LOAN_SOURCE_ID = 'loanDetail'

const LEDGER_COLS = {
  applicantName: 8,
  counterpartyName: 9,
  balance: 13,
  loanSerial: 16,
  financeId: 20,
  loanDate: 23,
  repaymentDate: 31,
  repaymentSerial: 33,
  repaymentAmount: 34
}

const LOAN_COLS = {
  loanDate: 16,
  amount: 49
}

// ========== 解析器实现 ==========

export async function parseWorkbook(
  workbook: Workbook,
  parseOptions?: BankCommonParseOptions
): Promise<BankCommonParsedData> {
  const options = {
    ledgerSheet: parseOptions?.ledgerSheet ?? DEFAULT_LEDGER_SHEET,
    loanSheet: parseOptions?.loanSheet ?? DEFAULT_LOAN_SHEET,
    ledgerDataStartRow: parseOptions?.ledgerDataStartRow ?? LEDGER_START_ROW,
    loanDataStartRow: parseOptions?.loanDataStartRow ?? LOAN_START_ROW
  }

  const ledgerSheet =
    typeof options.ledgerSheet === 'number'
      ? workbook.worksheets[options.ledgerSheet]
      : workbook.getWorksheet(options.ledgerSheet)
  if (!ledgerSheet) {
    throw new Error(`[bankCommon] 未找到台账工作表: ${options.ledgerSheet}`)
  }

  // 通过 excelToData 注入的额外数据源（放款明细）
  // 流式模式同样依赖预先载入的放款明细 workbook
  const extraSource = parseOptions?.extraSources?.[LOAN_SOURCE_ID]
  if (!extraSource) {
    throw new Error('[bankCommon] 缺少放款明细数据源')
  }

  const loanSheet =
    typeof options.loanSheet === 'number'
      ? extraSource.workbook.worksheets[options.loanSheet]
      : extraSource.workbook.getWorksheet(options.loanSheet)
  if (!loanSheet) {
    throw new Error(`[bankCommon] 未找到放款明细工作表: ${options.loanSheet}`)
  }

  const ledgerRows = collectLedgerRowsFromWorksheet(ledgerSheet, options.ledgerDataStartRow)
  const loanRows = collectLoanRowsFromWorksheet(loanSheet, options.loanDataStartRow)
  console.log(`[bankCommon] 解析完成，台账行: ${ledgerRows.length}，放款明细行: ${loanRows.length}`)
  return { ledgerRows, loanRows }
}

export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: BankCommonParseOptions
): Promise<BankCommonParsedData> {
  const options = {
    ledgerSheet: parseOptions?.ledgerSheet ?? DEFAULT_LEDGER_SHEET,
    loanSheet: parseOptions?.loanSheet ?? DEFAULT_LOAN_SHEET,
    ledgerDataStartRow: parseOptions?.ledgerDataStartRow ?? LEDGER_START_ROW,
    loanDataStartRow: parseOptions?.loanDataStartRow ?? LOAN_START_ROW
  }

  const extraSource = parseOptions?.extraSources?.[LOAN_SOURCE_ID]
  if (!extraSource) {
    throw new Error('[bankCommon] 缺少放款明细数据源')
  }

  const loanSheet =
    typeof options.loanSheet === 'number'
      ? extraSource.workbook.worksheets[options.loanSheet]
      : extraSource.workbook.getWorksheet(options.loanSheet)
  if (!loanSheet) {
    throw new Error(`[bankCommon] 未找到放款明细工作表: ${options.loanSheet}`)
  }

  const loanRows = collectLoanRowsFromWorksheet(loanSheet, options.loanDataStartRow)
  const ledgerRows = await collectLedgerRowsStream(filePath, options)
  console.log(
    `[bankCommon] 流式解析完成，台账行: ${ledgerRows.length}，放款明细行: ${loanRows.length}`
  )
  return { ledgerRows, loanRows }
}

async function collectLedgerRowsStream(
  filePath: string,
  options: { ledgerSheet: string | number; ledgerDataStartRow: number }
): Promise<LedgerRowRecord[]> {
  const ledgerRows: LedgerRowRecord[] = []
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    worksheets: 'emit'
  })

  let sheetIndex = 0
  const targetName = typeof options.ledgerSheet === 'string' ? options.ledgerSheet : null
  const targetIndex = typeof options.ledgerSheet === 'number' ? options.ledgerSheet : null
  let processed = false

  for await (const worksheetReader of workbookReader) {
    const sheetName = (worksheetReader as any).name
    const isTarget =
      (targetName && sheetName === targetName) ||
      (targetIndex !== null && sheetIndex === targetIndex)
    if (!isTarget) {
      sheetIndex++
      continue
    }

    let rowIndex = 0
    for await (const row of worksheetReader) {
      rowIndex++
      if (rowIndex < options.ledgerDataStartRow) {
        continue
      }
      const parsed = extractLedgerRow(row as Row)
      if (parsed) {
        ledgerRows.push(parsed)
      }
      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
    }
    processed = true
    break
  }

  if (!processed) {
    throw new Error(`[bankCommon] 无法找到台账工作表: ${options.ledgerSheet}`)
  }

  return ledgerRows
}

function collectLedgerRowsFromWorksheet(worksheet: Worksheet, startRow: number): LedgerRowRecord[] {
  const rows: LedgerRowRecord[] = []
  const endRow = worksheet.rowCount
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    if (!row.hasValues) continue
    const parsed = extractLedgerRow(row)
    if (parsed) {
      rows.push(parsed)
    }
  }
  return rows
}

function collectLoanRowsFromWorksheet(worksheet: Worksheet, startRow: number): LoanRowRecord[] {
  const rows: LoanRowRecord[] = []
  const endRow = worksheet.rowCount
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    if (!row.hasValues) continue
    const loanDate = parseDateValue(row.getCell(LOAN_COLS.loanDate).value)
    const amount = toNumber(row.getCell(LOAN_COLS.amount).value)
    if (loanDate || (amount !== null && amount !== 0)) {
      rows.push({ loanDate, amount })
    }
  }
  return rows
}

function extractLedgerRow(row: Row): LedgerRowRecord | null {
  const financeId = normalizeString(row.getCell(LEDGER_COLS.financeId).value)
  const loanDate = parseDateValue(row.getCell(LEDGER_COLS.loanDate).value)
  const balance = toNumber(row.getCell(LEDGER_COLS.balance).value)
  const repaymentDate = parseDateValue(row.getCell(LEDGER_COLS.repaymentDate).value)
  const repaymentAmount = toNumber(row.getCell(LEDGER_COLS.repaymentAmount).value)
  const loanSerial = normalizeString(row.getCell(LEDGER_COLS.loanSerial).value)
  const repaymentSerial = normalizeString(row.getCell(LEDGER_COLS.repaymentSerial).value)
  const counterpartyName = normalizeString(row.getCell(LEDGER_COLS.counterpartyName).value)
  const applicantName = normalizeString(row.getCell(LEDGER_COLS.applicantName).value)

  if (
    !financeId &&
    !loanSerial &&
    !repaymentSerial &&
    balance === null &&
    repaymentAmount === null &&
    !counterpartyName &&
    !applicantName &&
    !loanDate &&
    !repaymentDate
  ) {
    return null
  }

  return {
    financeId,
    loanDate,
    balance,
    repaymentDate,
    repaymentAmount,
    loanSerial,
    repaymentSerial,
    counterpartyName,
    applicantName
  }
}

// ========== ExcelJS 渲染器 ==========

async function renderWithExcelJS(
  parsedData: unknown,
  userInput: BankCommonUserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput || !userInput.period) {
    throw new Error('[bankCommon] 缺少统计周期参数')
  }

  const data = parsedData as BankCommonParsedData
  const period = parsePeriodRange(userInput.period)

  const currentAmount = (calculateLoanAmount(data.loanRows, period.start, period.end) ?? 0) / 10000
  const currentCount = calculateCurrentLoanCount(data.ledgerRows, period.start, period.end)
  const settlementAmount =
    (calculateRepaymentAmount(data.ledgerRows, period.start, period.end) ?? 0) / 10000
  const settlementCount = calculateSettlementCount(data.ledgerRows, period.start, period.end)
  const endingBalanceAmount = (calculateEndingBalance(data.ledgerRows) ?? 0) / 10000
  const endingFinanceCount = calculateEndingFinanceCount(data.ledgerRows)
  const endingPayerCount = calculateEndingCounterpartyCount(data.ledgerRows)
  const endingBorrowerCount = calculateEndingApplicantCount(data.ledgerRows)

  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('银行常见数据')
  sheet.columns = [{ width: 30 }, { width: 32 }]

  sheet.getCell('A1').value = '指标'
  sheet.getCell('A1').font = { bold: true }
  sheet.getCell('B1').value = period.raw
  sheet.getCell('B1').alignment = { horizontal: 'left', vertical: 'middle' }

  const metrics = [
    { label: '当期发放资产规模（万元）', value: currentAmount, format: '#,##0.00' },
    { label: '当期发放笔数（笔）', value: currentCount, format: '#,##0' },
    { label: '当期结清保理金额（万元）', value: settlementAmount, format: '#,##0.00' },
    { label: '当期结清笔数（笔）', value: settlementCount, format: '#,##0' },
    { label: '期末保理余额（万元）', value: endingBalanceAmount, format: '#,##0.00' },
    { label: '期末保理笔数（笔）', value: endingFinanceCount, format: '#,##0' },
    { label: '期末付款人数量（家）', value: endingPayerCount, format: '#,##0' },
    { label: '期末融资方数量（家）', value: endingBorrowerCount, format: '#,##0' }
  ]

  metrics.forEach((metric, index) => {
    const row = sheet.getRow(index + 2)
    row.getCell(1).value = metric.label
    row.getCell(2).value = metric.value ?? 0
    row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' }
    row.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' }
    row.getCell(2).numFmt = metric.format
  })

  await workbook.xlsx.writeFile(outputPath)
  console.log(`[bankCommon] Excel 报表已生成: ${outputPath}`)
}

// ========== 计算函数 ==========

function calculateLoanAmount(rows: LoanRowRecord[], start: Date, end: Date): number {
  return rows
    .filter((row) => row.loanDate && isWithinRange(row.loanDate, start, end))
    .reduce((sum, row) => sum + (row.amount ?? 0), 0)
}

function calculateCurrentLoanCount(rows: LedgerRowRecord[], start: Date, end: Date): number {
  const set = new Set<string>()
  rows.forEach((row) => {
    if (!row.financeId || !row.loanDate) return
    if (isWithinRange(row.loanDate, start, end)) {
      set.add(row.financeId)
    }
  })
  return set.size
}

function calculateRepaymentAmount(rows: LedgerRowRecord[], start: Date, end: Date): number {
  return rows
    .filter((row) => row.repaymentDate && isWithinRange(row.repaymentDate, start, end))
    .reduce((sum, row) => sum + (row.repaymentAmount ?? 0), 0)
}

function calculateSettlementCount(rows: LedgerRowRecord[], start: Date, end: Date): number {
  const groupMap = new Map<string, { allZero: boolean; loanSerials: Set<string> }>()
  rows.forEach((row) => {
    if (!row.financeId) return
    let group = groupMap.get(row.financeId)
    if (!group) {
      group = { allZero: true, loanSerials: new Set() }
      groupMap.set(row.financeId, group)
    }
    if (row.balance === null || row.balance !== 0) {
      group.allZero = false
    }
    if (row.loanSerial) {
      group.loanSerials.add(row.loanSerial)
    }
  })

  const repaymentDateMap = new Map<string, Date>()
  rows.forEach((row) => {
    if (!row.repaymentSerial || !row.repaymentDate) return
    const current = repaymentDateMap.get(row.repaymentSerial)
    if (!current || row.repaymentDate > current) {
      repaymentDateMap.set(row.repaymentSerial, row.repaymentDate)
    }
  })

  let count = 0
  groupMap.forEach((group) => {
    if (!group.allZero || group.loanSerials.size === 0) return
    let maxDate: Date | null = null
    group.loanSerials.forEach((serial) => {
      const candidate = repaymentDateMap.get(serial)
      if (candidate && (!maxDate || candidate > maxDate)) {
        maxDate = candidate
      }
    })
    if (maxDate && isWithinRange(maxDate, start, end)) {
      count++
    }
  })
  return count
}

function calculateEndingBalance(rows: LedgerRowRecord[]): number {
  return rows.reduce((sum, row) => sum + (row.balance ?? 0), 0)
}

function hasValidBalance(row: LedgerRowRecord): boolean {
  return row.balance !== null
}

function calculateEndingFinanceCount(rows: LedgerRowRecord[]): number {
  const set = new Set<string>()
  rows.forEach((row) => {
    if (!row.financeId || !hasValidBalance(row)) return
    set.add(row.financeId)
  })
  return set.size
}

function calculateEndingCounterpartyCount(rows: LedgerRowRecord[]): number {
  const set = new Set<string>()
  rows.forEach((row) => {
    if (!row.counterpartyName || !hasValidBalance(row)) return
    set.add(row.counterpartyName)
  })
  return set.size
}

function calculateEndingApplicantCount(rows: LedgerRowRecord[]): number {
  const set = new Set<string>()
  rows.forEach((row) => {
    if (!row.applicantName || !hasValidBalance(row)) return
    set.add(row.applicantName)
  })
  return set.size
}

// ========== 工具函数 ==========

function unwrap(value: CellValue): any {
  if (value && typeof value === 'object') {
    if ('result' in value && typeof value.result !== 'undefined') {
      return unwrap(value.result as CellValue)
    }
    if ('text' in value && typeof value.text === 'string') {
      return value.text
    }
  }
  return value
}

function isPlaceholder(value: string): boolean {
  const trimmed = value.trim()
  return trimmed === '' || trimmed === '-' || trimmed === '/'
}

function normalizeString(value: CellValue): string {
  if (value === null || typeof value === 'undefined') return ''
  const unwrapped = unwrap(value)
  if (typeof unwrapped === 'string') {
    return isPlaceholder(unwrapped) ? '' : unwrapped.trim()
  }
  if (typeof unwrapped === 'number') {
    return Number.isFinite(unwrapped) ? String(unwrapped) : ''
  }
  return ''
}

function toNumber(value: CellValue): number | null {
  if (value === null || typeof value === 'undefined') return null
  const unwrapped = unwrap(value)
  if (typeof unwrapped === 'number') {
    return Number.isFinite(unwrapped) ? unwrapped : null
  }
  if (typeof unwrapped === 'string') {
    if (isPlaceholder(unwrapped)) return null
    const normalized = unwrapped.replace(/,/g, '').trim()
    if (!normalized) return null
    const parsed = Number(normalized)
    return Number.isFinite(parsed) ? parsed : null
  }
  return null
}

function parseDateValue(value: CellValue): Date | null {
  if (value === null || typeof value === 'undefined') return null
  const unwrapped = unwrap(value)
  if (unwrapped instanceof Date) {
    return Number.isFinite(unwrapped.getTime()) ? unwrapped : null
  }
  if (typeof unwrapped === 'number') {
    if (!Number.isFinite(unwrapped)) return null
    const epoch = new Date(1899, 11, 30)
    return new Date(epoch.getTime() + unwrapped * 86400000)
  }
  if (typeof unwrapped === 'string') {
    if (isPlaceholder(unwrapped)) return null
    const parsed = new Date(unwrapped)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  return null
}

function parsePeriodRange(value: string): { start: Date; end: Date; raw: string } {
  const trimmed = String(value || '').trim()
  const match = /^([0-9]{8})\s*-\s*([0-9]{8})$/.exec(trimmed)
  if (!match) {
    throw new Error('统计周期格式应为 YYYYMMDD-YYYYMMDD')
  }
  const start = toDate(match[1])
  const end = toDate(match[2])
  if (start.getTime() > end.getTime()) {
    throw new Error('统计周期的开始日期不能晚于结束日期')
  }
  return { start, end, raw: `${formatYmd(start)}-${formatYmd(end)}` }
}

function toDate(ymd: string): Date {
  const year = Number(ymd.slice(0, 4))
  const month = Number(ymd.slice(4, 6)) - 1
  const day = Number(ymd.slice(6, 8))
  const date = new Date(year, month, day)
  if (Number.isNaN(date.getTime())) {
    throw new Error(`无效日期: ${ymd}`)
  }
  return date
}

function formatYmd(date: Date): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}${m}${d}`
}

function isWithinRange(date: Date, start: Date, end: Date): boolean {
  const time = date.getTime()
  return time >= start.getTime() && time <= end.getTime()
}

// ========== 表单规则 ==========

const periodRule: FormCreateRule = {
  type: 'Input',
  field: 'period',
  title: '统计周期',
  value: '',
  props: {
    placeholder: '例如 20240101-20240131'
  },
  validate: [
    {
      required: true,
      message: '请输入统计周期',
      trigger: 'blur'
    },
    {
      pattern: '^\\d{8}\\s*-\\s*\\d{8}$',
      message: '格式应为 YYYYMMDD-YYYYMMDD',
      trigger: 'blur'
    }
  ]
}

// ========== 模板定义 ==========

export const bankCommonTemplate: TemplateDefinition<BankCommonUserInput> = {
  meta: {
    id: 'bank-common-data',
    name: '银行常见数据',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '汇总“台账 + 放款明细”常见指标（放款、结清、余额、客户数）',
    sourceLabel: '「台账」- 融资及还款明细',
    sourceDescription: '请选择包含“融资及还款明细”工作表的台账数据文件',
    extraSources: [
      {
        id: LOAN_SOURCE_ID,
        label: '「放款明细」- sheet1',
        description: '请选择放款明细工作簿（sheet1）',
        required: true,
        supportedExts: ['xlsx']
      }
    ]
  },
  engine: 'exceljs',
  inputRule: {
    rules: [periodRule],
    options: {
      labelWidth: '120px',
      submitBtn: false,
      resetBtn: false
    },
    description: '输入统计周期（YYYYMMDD-YYYYMMDD），系统会根据台账与放款明细计算常见指标。'
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  excelRenderer: renderWithExcelJS
}
