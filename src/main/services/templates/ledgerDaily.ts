/**
 * 台账报表模板：ledgerDaily
 * 融资及还款明细工作表更新（流式处理，支持十万级数据）
 *
 * 数据来源：
 * - 放款明细（主数据源）
 * - 保理融资还款明细
 * - 再保理融资还款明细
 * - 台账基线文件
 */

import ExcelJS from 'exceljs'
import type { Row, Worksheet, Workbook } from 'exceljs'
import path from 'node:path'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule, ExtraSourceContext } from './types'
import { streamWorksheetRows } from './streamUtils'
import { sanitizeFilename } from '../utils/naming'

// ========== 类型定义 ==========

interface LedgerParseOptions extends ParseOptions {
  /** 放款明细工作表索引/名称 */
  loanSheet?: string | number
  /** 还款明细工作表索引/名称 */
  repaySheet?: string | number
  /** 放款明细数据起始行 */
  loanDataStartRow?: number
  /** 还款明细数据起始行 */
  repayDataStartRow?: number
}

/**
 * 放款记录（从《放款明细》提取）
 */
interface LoanRowRecord {
  loanDate: Date | null // P列(16)：实际放款日期
  colB: any // 业务来源
  colAE: any // AE列(31)
  colAG: any // AG列(33)
  colK: any // 资产编号
  colC: any // 保理/再保理申请人名称
  colG: any // 基础交易对手方名称
  colAZ: any // AZ列(52)
  colJ: any // J列(10)
  colAA: any // 所属行业
  colAK: any // AK列(37)
  colN: any // N列(14)
  colM: any // M列(13)
  colL: any // 融资申请号
  colY: any // Y列(25)：直接投放/再保理
  colAW: any // AW列(49)
  colBC: any // BC列(55)
  colBF: any // BF列(58)
  colQ: any // Q列(17)
  colT: any // T列(20)
  colU: any // U列(21)
}

/**
 * 还款记录（从《保理融资还款明细》或《再保理融资还款明细》提取）
 */
interface RepayRowRecord {
  repayDate: Date | null // AE列(31)：还款日期
  feeType: string // AB列(28)：费用类型
  colG: any // G列(7)
  colH: any // H列(8)
  colJ: any // J列(10)
  colM: any // M列(13)
  colB: any // B列(2)
  colC: any // C列(3)
  colF: any // F列(6)
  colAE: any // AE列(31)
  colO: any // O列(15)
  colAG: any // AG列(33)：金额
  colAH: any // AH列(34)：交易银行流水号
}

/**
 * 解析后的完整数据
 */
interface LedgerParsedData {
  loans: LoanRowRecord[]
  factoringRepays: RepayRowRecord[]
  refactoringRepays: RepayRowRecord[]
  ledgerPath: string
  ledgerFilename: string
}

/**
 * 用户输入参数
 */
interface LedgerUserInput {
  /** YYYYMMDD，可包含分隔符 */
  date: string
}

// ========== 常量定义 ==========

const OUTPUT_SHEET_NAME = '融资及还款明细'
const TEMPLATE_ROW_INDEX = 10 // 第10行为样式模板行
const DEFAULT_LOAN_SHEET = 0
const DEFAULT_REPAY_SHEET = 0
const DEFAULT_LOAN_START_ROW = 2
const DEFAULT_REPAY_START_ROW = 2
const MAX_ROWS = 300000
const ROW_YIELD_INTERVAL = 3000
const FACTORING_REPAY_SOURCE_ID = 'factoringRepay'
const REFACTORING_REPAY_SOURCE_ID = 'refactoringRepay'
const LEDGER_SOURCE_ID = 'ledgerWorkbook'

/** 放款记录输出列 */
const LOAN_OUTPUT_COLUMNS = [
  'A',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Z',
  'AB',
  'AC',
  'AD'
]

/** 还款记录输出列 */
const REPAY_OUTPUT_COLUMNS = [
  'A',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Z',
  'AB',
  'AC',
  'AD',
  'AE',
  'AG',
  'AH',
  'AI',
  'AJ',
  'AK',
  'AR'
]

// ========== Step 1: 解析器实现 ==========

export async function parseWorkbook(
  workbook: Workbook,
  parseOptions?: LedgerParseOptions
): Promise<LedgerParsedData> {
  const options = resolveParseOptions(parseOptions)

  // 校验台账数据源
  const ledgerSource = parseOptions?.extraSources?.[LEDGER_SOURCE_ID]
  assertLedgerSource(ledgerSource)

  // 解析放款明细
  const loans = collectLoanRowsFromWorkbook(workbook, options.loanSheet, options.loanDataStartRow)

  // 解析保理还款明细
  const factoringRepays = await collectRepayRowsFromSource(
    parseOptions?.extraSources?.[FACTORING_REPAY_SOURCE_ID],
    options
  )

  // 解析再保理还款明细
  const refactoringRepays = await collectRepayRowsFromSource(
    parseOptions?.extraSources?.[REFACTORING_REPAY_SOURCE_ID],
    options
  )

  return {
    loans,
    factoringRepays,
    refactoringRepays,
    ledgerPath: ledgerSource.path,
    ledgerFilename: path.basename(ledgerSource.path)
  }
}

/**
 * 流式解析器（推荐用于大文件）
 */
export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: LedgerParseOptions
): Promise<LedgerParsedData> {
  const options = resolveParseOptions(parseOptions)

  const ledgerSource = parseOptions?.extraSources?.[LEDGER_SOURCE_ID]
  assertLedgerSource(ledgerSource)

  // 流式解析放款明细
  const loans = await collectLoanRowsFromStream(
    filePath,
    options.loanSheet,
    options.loanDataStartRow
  )

  // 解析保理还款明细
  const factoringRepays = await collectRepayRowsFromSource(
    parseOptions?.extraSources?.[FACTORING_REPAY_SOURCE_ID],
    options
  )

  // 解析再保理还款明细
  const refactoringRepays = await collectRepayRowsFromSource(
    parseOptions?.extraSources?.[REFACTORING_REPAY_SOURCE_ID],
    options
  )

  return {
    loans,
    factoringRepays,
    refactoringRepays,
    ledgerPath: ledgerSource.path,
    ledgerFilename: path.basename(ledgerSource.path)
  }
}

function resolveParseOptions(parseOptions?: LedgerParseOptions) {
  return {
    loanSheet: parseOptions?.loanSheet ?? DEFAULT_LOAN_SHEET,
    repaySheet: parseOptions?.repaySheet ?? DEFAULT_REPAY_SHEET,
    loanDataStartRow: parseOptions?.loanDataStartRow ?? DEFAULT_LOAN_START_ROW,
    repayDataStartRow: parseOptions?.repayDataStartRow ?? DEFAULT_REPAY_START_ROW
  }
}

function assertLedgerSource(source?: ExtraSourceContext): asserts source is ExtraSourceContext {
  if (!source?.path) {
    throw new Error('[ledgerDaily] 缺少台账文件数据源')
  }
}

// ========== 放款明细解析 ==========

function collectLoanRowsFromWorkbook(
  workbook: Workbook,
  sheetRef: string | number,
  startRow: number
): LoanRowRecord[] {
  const worksheet =
    typeof sheetRef === 'number' ? workbook.worksheets[sheetRef] : workbook.getWorksheet(sheetRef)
  if (!worksheet) {
    throw new Error(`[ledgerDaily] 未找到放款明细工作表: ${sheetRef}`)
  }
  return collectLoanRowsFromWorksheet(worksheet, startRow)
}

async function collectLoanRowsFromStream(
  filePath: string,
  sheetRef: string | number,
  startRow: number
): Promise<LoanRowRecord[]> {
  const loans: LoanRowRecord[] = []
  await streamWorksheetRows(
    {
      readerFactory: () =>
        new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
          sharedStrings: 'cache',
          hyperlinks: 'ignore',
          styles: 'ignore',
          worksheets: 'emit'
        }),
      sheet: sheetRef,
      startRow,
      maxRows: MAX_ROWS,
      rowYieldInterval: ROW_YIELD_INTERVAL
    },
    (row) => {
      const rowReader = row as Row
      if (!rowReader.hasValues) return
      const parsed = extractLoanRow(rowReader)
      if (parsed) {
        loans.push(parsed)
      }
    }
  )
  return loans
}

function collectLoanRowsFromWorksheet(worksheet: Worksheet, startRow: number): LoanRowRecord[] {
  const loans: LoanRowRecord[] = []
  const endRow = Math.min(worksheet.rowCount, MAX_ROWS + startRow)
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    if (!row.hasValues) continue
    const parsed = extractLoanRow(row)
    if (parsed) {
      loans.push(parsed)
    }
  }
  return loans
}

// ========== 还款明细解析 ==========

async function collectRepayRowsFromSource(
  source: ExtraSourceContext | undefined,
  options: ReturnType<typeof resolveParseOptions>
): Promise<RepayRowRecord[]> {
  if (!source) {
    throw new Error('[ledgerDaily] 缺少还款明细数据源')
  }

  // 优先使用 workbook（如果已加载）
  if (source.workbook) {
    const worksheet =
      typeof options.repaySheet === 'number'
        ? source.workbook.worksheets[options.repaySheet]
        : source.workbook.getWorksheet(options.repaySheet)
    if (!worksheet) {
      throw new Error(`[ledgerDaily] 未找到还款明细工作表: ${options.repaySheet}`)
    }
    return collectRepayRowsFromWorksheet(worksheet, options.repayDataStartRow)
  }

  // 使用流式读取
  if (source.createReader) {
    const repays: RepayRowRecord[] = []
    await streamWorksheetRows(
      {
        readerFactory: source.createReader,
        sheet: options.repaySheet,
        startRow: options.repayDataStartRow,
        maxRows: MAX_ROWS,
        rowYieldInterval: ROW_YIELD_INTERVAL
      },
      (row) => {
        const parsed = extractRepayRow(row as Row)
        if (parsed) {
          repays.push(parsed)
        }
      }
    )
    return repays
  }

  throw new Error('[ledgerDaily] 还款明细数据源未提供可用的访问方式')
}

function collectRepayRowsFromWorksheet(worksheet: Worksheet, startRow: number): RepayRowRecord[] {
  const repays: RepayRowRecord[] = []
  const endRow = Math.min(worksheet.rowCount, MAX_ROWS + startRow)
  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    if (!row.hasValues) continue
    const parsed = extractRepayRow(row)
    if (parsed) {
      repays.push(parsed)
    }
  }
  return repays
}

function extractLoanRow(row: Row): LoanRowRecord | null {
  const loanDate = parseExcelDate(row.getCell(16).value) // P列
  const colB = row.getCell(2).value
  const colAE = row.getCell(31).value
  const colAG = row.getCell(33).value
  const colK = row.getCell(11).value
  const colC = row.getCell(3).value
  const colG = row.getCell(7).value
  const colAZ = row.getCell(52).value
  const colJ = row.getCell(10).value
  const colAA = row.getCell(27).value
  const colAK = row.getCell(37).value
  const colN = row.getCell(14).value
  const colM = row.getCell(13).value
  const colL = row.getCell(12).value
  const colY = row.getCell(25).value
  const colAW = row.getCell(49).value
  const colBC = row.getCell(55).value
  const colBF = row.getCell(58).value
  const colQ = row.getCell(17).value
  const colT = row.getCell(20).value
  const colU = row.getCell(21).value

  // 判断空行
  if (
    !loanDate &&
    !colB &&
    !colAE &&
    !colAG &&
    !colK &&
    !colC &&
    !colG &&
    !colAZ &&
    !colJ &&
    !colAA &&
    !colAK &&
    !colN &&
    !colM &&
    !colL &&
    !colY &&
    !colAW &&
    !colBC &&
    !colBF &&
    !colQ &&
    !colT &&
    !colU
  ) {
    return null
  }

  return {
    loanDate,
    colB,
    colAE,
    colAG,
    colK,
    colC,
    colG,
    colAZ,
    colJ,
    colAA,
    colAK,
    colN,
    colM,
    colL,
    colY,
    colAW,
    colBC,
    colBF,
    colQ,
    colT,
    colU
  }
}

function extractRepayRow(row: Row): RepayRowRecord | null {
  const repayDate = parseExcelDate(row.getCell(31).value) // AE列
  const feeType = normalizeString(row.getCell(28).value) // AB列
  const colG = row.getCell(7).value
  const colH = row.getCell(8).value
  const colJ = row.getCell(10).value
  const colM = row.getCell(13).value
  const colB = row.getCell(2).value
  const colC = row.getCell(3).value
  const colF = row.getCell(6).value
  const colAE = row.getCell(31).value
  const colO = row.getCell(15).value
  const colAG = row.getCell(33).value
  const colAH = row.getCell(34).value

  // 判断空行
  if (!repayDate && !colG && !colH && !colJ && !colM && !colB && !colC && !colAG) {
    return null
  }

  return {
    repayDate,
    feeType,
    colG,
    colH,
    colJ,
    colM,
    colB,
    colC,
    colF,
    colAE,
    colO,
    colAG,
    colAH
  }
}

// ========== 工具函数 ==========

function parseExcelDate(value: any): Date | null {
  if (!value) return null
  if (value instanceof Date) {
    return Number.isFinite(value.getTime()) ? value : null
  }
  if (typeof value === 'number') {
    const epoch = new Date(1899, 11, 30)
    const parsed = new Date(epoch.getTime() + value * 86400000)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return null
    const parsed = new Date(trimmed)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  if (typeof value === 'object' && value) {
    if ('text' in value) return parseExcelDate((value as any).text)
    if ('result' in value) return parseExcelDate((value as any).result)
  }
  return null
}

function normalizeString(value: any): string {
  if (value === null || typeof value === 'undefined') return ''
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : ''
  if (value instanceof Date) return formatYmd(value)
  if (typeof value === 'object' && value) {
    if ('text' in value) return normalizeString((value as any).text)
    if ('result' in value) return normalizeString((value as any).result)
  }
  return String(value).trim()
}

function formatYmd(date: Date): string {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(
    date.getDate()
  ).padStart(2, '0')}`
}

function normalizeInputDate(raw: string): { date: Date; ymd: string } {
  const digits = String(raw ?? '').replace(/[^0-9]/g, '')
  if (digits.length !== 8) {
    throw new Error('[ledgerDaily] 日期格式需为 YYYYMMDD')
  }
  const year = Number(digits.slice(0, 4))
  const month = Number(digits.slice(4, 6))
  const day = Number(digits.slice(6, 8))
  const date = new Date(year, month - 1, day)
  if (!Number.isFinite(date.getTime()) || date.getMonth() + 1 !== month || date.getDate() !== day) {
    throw new Error('[ledgerDaily] 日期不合法，请确认输入')
  }
  const ymd = `${year}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`
  return { date, ymd }
}

function toNumber(value: any): number {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0
  }
  if (typeof value === 'string') {
    const normalized = value.replace(/,/g, '').trim()
    const parsed = Number(normalized)
    return Number.isFinite(parsed) ? parsed : 0
  }
  if (typeof value === 'object' && value) {
    if ('text' in value) return toNumber((value as any).text)
    if ('result' in value) return toNumber((value as any).result)
  }
  return 0
}

function isSameDay(date: Date | null, target: Date): boolean {
  if (!date) return false
  return (
    date.getFullYear() === target.getFullYear() &&
    date.getMonth() === target.getMonth() &&
    date.getDate() === target.getDate()
  )
}

function isAfterDay(date: Date, base: Date): boolean {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime()
  const b = new Date(base.getFullYear(), base.getMonth(), base.getDate()).getTime()
  return d > b
}

function buildConcat(left: any, right: any): string {
  const l = normalizeString(left)
  const r = normalizeString(right)
  if (!l && !r) return ''
  if (!l) return r
  if (!r) return l
  return `${l}-${r}`
}

function compareTxn(a: string, b: string): number {
  if (a === b) return 0
  if (!a) return -1
  if (!b) return 1
  return a.localeCompare(b, 'zh-CN', { numeric: true, sensitivity: 'base' })
}

// ========== Step 3: 流式渲染器实现 ==========

/**
 * 台账信息（流式扫描获取）
 */
interface LedgerInfo {
  hasTargetSheet: boolean
  lastDate: Date | null
  lastRowNumber: number
  templateRowData: TemplateRowData | null
}

/**
 * 模板行样式数据
 */
interface TemplateRowData {
  height: number | undefined
  styles: Record<string, Partial<ExcelJS.Style>>
}

/**
 * 流式复制与追加选项
 */
interface StreamCopyOptions {
  sourcePath: string
  outputPath: string
  targetSheetName: string
  templateRowIndex: number
  templateRowData: TemplateRowData | null
  rowsToAppend: Array<{ type: 'loan' | 'factoring' | 'refactoring'; record: any }>
  targetDate: Date
}

/**
 * 流式渲染函数（主入口）
 */
async function renderWithStreamWriter(
  parsedData: unknown,
  userInput: LedgerUserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput?.date) {
    throw new Error('[ledgerDaily] 缺少日期参数')
  }
  const { date: targetDate, ymd } = normalizeInputDate(userInput.date)
  const data = parsedData as LedgerParsedData

  // 第一步：流式读取台账，提取必要信息
  console.log('[ledgerDaily] 开始流式读取台账...')
  const ledgerInfo = await extractLedgerInfo(data.ledgerPath)

  if (!ledgerInfo.hasTargetSheet) {
    throw new Error(`[ledgerDaily] 未找到台账工作表: ${OUTPUT_SHEET_NAME}`)
  }

  // 检查日期是否需要更新
  if (ledgerInfo.lastDate && !isAfterDay(targetDate, ledgerInfo.lastDate)) {
    console.log(
      `[ledgerDaily] 日期 ${ymd} 不大于现有日期 ${formatYmd(ledgerInfo.lastDate)}，复制原文件`
    )
    const fs = await import('node:fs/promises')
    await fs.copyFile(data.ledgerPath, outputPath)
    return
  }

  // 构建待追加的行数据
  const rowsToAppend = buildRowsToAppend(data, targetDate)
  if (!rowsToAppend.length) {
    console.warn(`[ledgerDaily] 日期 ${ymd} 无匹配数据，复制原文件`)
    const fs = await import('node:fs/promises')
    await fs.copyFile(data.ledgerPath, outputPath)
    return
  }

  console.log(`[ledgerDaily] 准备追加 ${rowsToAppend.length} 行，开始流式写入...`)

  // 第二步：流式写入新文件
  await streamCopyAndAppend({
    sourcePath: data.ledgerPath,
    outputPath,
    targetSheetName: OUTPUT_SHEET_NAME,
    templateRowIndex: TEMPLATE_ROW_INDEX,
    templateRowData: ledgerInfo.templateRowData,
    rowsToAppend,
    targetDate
  })

  console.log(`[ledgerDaily] 流式写入完成: ${outputPath}`)
}

/**
 * 流式读取台账文件，提取元信息
 */
async function extractLedgerInfo(ledgerPath: string): Promise<LedgerInfo> {
  const result: LedgerInfo = {
    hasTargetSheet: false,
    lastDate: null,
    lastRowNumber: 0,
    templateRowData: null
  }

  const reader = new ExcelJS.stream.xlsx.WorkbookReader(ledgerPath, {
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'cache',
    worksheets: 'emit'
  })

  try {
    for await (const worksheetReader of reader) {
      const sheetName = (worksheetReader as any).name as string | undefined
      if (sheetName !== OUTPUT_SHEET_NAME) {
        continue
      }

      result.hasTargetSheet = true
      let rowNumber = 0

      for await (const row of worksheetReader) {
        rowNumber++
        const rowData = row as Row

        // 提取模板行样式
        if (rowNumber === TEMPLATE_ROW_INDEX && !result.templateRowData) {
          result.templateRowData = extractTemplateRowData(rowData)
        }

        // 查找最后一个有效数据行
        if (rowNumber > TEMPLATE_ROW_INDEX && rowHasData(rowData)) {
          result.lastRowNumber = rowNumber
          const dateValue = rowData.getCell('W').value || rowData.getCell('AE').value || undefined
          const parsedDate = parseExcelDate(dateValue)
          if (parsedDate) {
            result.lastDate = parsedDate
          }
        }
      }

      break
    }
  } catch (error) {
    console.error('[ledgerDaily] 读取台账信息失败:', error)
    throw error
  }

  return result
}

/**
 * 从模板行提取样式数据
 */
function extractTemplateRowData(templateRow: Row): TemplateRowData {
  const styleColumns = Array.from(new Set([...LOAN_OUTPUT_COLUMNS, ...REPAY_OUTPUT_COLUMNS]))
  const styles: Record<string, Partial<ExcelJS.Style>> = {}

  styleColumns.forEach((col) => {
    const cell = templateRow.getCell(col)
    if (cell.style && Object.keys(cell.style).length > 0) {
      styles[col] = JSON.parse(JSON.stringify(cell.style))
    }
  })

  return {
    height: templateRow.height,
    styles
  }
}

/**
 * 流式复制现有数据并追加新行
 */
async function streamCopyAndAppend(options: StreamCopyOptions): Promise<void> {
  const {
    sourcePath,
    outputPath,
    targetSheetName,
    templateRowIndex,
    templateRowData,
    rowsToAppend,
    targetDate
  } = options

  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: outputPath,
    useStyles: true,
    useSharedStrings: true
  })

  const reader = new ExcelJS.stream.xlsx.WorkbookReader(sourcePath, {
    sharedStrings: 'cache',
    hyperlinks: 'cache',
    styles: 'cache',
    worksheets: 'emit'
  })

  let targetSheetProcessed = false
  let sheetIndex = 0

  try {
    for await (const worksheetReader of reader) {
      const sheetName = (worksheetReader as any).name as string | undefined
      const isTargetSheet = sheetName === targetSheetName

      const outputSheet = writer.addWorksheet(sheetName || `Sheet${sheetIndex + 1}`)

      if (isTargetSheet) {
        await processTargetSheet(
          worksheetReader,
          outputSheet,
          templateRowIndex,
          templateRowData,
          rowsToAppend,
          targetDate
        )
        targetSheetProcessed = true
      } else {
        await copySheetRows(worksheetReader, outputSheet)
      }

      sheetIndex++
    }

    if (!targetSheetProcessed) {
      throw new Error(`[ledgerDaily] 未找到目标工作表: ${targetSheetName}`)
    }

    await writer.commit()
  } catch (error) {
    console.error('[ledgerDaily] 流式写入失败:', error)
    throw error
  }
}

/**
 * 处理目标工作表：复制现有行 + 追加新行
 */
async function processTargetSheet(
  worksheetReader: any,
  outputSheet: ExcelJS.Worksheet,
  templateRowIndex: number,
  templateRowData: TemplateRowData | null,
  rowsToAppend: Array<{ type: 'loan' | 'factoring' | 'refactoring'; record: any }>,
  targetDate: Date
): Promise<void> {
  let rowNumber = 0
  let lastDataRowNumber = 0
  let batchCount = 0
  const BATCH_SIZE = 500

  // 复制所有现有行
  for await (const row of worksheetReader) {
    rowNumber++
    const rowData = row as Row

    const newRow = outputSheet.addRow(rowData.values)
    newRow.height = rowData.height

    // 复制样式
    rowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber)
      if (cell.style && Object.keys(cell.style).length > 0) {
        newCell.style = cell.style
      }
    })

    await newRow.commit()

    // 记录最后一个数据行
    if (rowNumber > templateRowIndex && rowHasData(rowData)) {
      lastDataRowNumber = rowNumber
    }

    // 定期让出事件循环
    batchCount++
    if (batchCount >= BATCH_SIZE) {
      await setImmediatePromise()
      batchCount = 0
    }
  }

  console.log(`[ledgerDaily] 已复制 ${rowNumber} 行，准备追加 ${rowsToAppend.length} 行`)

  // 追加新行
  const startRowNumber = Math.max(lastDataRowNumber + 1, templateRowIndex + 1)
  const newRowValues = rowsToAppend.map((rowData) => buildRowValues(rowData))
  const mergeGroups = collectMergeGroups(rowsToAppend, startRowNumber, targetDate)
  const overrideByIndex = new Map<number, { ai: number; ar: string }>()
  mergeGroups.forEach((group) => {
    const relativeIndex = group.start - startRowNumber
    if (relativeIndex >= 0) {
      overrideByIndex.set(relativeIndex, { ai: group.sum, ar: group.ar })
    }
  })

  // 分批追加新行并应用样式
  for (let i = 0; i < newRowValues.length; i++) {
    const override = overrideByIndex.get(i)
    if (override) {
      setValue(newRowValues[i], 'AI', override.ai)
      setValue(newRowValues[i], 'AR', override.ar)
    }
    const newRow = outputSheet.addRow(newRowValues[i])

    // 应用样式
    if (templateRowData) {
      if (templateRowData.height !== undefined) {
        newRow.height = templateRowData.height
      }
      Object.entries(templateRowData.styles).forEach(([col, style]) => {
        const cell = newRow.getCell(col)
        cell.style = style
      })
    }

    await newRow.commit()

    // 定期让出事件循环
    if (i > 0 && i % BATCH_SIZE === 0) {
      await setImmediatePromise()
      console.log(`[ledgerDaily] 已追加 ${i + 1}/${newRowValues.length} 行`)
    }
  }
  console.log(`[ledgerDaily] 目标工作表处理完成，共 ${rowNumber + newRowValues.length} 行`)
}

/**
 * 复制工作表的所有行（非目标工作表）
 */
async function copySheetRows(worksheetReader: any, outputSheet: ExcelJS.Worksheet): Promise<void> {
  let rowCount = 0
  const BATCH_SIZE = 1000

  for await (const row of worksheetReader) {
    const rowData = row as Row
    const newRow = outputSheet.addRow(rowData.values)
    newRow.height = rowData.height

    // 复制样式
    rowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber)
      if (cell.style && Object.keys(cell.style).length > 0) {
        newCell.style = cell.style
      }
    })

    await newRow.commit()
    rowCount++

    if (rowCount % BATCH_SIZE === 0) {
      await setImmediatePromise()
    }
  }
}

/**
 * 构建待追加的行数据
 */
function buildRowsToAppend(data: LedgerParsedData, targetDate: Date) {
  const rows: Array<{ type: 'loan' | 'factoring' | 'refactoring'; record: any }> = []

  // 1. 收集当天放款
  data.loans.forEach((loan) => {
    if (loan.loanDate && isSameDay(loan.loanDate, targetDate)) {
      rows.push({ type: 'loan', record: loan })
    }
  })

  // 2. 收集保理回款（本金，按流水号排序）
  const factoring = data.factoringRepays
    .filter(
      (row) => row.repayDate && isSameDay(row.repayDate, targetDate) && row.feeType === '本金'
    )
    .sort((a, b) => compareTxn(normalizeString(a.colAH), normalizeString(b.colAH)))
  factoring.forEach((row) => rows.push({ type: 'factoring', record: row }))

  // 3. 收集再保理回款（本金，按流水号排序）
  const refactoring = data.refactoringRepays
    .filter(
      (row) => row.repayDate && isSameDay(row.repayDate, targetDate) && row.feeType === '本金'
    )
    .sort((a, b) => compareTxn(normalizeString(a.colAH), normalizeString(b.colAH)))
  refactoring.forEach((row) => rows.push({ type: 'refactoring', record: row }))

  return rows
}

/**
 * 构建单行的值
 */
function buildRowValues(rowData: {
  type: 'loan' | 'factoring' | 'refactoring'
  record: any
}): ExcelJS.RowValues {
  if (rowData.type === 'loan') {
    const loan = rowData.record as LoanRowRecord
    const values: ExcelJS.RowValues = []
    setValue(values, 'A', { formula: 'ROW()-2' })
    setValue(values, 'D', loan.colB)
    setValue(values, 'E', loan.colAE)
    setValue(values, 'F', loan.colAG)
    setValue(values, 'G', loan.colK)
    setValue(values, 'H', loan.colC)
    setValue(values, 'I', loan.colG)
    setValue(values, 'J', loan.colAZ)
    setValue(values, 'K', loan.colJ)
    setValue(values, 'L', loan.colAA)
    setValue(values, 'N', loan.colAK)
    setValue(values, 'O', buildConcat(loan.colK, loan.colN))
    setValue(values, 'P', loan.colM)
    setValue(values, 'Q', loan.colL)
    setValue(values, 'R', loan.colAK)
    setValue(values, 'S', normalizeString(loan.colY) === '直接投放' ? '保理' : '再保理')
    setValue(values, 'T', loan.colN)
    setValue(values, 'U', loan.colAW)
    setValue(values, 'V', loan.colAW)
    setValue(values, 'W', loan.loanDate ?? loan.colAE ?? null)
    setValue(values, 'X', loan.colBC)
    setValue(values, 'Z', loan.colBF)
    setValue(values, 'AB', loan.colQ)
    setValue(values, 'AC', loan.colT)
    setValue(values, 'AD', loan.colU)
    return values
  }

  const repay = rowData.record as RepayRowRecord
  const isFactoring = rowData.type === 'factoring'
  const values: ExcelJS.RowValues = []
  setValue(values, 'A', { formula: 'ROW()-2' })
  setValue(values, 'D', repay.colG)
  setValue(values, 'E', repay.colH)
  setValue(values, 'F', repay.colJ)
  setValue(values, 'G', repay.colM)
  setValue(values, 'H', repay.colB)
  setValue(values, 'I', repay.colC)
  setValue(values, 'J', '/')
  setValue(values, 'K', '/')
  setValue(values, 'L', repay.colF)
  setValue(values, 'N', null)
  setValue(values, 'O', null)
  setValue(values, 'P', null)
  setValue(values, 'Q', null)
  setValue(values, 'R', null)
  setValue(values, 'S', null)
  setValue(values, 'T', null)
  setValue(values, 'U', null)
  setValue(values, 'V', null)
  setValue(values, 'W', null)
  setValue(values, 'X', null)
  setValue(values, 'Z', null)
  setValue(values, 'AB', null)
  setValue(values, 'AC', null)
  setValue(values, 'AD', null)
  setValue(values, 'AE', repay.colAE)
  setValue(values, 'AG', repay.colO)
  setValue(values, 'AH', repay.colAG)
  setValue(values, 'AI', repay.colAG)
  setValue(values, 'AJ', repay.colAE)
  setValue(values, 'AK', isFactoring ? '保理' : '再保理')
  setValue(values, 'AR', repay.colAH)
  return values
}

/**
 * 收集合并分组（AI列需要合并）
 */
function collectMergeGroups(
  rowsToAppend: Array<{ type: 'loan' | 'factoring' | 'refactoring'; record: any }>,
  startRowIndex: number,
  targetDate: Date
): Array<{ start: number; end: number; ar: string; sum: number }> {
  const mergeGroups: Array<{ start: number; end: number; ar: string; sum: number }> = []
  rowsToAppend.forEach((rowData, idx) => {
    if (rowData.type === 'loan') return
    const repay = rowData.record as RepayRowRecord
    if (repay.repayDate && isSameDay(repay.repayDate, targetDate)) {
      const arValue = normalizeString(repay.colAH)
      if (arValue) {
        mergeGroups.push({
          start: startRowIndex + idx,
          end: startRowIndex + idx,
          ar: arValue,
          sum: toNumber(repay.colAG)
        })
      }
    }
  })
  return mergeGroups
}

function rowHasData(row: Row): boolean {
  if (!row || !row.hasValues || !row.values) return false
  const values = Array.isArray(row.values) ? row.values : Object.values(row.values)
  return values.some((value, idx) => {
    if (idx === 0) return false
    if (value === null || typeof value === 'undefined') return false
    if (typeof value === 'string' && value.trim() === '') return false
    return true
  })
}

function setValue(values: ExcelJS.RowValues, column: string, value: any) {
  const idx = columnLetterToNumber(column)
  ;(values as any)[idx] = value === undefined ? null : value
}

function columnLetterToNumber(letter: string): number {
  let num = 0
  const upper = letter.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    num = num * 26 + (upper.charCodeAt(i) - 64)
  }
  return num
}

// ========== 输入规则定义 ==========

const inputRules: FormCreateRule[] = [
  {
    type: 'Input',
    field: 'date',
    title: '日期',
    value: '',
    props: { placeholder: 'YYYYMMDD，可输入 2025-01-31 形式' },
    validate: [
      { required: true, message: '请输入日期', trigger: 'blur' },
      { pattern: '^\\d{4}[-/ ]?\\d{2}[-/ ]?\\d{2}$', message: '格式示例：20250131 或 2025-01-31' }
    ]
  }
]

// ========== 模板定义 ==========

export const ledgerDailyTemplate: TemplateDefinition<LedgerUserInput> = {
  meta: {
    id: 'ledgerDaily',
    name: '台账-融资及还款明细',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '基于放款/保理还款/再保理还款明细，按日追加台账「融资及还款明细」sheet',
    sourceLabel: '放款明细',
    extraSources: [
      {
        id: FACTORING_REPAY_SOURCE_ID,
        label: '保理融资还款明细',
        required: true,
        supportedExts: ['xlsx']
      },
      {
        id: REFACTORING_REPAY_SOURCE_ID,
        label: '再保理融资还款明细',
        required: true,
        supportedExts: ['xlsx']
      },
      {
        id: LEDGER_SOURCE_ID,
        label: '台账基线文件',
        required: true,
        supportedExts: ['xlsx'],
        description: '将基于此文件的「融资及还款明细」sheet 追加数据并输出到指定目录'
      }
    ]
  },
  engine: 'exceljs',
  inputRule: {
    rules: inputRules,
    options: { labelWidth: '80px', submitBtn: false, resetBtn: false },
    description: '输入目标日期（YYYYMMDD），系统将按顺序追加放款/保理还款/再保理还款记录。'
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  excelRenderer: renderWithStreamWriter,
  resolveReportName: ({ defaultName, parsedData, userInput }) => {
    const { ymd } = normalizeInputDate((userInput as LedgerUserInput | undefined)?.date ?? '')
    const ledgerFilename =
      (parsedData as LedgerParsedData | undefined)?.ledgerFilename || defaultName || 'ledger.xlsx'
    const parsed = path.parse(ledgerFilename)
    const base = sanitizeFilename(parsed.name || 'ledger')
    const ext = parsed.ext || '.xlsx'
    return `${base}-${ymd}${ext}`
  }
}
