import ExcelJS from 'exceljs'
import type { Workbook, CellValue, Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions } from './types'

interface Top10ParseOptions extends ParseOptions {
  financingSheet?: string | number
  customerSheet?: string | number
  financingDataStartRow?: number
  customerDataStartRow?: number
}

interface CustomerInfo {
  name: string
  industry: string
  region: string
}

interface Top10ParsedData {
  entries: AggregatedCustomer[]
  totalBalance: number
  customers: Record<string, CustomerInfo>
}

interface AggregatedCustomer {
  name: string
  balance: number
  principal: number
  tenors: number[]
  recourseStates: string[]
}

const DEFAULT_FINANCING_SHEET = '融资及还款明细'
const DEFAULT_CUSTOMER_SHEET = '客户表'
const DATA_START_ROW = 2
const MAX_CUSTOMERS = 10
const ROW_YIELD_INTERVAL = 2000
const PROGRESS_INTERVAL = 10000

const FINANCING_COLS = {
  customerName: 8,
  balance: 13,
  principal: 21,
  loanDate: 23,
  dueDate: 24,
  recourse: 40
}

const CUSTOMER_COLS = {
  region: 4,
  name: 5,
  industry: 8
}

// ========== 解析器实现 ==========

export async function parseWorkbook(
  workbook: Workbook,
  parseOptions?: Top10ParseOptions
): Promise<Top10ParsedData> {
  const options = {
    financingSheet: parseOptions?.financingSheet ?? DEFAULT_FINANCING_SHEET,
    customerSheet: parseOptions?.customerSheet ?? DEFAULT_CUSTOMER_SHEET,
    financingDataStartRow: parseOptions?.financingDataStartRow ?? DATA_START_ROW,
    customerDataStartRow: parseOptions?.customerDataStartRow ?? DATA_START_ROW
  }

  console.log('[top10customers] 开始解析来源工作簿', {
    financingSheet: options.financingSheet,
    customerSheet: options.customerSheet
  })

  const financingSheet =
    typeof options.financingSheet === 'number'
      ? workbook.worksheets[options.financingSheet]
      : workbook.getWorksheet(options.financingSheet)
  if (!financingSheet) {
    throw new Error(`[top10customers] 未找到工作表: ${options.financingSheet}`)
  }
  console.log(
    `[top10customers] 已加载融资表 ${financingSheet.name} (rows=${financingSheet.rowCount})`
  )

  const customerSheet =
    typeof options.customerSheet === 'number'
      ? workbook.worksheets[options.customerSheet]
      : workbook.getWorksheet(options.customerSheet)
  if (!customerSheet) {
    throw new Error(`[top10customers] 未找到工作表: ${options.customerSheet}`)
  }
  console.log(
    `[top10customers] 已加载客户表 ${customerSheet.name} (rows=${customerSheet.rowCount})`
  )

  const summary = new Map<string, AggregatedCustomer>()
  const financingEndRow = financingSheet.rowCount
  for (let rowIndex = options.financingDataStartRow; rowIndex <= financingEndRow; rowIndex++) {
    const row = financingSheet.getRow(rowIndex)
    if (!row.hasValues) {
      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
      continue
    }

    const customerName = normalizeString(row.getCell(FINANCING_COLS.customerName).value)
    if (!customerName) {
      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
      continue
    }

    let current = summary.get(customerName)
    if (!current) {
      current = { name: customerName, balance: 0, principal: 0, tenors: [], recourseStates: [] }
      summary.set(customerName, current)
    }

    const balance = toNumber(row.getCell(FINANCING_COLS.balance).value)
    if (isValidNumber(balance)) {
      current.balance += balance
    }

    const principal = toNumber(row.getCell(FINANCING_COLS.principal).value)
    if (isValidNumber(principal)) {
      current.principal += principal
    }

    const tenor = calculateTenorMonths(
      parseExcelDate(row.getCell(FINANCING_COLS.loanDate).value),
      parseExcelDate(row.getCell(FINANCING_COLS.dueDate).value)
    )
    if (typeof tenor === 'number') {
      current.tenors.push(tenor)
    }

    if (isValidNumber(balance) && balance !== 0) {
      const status = normalizeRecourseState(
        normalizeString(row.getCell(FINANCING_COLS.recourse).value)
      )
      if (status) {
        current.recourseStates.push(status)
      }
    }

    if (rowIndex % PROGRESS_INTERVAL === 0) {
      console.log(
        `[top10customers] 融资数据处理进度: ${rowIndex}/${financingEndRow} (uniqueCustomers=${summary.size})`
      )
    }

    if (rowIndex % ROW_YIELD_INTERVAL === 0) {
      await setImmediatePromise()
    }
  }

  const entries = Array.from(summary.values()).sort((a, b) => b.balance - a.balance)
  const totalBalance = entries.reduce((sum, entry) => sum + entry.balance, 0)
  console.log('[top10customers] 融资数据聚合完成', {
    customerCount: entries.length,
    totalBalance
  })

  const customers: Record<string, CustomerInfo> = {}
  const customerEndRow = customerSheet.rowCount
  for (let rowIndex = options.customerDataStartRow; rowIndex <= customerEndRow; rowIndex++) {
    const row = customerSheet.getRow(rowIndex)
    if (!row.hasValues) {
      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
      continue
    }

    const name = normalizeString(row.getCell(CUSTOMER_COLS.name).value)
    if (!name || customers[name]) {
      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
      continue
    }

    customers[name] = {
      name,
      industry: normalizeString(row.getCell(CUSTOMER_COLS.industry).value),
      region: normalizeString(row.getCell(CUSTOMER_COLS.region).value)
    }

    if (rowIndex % ROW_YIELD_INTERVAL === 0) {
      await setImmediatePromise()
    }
  }

  console.log(`[top10customers] 客户信息加载完成，共 ${Object.keys(customers).length} 条记录`)

  return { entries, totalBalance, customers }
}

/**
 * 流式解析器(优化版,用于大文件)
 * 优化点:
 * 1. 使用流式API,避免全量加载到内存
 * 2. 只处理指定的sheet,跳过其他sheet
 * 3. 遇到连续空行时提前终止
 * 4. 更频繁的yield控制,避免阻塞事件循环
 */
export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: ParseOptions
): Promise<Top10ParsedData> {
  const opts = parseOptions as Top10ParseOptions | undefined
  const options = {
    financingSheet: opts?.financingSheet ?? DEFAULT_FINANCING_SHEET,
    customerSheet: opts?.customerSheet ?? DEFAULT_CUSTOMER_SHEET,
    financingDataStartRow: opts?.financingDataStartRow ?? DATA_START_ROW,
    customerDataStartRow: opts?.customerDataStartRow ?? DATA_START_ROW
  }

  console.log('[top10customers] 开始流式解析', {
    filePath,
    financingSheet: options.financingSheet,
    customerSheet: options.customerSheet
  })

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'ignore',
    worksheets: 'emit'
  })

  const summary = new Map<string, AggregatedCustomer>()
  const customers: Record<string, CustomerInfo> = {}
  let currentSheetIndex = 0
  let processedFinancing = false
  let processedCustomer = false

  const targetFinancingName =
    typeof options.financingSheet === 'string' ? options.financingSheet : null
  const targetFinancingIndex =
    typeof options.financingSheet === 'number' ? options.financingSheet : null
  const targetCustomerName =
    typeof options.customerSheet === 'string' ? options.customerSheet : null
  const targetCustomerIndex =
    typeof options.customerSheet === 'number' ? options.customerSheet : null

  for await (const worksheetReader of workbookReader) {
    const sheetName = (worksheetReader as any).name
    const isFinancingSheet =
      (targetFinancingName && sheetName === targetFinancingName) ||
      (targetFinancingIndex !== null && currentSheetIndex === targetFinancingIndex)
    const isCustomerSheet =
      (targetCustomerName && sheetName === targetCustomerName) ||
      (targetCustomerIndex !== null && currentSheetIndex === targetCustomerIndex)

    if (isFinancingSheet && !processedFinancing) {
      console.log(`[top10customers] 流式处理融资表: ${sheetName}`)
      let rowIndex = 0

      for await (const row of worksheetReader) {
        rowIndex++

        if (rowIndex < options.financingDataStartRow) {
          continue
        }

        const rowData = row as Row
        if (!rowData.hasValues) {
          if (rowIndex % ROW_YIELD_INTERVAL === 0) {
            await setImmediatePromise()
          }
          continue
        }

        const customerName = normalizeString(rowData.getCell(FINANCING_COLS.customerName).value)
        if (!customerName) {
          if (rowIndex % ROW_YIELD_INTERVAL === 0) {
            await setImmediatePromise()
          }
          continue
        }

        let current = summary.get(customerName)
        if (!current) {
          current = { name: customerName, balance: 0, principal: 0, tenors: [], recourseStates: [] }
          summary.set(customerName, current)
        }

        const balance = toNumber(rowData.getCell(FINANCING_COLS.balance).value)
        if (isValidNumber(balance)) {
          current.balance += balance
        }

        const principal = toNumber(rowData.getCell(FINANCING_COLS.principal).value)
        if (isValidNumber(principal)) {
          current.principal += principal
        }

        const tenor = calculateTenorMonths(
          parseExcelDate(rowData.getCell(FINANCING_COLS.loanDate).value),
          parseExcelDate(rowData.getCell(FINANCING_COLS.dueDate).value)
        )
        if (typeof tenor === 'number') {
          current.tenors.push(tenor)
        }

        if (isValidNumber(balance) && balance !== 0) {
          const status = normalizeRecourseState(
            normalizeString(rowData.getCell(FINANCING_COLS.recourse).value)
          )
          if (status) {
            current.recourseStates.push(status)
          }
        }

        if (rowIndex % PROGRESS_INTERVAL === 0) {
          console.log(
            `[top10customers] 融资数据处理进度: ${rowIndex} (uniqueCustomers=${summary.size})`
          )
        }

        if (rowIndex % ROW_YIELD_INTERVAL === 0) {
          await setImmediatePromise()
        }
      }

      console.log(
        `[top10customers] 融资表处理完成 (totalRows=${rowIndex}, customers=${summary.size})`
      )
      processedFinancing = true
    } else if (isCustomerSheet && !processedCustomer) {
      console.log(`[top10customers] 流式处理客户表: ${sheetName}`)
      let rowIndex = 0

      for await (const row of worksheetReader) {
        rowIndex++

        if (rowIndex < options.customerDataStartRow) {
          continue
        }

        const rowData = row as Row
        if (!rowData.hasValues) {
          if (rowIndex % ROW_YIELD_INTERVAL === 0) {
            await setImmediatePromise()
          }
          continue
        }

        const name = normalizeString(rowData.getCell(CUSTOMER_COLS.name).value)
        if (!name || customers[name]) {
          if (rowIndex % ROW_YIELD_INTERVAL === 0) {
            await setImmediatePromise()
          }
          continue
        }

        customers[name] = {
          name,
          industry: normalizeString(rowData.getCell(CUSTOMER_COLS.industry).value),
          region: normalizeString(rowData.getCell(CUSTOMER_COLS.region).value)
        }

        if (rowIndex % ROW_YIELD_INTERVAL === 0) {
          await setImmediatePromise()
        }
      }

      console.log(
        `[top10customers] 客户表处理完成 (totalRows=${rowIndex}, customers=${Object.keys(customers).length})`
      )
      processedCustomer = true
    } else {
      // 跳过不需要的sheet - worksheetReader是async iterator,不处理即自动跳过
      console.log(`[top10customers] 跳过sheet: ${sheetName}`)
      // 不需要显式destroy,for await会自动处理
    }

    currentSheetIndex++

    // 如果两个表都处理完了,提前退出
    if (processedFinancing && processedCustomer) {
      console.log('[top10customers] 所有目标sheet已处理完成,提前退出')
      break
    }
  }

  if (!processedFinancing) {
    throw new Error(`[top10customers] 未找到融资表: ${options.financingSheet}`)
  }
  if (!processedCustomer) {
    throw new Error(`[top10customers] 未找到客户表: ${options.customerSheet}`)
  }

  const entries = Array.from(summary.values()).sort((a, b) => b.balance - a.balance)
  const totalBalance = entries.reduce((sum, entry) => sum + entry.balance, 0)

  console.log('[top10customers] 流式解析完成', {
    customerCount: entries.length,
    totalBalance,
    customerInfoCount: Object.keys(customers).length
  })

  return { entries, totalBalance, customers }
}

// ========== 渲染器实现 ==========

async function renderWithExcelJS(
  parsedData: unknown,
  _userInput: unknown | undefined,
  outputPath: string
): Promise<void> {
  console.log('[top10customers] ExcelJS 渲染开始')
  const data = parsedData as Top10ParsedData
  const { entries, totalBalance, customers } = data
  console.log(
    '[top10customers] 排名前十客户',
    entries.slice(0, MAX_CUSTOMERS).map((entry, idx) => ({
      rank: idx + 1,
      name: entry?.name,
      balance: entry?.balance
    }))
  )

  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Sheet1')
  worksheet.columns = [
    { width: 24 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 12 },
    { width: 14 },
    { width: 16 },
    { width: 18 }
  ]

  worksheet.mergeCells('A1:H1')
  const titleCell = worksheet.getCell('A1')
  titleCell.value = '根据余额统计前十客户信息'
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  titleCell.font = { bold: true, size: 14 }
  worksheet.getRow(1).height = 28

  const headers = ['客户名称', '行业', '地区', '发放金额', '占比', '期限（月）', '余额', '是否有追']
  const headerRow = worksheet.getRow(2)
  headerRow.height = 22
  headers.forEach((header, idx) => {
    const cell = headerRow.getCell(idx + 1)
    cell.value = header
    cell.font = { bold: true }
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    }
  })

  for (let i = 0; i < MAX_CUSTOMERS; i++) {
    const row = worksheet.getRow(3 + i)
    row.height = 20
    const entry = entries[i]
    if (!entry) {
      for (let col = 1; col <= headers.length; col++) {
        row.getCell(col).value = null
      }
      continue
    }

    const customerInfo = customers[entry.name]
    row.getCell(1).value = entry.name
    row.getCell(2).value = customerInfo?.industry ?? ''
    row.getCell(3).value = customerInfo?.region ?? ''

    const principalCell = row.getCell(4)
    principalCell.value = entry.principal ?? 0
    principalCell.numFmt = '#,##0.00'

    const ratioCell = row.getCell(5)
    ratioCell.value = totalBalance > 0 ? entry.balance / totalBalance : 0
    ratioCell.numFmt = '0.00%'

    row.getCell(6).value = formatTenor(entry.tenors)

    const balanceCell = row.getCell(7)
    balanceCell.value = entry.balance ?? 0
    balanceCell.numFmt = '#,##0.00'

    row.getCell(8).value = formatRecourse(entry.recourseStates)
  }

  console.log('[top10customers] 写入 Excel 文件', outputPath)
  await workbook.xlsx.writeFile(outputPath)
  console.log(`[top10customers] ExcelJS 渲染完成: ${outputPath}`)
}

// ========== 工具函数 ==========

function unwrapCellValue(value: CellValue): any {
  if (value && typeof value === 'object') {
    if ('result' in value && typeof value.result !== 'undefined') {
      return unwrapCellValue(value.result as CellValue)
    }
    if ('text' in value && typeof value.text === 'string') {
      return value.text
    }
    if ('richText' in (value as any) && Array.isArray((value as any).richText)) {
      return (value as any).richText.map((item: any) => item.text ?? '').join('')
    }
  }
  return value
}

function isPlaceholder(value: string): boolean {
  const trimmed = value.trim()
  return trimmed === '' || trimmed === '/' || trimmed === '-'
}

function normalizeString(value: CellValue): string {
  if (value === null || typeof value === 'undefined') return ''
  const unwrapped = unwrapCellValue(value)
  if (typeof unwrapped === 'string') {
    return isPlaceholder(unwrapped) ? '' : unwrapped.trim()
  }
  if (typeof unwrapped === 'number') {
    return Number.isNaN(unwrapped) ? '' : String(unwrapped)
  }
  return ''
}

function toNumber(value: CellValue): number | null {
  if (value === null || typeof value === 'undefined') return null
  const unwrapped = unwrapCellValue(value)
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

function parseExcelDate(value: CellValue): Date | null {
  if (value === null || typeof value === 'undefined') return null
  const unwrapped = unwrapCellValue(value)
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

function isValidNumber(value: number | null): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function calculateTenorMonths(loanDate: Date | null, dueDate: Date | null): number | null {
  if (!loanDate || !dueDate) return null
  const diff = dueDate.getTime() - loanDate.getTime()
  if (!Number.isFinite(diff)) return null
  const days = diff / 86400000
  const months = Math.ceil(days / 30)
  if (!Number.isFinite(months)) return null
  return Math.max(0, months)
}

function formatTenor(values: number[]): string {
  if (!values.length) return ''
  const unique = Array.from(new Set(values)).sort((a, b) => a - b)
  if (unique.length === 1) {
    return `${unique[0]}`
  }
  return `${unique[0]}-${unique[unique.length - 1]}`
}

function normalizeRecourseState(value: string): string {
  if (!value) return ''
  const trimmed = value.trim()
  if (isPlaceholder(trimmed)) {
    return ''
  }
  return trimmed
}

function formatRecourse(states: string[]): string {
  const valid = states.filter((state) => !!state && !isPlaceholder(state))
  if (!valid.length) return ''

  const allNoRecourse = valid.every((state) => state === '无追索权')
  if (allNoRecourse) {
    return '无追索权'
  }

  const noneNoRecourse = valid.every((state) => state !== '无追索权')
  if (noneNoRecourse) {
    return '有追索权'
  }

  return '有追索权/无追索权'
}

// ========== 模板定义 ==========

export const top10CustomersTemplate: TemplateDefinition = {
  meta: {
    id: 'top10customers',
    name: '前十大客户',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '根据融资余额汇总并展示前十名客户的行业、地区、金额、占比与追索权情况'
  },
  engine: 'exceljs',
  parser: parseWorkbook,
  streamParser: streamParseWorkbook, // 优化: 使用流式解析器处理大文件
  excelRenderer: renderWithExcelJS
}
