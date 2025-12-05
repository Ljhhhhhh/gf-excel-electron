import ExcelJS from 'exceljs'
import type { Workbook, Cell, Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule } from './types'

// ========== 类型定义 ==========

interface Month4ParseOptions extends ParseOptions {
  sheet?: string | number
  dataStartRow?: number
  maxRows?: number
}

interface Month4ParsedRow {
  actualLoanDate: any
  industry: any
  businessLevel: any
  loanAmount: any
}

interface Month4ParsedData {
  rows: Month4ParsedRow[]
}

interface Month4UserInput {
  year: number
  month: number
}

const STREAM_WORKBOOK_OPTIONS = {
  sharedStrings: 'cache' as const,
  hyperlinks: 'ignore' as const,
  styles: 'ignore' as const,
  worksheets: 'emit' as const
}
const ROW_YIELD_INTERVAL = 2000

// ========== 解析器实现 ==========

export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: Month4ParseOptions
): Month4ParsedData {
  const options = {
    sheet: parseOptions?.sheet ?? 0,
    dataStartRow: parseOptions?.dataStartRow ?? 2,
    maxRows: parseOptions?.maxRows ?? 100000
  }

  const worksheet =
    typeof options.sheet === 'number'
      ? workbook.worksheets[options.sheet]
      : workbook.getWorksheet(options.sheet)

  if (!worksheet) {
    throw new Error(`无法找到工作表: ${options.sheet}`)
  }

  const rows: Month4ParsedData['rows'] = []
  const endRow = Math.min(worksheet.rowCount, options.dataStartRow + options.maxRows - 1)

  for (let rowIndex = options.dataStartRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    const parsed = extractDataRow(row)
    if (parsed) {
      rows.push(parsed)
    }
  }

  console.log(`[month4excel] 解析完成，共读取 ${rows.length} 行数据`)
  return { rows }
}

export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: Month4ParseOptions
): Promise<Month4ParsedData> {
  const options = {
    sheet: parseOptions?.sheet ?? 0,
    dataStartRow: parseOptions?.dataStartRow ?? 2,
    maxRows: parseOptions?.maxRows ?? 100000
  }

  console.log('[month4excel] 开始流式解析', {
    filePath,
    sheet: options.sheet,
    dataStartRow: options.dataStartRow,
    maxRows: options.maxRows
  })

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_WORKBOOK_OPTIONS)
  const rows: Month4ParsedData['rows'] = []
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

      const parsed = extractDataRow(row as Row)
      if (parsed) {
        rows.push(parsed)
      }

      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
    }

    processedSheet = true
    console.log(
      `[month4excel] 流式解析完成，共读取 ${rows.length} 行 (sheet=${sheetName ?? `#${sheetIndex}`})`
    )
    break
  }

  if (!processedSheet) {
    throw new Error(`[month4excel] 无法找到工作表: ${options.sheet}`)
  }

  return { rows }
}

function extractDataRow(row: Row): Month4ParsedRow | null {
  if (!row.hasValues) {
    return null
  }

  return {
    actualLoanDate: row.getCell(16).value, // P 列
    industry: row.getCell(27).value, // AA 列
    businessLevel: row.getCell(31).value, // AE 列
    loanAmount: row.getCell(49).value // AW 列
  }
}

// ========== 工具函数 ==========

function parseDate(value: any): Date | null {
  if (!value) return null
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? value : null
  if (typeof value === 'string') {
    const parsed = new Date(value)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  if (typeof value === 'number') {
    // Excel 序列号纪元：1899-12-30 (UTC)，避免本地时区偏移
    const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30)
    return new Date(EXCEL_EPOCH_MS + value * 86400000)
  }
  return null
}

function calculateAmount(
  rows: Month4ParsedData['rows'],
  year: number,
  month: number,
  businessLevel: string,
  industry: string
): number {
  const sum = rows
    .filter((row) => {
      const date = parseDate(row.actualLoanDate)
      if (!date) return false

      // 使用 UTC 方法保持与 parseDate 的一致性
      const matchDate = date.getUTCFullYear() === year && date.getUTCMonth() + 1 === month
      if (!matchDate) return false

      const level = String(row.businessLevel ?? '').trim()
      if (level !== businessLevel) return false

      const industryValue = String(row.industry ?? '').trim()
      return industryValue === industry
    })
    .reduce((acc, row) => acc + (Number(row.loanAmount) || 0), 0)

  return sum / 10000
}

function setCellStyle(
  cell: Cell,
  options: {
    bgColor: string
    fontColor: string
    bold?: boolean
    horizontal?: 'left' | 'center' | 'right'
  }
): void {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: options.bgColor }
  }
  cell.font = {
    color: { argb: options.fontColor },
    bold: options.bold ?? false,
    size: 11
  }
  cell.alignment = {
    horizontal: options.horizontal ?? 'center',
    vertical: 'middle'
  }
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  }
}

const HEADER_BG = 'FFED7D31'
const HEADER_FONT = 'FFFFFFFF'
const DATA_BG = 'FFFCECE8'
const DATA_FONT = 'FF000000'

function applyHeaderStyle(cell: Cell, horizontal: 'left' | 'center' | 'right' = 'center'): void {
  setCellStyle(cell, {
    bgColor: HEADER_BG,
    fontColor: HEADER_FONT,
    bold: true,
    horizontal
  })
}

function applyDataStyle(cell: Cell): void {
  setCellStyle(cell, {
    bgColor: DATA_BG,
    fontColor: DATA_FONT,
    bold: false,
    horizontal: 'center'
  })
}

// ========== ExcelJS 渲染器 ==========

async function renderWithExcelJS(
  parsedData: unknown,
  userInput: Month4UserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput) {
    throw new Error('[month4excel] 缺少用户输入参数')
  }

  const { year, month } = userInput
  if (month < 1 || month > 12) {
    throw new Error('[month4excel] 月份必须在 1-12 之间')
  }
  if (year < 1990 || year > 2100) {
    throw new Error('[month4excel] 年份必须在 1990-2100 之间')
  }

  const data = parsedData as Month4ParsedData

  const industries = [
    { label: '基建工程（万元）', source: '基建工程' },
    { label: '医药医疗（万元）', source: '医药医疗' },
    { label: '再保理（万元）', source: '大宗商品' }
  ]

  const businessLevels = [
    { label: '三级业务', value: '三级业务' },
    { label: '二级业务', value: '二级业务' },
    { label: '一级业务', value: '一级业务' }
  ]

  const matrix = businessLevels.map((level) => ({
    level: level.label,
    values: industries.map((industry) =>
      calculateAmount(data.rows, year, month, level.value, industry.source)
    )
  }))

  const rowTotals = matrix.map((row) => row.values.reduce((sum, value) => sum + value, 0))
  const grandTotal = rowTotals.reduce((sum, value) => sum + value, 0)
  const lowRiskTotal = rowTotals[0] + rowTotals[1]
  const lowRiskRatio = grandTotal === 0 ? 0 : lowRiskTotal / grandTotal

  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Sheet1')
  worksheet.columns = [
    { width: 18 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
    { width: 16 }
  ]

  worksheet.mergeCells('A1:F1')
  const cellA1 = worksheet.getCell('A1')
  const totalText = grandTotal.toFixed(2)
  const ratioText = (lowRiskRatio * 100).toFixed(2)
  cellA1.value = `${year}年${month}月，合计放款${totalText}万元，保理业务中低风险保理业务（明保+暗保核心企业应收）占比${ratioText}%。`
  applyHeaderStyle(cellA1)
  worksheet.getRow(1).height = 32

  const headers = [
    '业务等级',
    ...industries.map((industry) => industry.label),
    '总计（万元）',
    '比例'
  ]
  const row2 = worksheet.getRow(2)
  headers.forEach((header, idx) => {
    const cell = row2.getCell(idx + 1)
    cell.value = header
    applyHeaderStyle(cell)
  })
  row2.height = 25

  matrix.forEach((rowData, rowIdx) => {
    const excelRow = worksheet.getRow(3 + rowIdx)
    excelRow.height = 24

    const levelCell = excelRow.getCell(1)
    levelCell.value = rowData.level
    applyHeaderStyle(levelCell)

    rowData.values.forEach((value, colIdx) => {
      const cell = excelRow.getCell(colIdx + 2)
      cell.value = value
      cell.numFmt = '#,##0.00'
      applyDataStyle(cell)
    })

    const totalCell = excelRow.getCell(5)
    totalCell.value = { formula: `SUM(B${excelRow.number}:D${excelRow.number})` }
    totalCell.numFmt = '#,##0.00'
    applyDataStyle(totalCell)

    const ratioCell = excelRow.getCell(6)
    ratioCell.value = { formula: `IF($E$6=0,0,E${excelRow.number}/$E$6)` }
    ratioCell.numFmt = '0.00%'
    applyDataStyle(ratioCell)
  })

  const row6 = worksheet.getRow(6)
  row6.height = 25

  const cellA6 = row6.getCell(1)
  cellA6.value = '总计（万元）'
  applyHeaderStyle(cellA6, 'left')

  const summaryColumns: Array<{ letter: string; columnIndex: number }> = [
    { letter: 'B', columnIndex: 2 },
    { letter: 'C', columnIndex: 3 },
    { letter: 'D', columnIndex: 4 },
    { letter: 'E', columnIndex: 5 }
  ]

  summaryColumns.forEach(({ letter, columnIndex }) => {
    const cell = row6.getCell(columnIndex)
    cell.value = { formula: `SUM(${letter}3:${letter}5)` }
    cell.numFmt = '#,##0.00'
    applyHeaderStyle(cell)
  })

  const cellF6 = row6.getCell(6)
  cellF6.value = null
  applyHeaderStyle(cellF6)

  await workbook.xlsx.writeFile(outputPath)
  console.log(`[month4excel] ExcelJS 渲染完成: ${outputPath}`)
}

// ========== 表单规则 ==========

const inputRules: FormCreateRule[] = [
  {
    type: 'InputNumber',
    field: 'year',
    title: '年份',
    value: new Date().getFullYear(),
    props: {
      placeholder: '请输入年份（2020-2099）',
      min: 2020,
      max: 2099,
      step: 1
    },
    validate: [
      { required: true, message: '请输入年份', trigger: 'blur' },
      {
        type: 'number',
        min: 2020,
        max: 2099,
        message: '年份需在 2020-2099 之间',
        trigger: 'blur'
      }
    ]
  },
  {
    type: 'InputNumber',
    field: 'month',
    title: '月份',
    value: new Date().getMonth() + 1,
    props: {
      placeholder: '请输入月份（1-12）',
      min: 1,
      max: 12,
      step: 1
    },
    validate: [
      { required: true, message: '请输入月份', trigger: 'blur' },
      {
        type: 'number',
        min: 1,
        max: 12,
        message: '月份需在 1-12 之间',
        trigger: 'blur'
      }
    ]
  }
]

// ========== 模板定义 ==========

export const month4excelTemplate: TemplateDefinition<Month4UserInput> = {
  meta: {
    id: 'month4excel',
    name: '月报4',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description:
      '根据指定年月统计不同业务等级（一级-三级）在基建工程、医药医疗、再保理中的放款及占比',
    sourceLabel: '放款明细表'
  },
  engine: 'exceljs',
  inputRule: {
    rules: inputRules,
    options: {
      labelWidth: '120px',
      labelPosition: 'right',
      submitBtn: false,
      resetBtn: false
    },
    example: {
      year: 2024,
      month: 5
    },
    description: `
### 参数说明
- **年份**：筛选放款数据的年份
- **月份**：筛选放款数据的月份（1-12）

### 数据来源
- **P列**：实际放款日期（用于年月过滤）
- **AA列**：所属行业（基建工程 / 医药医疗 / 大宗商品）
- **AE列**：业务等级（一/二/三级业务）
- **AW列**：放款金额（元，最终换算为万元）

### 表格逻辑
- B3~D5：按业务等级与行业匹配的 AW 列金额求和（万元）
- E3~E5：对应行 B~D 的合计
- F3~F5：对应行总计占整体总计（E3~E5 / E6），保留两位小数
- B6~E6：第 3-5 行同列的合计，F6 留空
- A1：展示用户输入的年月、总计（E6）及低风险（三级+二级）占比
    `.trim()
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  excelRenderer: renderWithExcelJS
}
