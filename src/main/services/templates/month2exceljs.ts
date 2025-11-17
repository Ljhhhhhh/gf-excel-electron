/**
 * 月报2模板：month2exceljs
 * 行业月度新增放款统计表
 * 纯 ExcelJS 渲染，不依赖 Carbone 模板文件
 */

import ExcelJS from 'exceljs'
import type { Workbook, Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule } from './types'

// ========== 类型定义 ==========

interface Month2ParseOptions extends ParseOptions {
  /** 指定要解析的 sheet 名称或索引，默认解析首表 */
  sheet?: string | number
  /** 数据起始行索引（1-based），默认为 2（第1行是标题） */
  dataStartRow?: number
  /** 最大行数限制，默认 100000 */
  maxRows?: number
}

interface Month2ParsedData {
  /** 原始行数据 */
  rows: Array<{
    实际放款日期: any
    所属行业: any
    放款金额: any
  }>
}

const DEFAULT_DATA_START_ROW = 2
const DEFAULT_MAX_ROWS = 100000
const ROW_YIELD_INTERVAL = 2000
const PROGRESS_INTERVAL = 10000
const STREAM_WORKBOOK_OPTIONS = {
  sharedStrings: 'cache' as const,
  hyperlinks: 'ignore' as const,
  styles: 'ignore' as const,
  worksheets: 'emit' as const
}

const COLUMN_INDEX = {
  actualLoanDate: 16,
  industry: 27,
  amount: 49
}

interface Month2UserInput {
  /** 查询年份 */
  queryYear: number
  /** 截止月份（1-12） */
  endMonth: number
}

// ========== 解析器实现 ==========

/**
 * 解析 workbook，提取关键列数据
 */
export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: Month2ParseOptions
): Month2ParsedData {
  const options = {
    sheet: parseOptions?.sheet ?? 0,
    dataStartRow: parseOptions?.dataStartRow ?? DEFAULT_DATA_START_ROW,
    maxRows: parseOptions?.maxRows ?? DEFAULT_MAX_ROWS
  }

  // 获取工作表
  const worksheet =
    typeof options.sheet === 'number'
      ? workbook.worksheets[options.sheet]
      : workbook.getWorksheet(options.sheet)

  if (!worksheet) {
    throw new Error(`无法找到工作表: ${options.sheet}`)
  }

  const rows: Month2ParsedData['rows'] = []
  const endRow = Math.min(worksheet.rowCount, options.dataStartRow + options.maxRows - 1)

  // 从指定行开始读取数据
  for (let rowNumber = options.dataStartRow; rowNumber <= endRow; rowNumber++) {
    const row = worksheet.getRow(rowNumber)

    if (!row.hasValues) {
      continue
    }

    const parsed = extractRowData(row)
    if (parsed) {
      rows.push(parsed)
    }
  }

  console.log(`[month2exceljs] 解析完成，共读取 ${rows.length} 行数据`)
  return { rows }
}

/**
 * 流式解析 workbook，适合大文件
 */
export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: Month2ParseOptions
): Promise<Month2ParsedData> {
  const options = {
    sheet: parseOptions?.sheet ?? 0,
    dataStartRow: parseOptions?.dataStartRow ?? DEFAULT_DATA_START_ROW,
    maxRows: parseOptions?.maxRows ?? DEFAULT_MAX_ROWS
  }

  console.log('[month2exceljs] 开始流式解析', {
    filePath,
    sheet: options.sheet,
    dataStartRow: options.dataStartRow,
    maxRows: options.maxRows
  })

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_WORKBOOK_OPTIONS)
  const rows: Month2ParsedData['rows'] = []

  const targetSheetName = typeof options.sheet === 'string' ? options.sheet : null
  const targetSheetIndex = typeof options.sheet === 'number' ? options.sheet : null
  let currentSheetIndex = 0
  let processedTargetSheet = false

  for await (const worksheetReader of workbookReader) {
    const sheetName = (worksheetReader as any).name as string | undefined
    const isTargetSheet =
      (targetSheetName && sheetName === targetSheetName) ||
      (targetSheetIndex !== null && currentSheetIndex === targetSheetIndex)

    if (!isTargetSheet || processedTargetSheet) {
      currentSheetIndex++
      continue
    }

    console.log(`[month2exceljs] 流式处理工作表: ${sheetName ?? `#${currentSheetIndex + 1}`}`)
    let rowIndex = 0

    for await (const row of worksheetReader) {
      rowIndex++

      if (rowIndex < options.dataStartRow) {
        continue
      }

      const rowData = row as Row
      if (!rowData.hasValues) {
        if (rowIndex % ROW_YIELD_INTERVAL === 0) {
          await setImmediatePromise()
        }
        continue
      }

      if (rows.length >= options.maxRows) {
        if (rowIndex % ROW_YIELD_INTERVAL === 0) {
          await setImmediatePromise()
        }
        continue
      }

      const parsed = extractRowData(rowData)
      if (parsed) {
        rows.push(parsed)
      }

      if (rowIndex % PROGRESS_INTERVAL === 0) {
        console.log(`[month2exceljs] 流式解析进度: row=${rowIndex}, collected=${rows.length}`)
      }

      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
    }

    processedTargetSheet = true
    console.log(
      `[month2exceljs] 流式解析完成: 读取 ${rows.length} 行 (targetSheet=${sheetName ?? targetSheetIndex})`
    )
    currentSheetIndex++

    if (processedTargetSheet) {
      break
    }
  }

  if (!processedTargetSheet) {
    throw new Error(`[month2exceljs] 未找到工作表: ${options.sheet}`)
  }

  return { rows }
}

// ========== 工具函数 ==========

/**
 * 解析日期值（支持多种格式）
 */
function parseDate(value: any): Date | null {
  if (!value) return null

  // ExcelJS 已经解析为 Date 对象
  if (value instanceof Date) {
    return value
  }

  // 字符串格式
  if (typeof value === 'string') {
    const parsed = new Date(value)
    return isNaN(parsed.getTime()) ? null : parsed
  }

  // Excel 序列号格式
  if (typeof value === 'number') {
    // Excel 日期从 1900-01-01 开始
    const excelEpoch = new Date(1899, 11, 30)
    return new Date(excelEpoch.getTime() + value * 86400000)
  }

  return null
}

/**
 * 计算指定年月和行业的放款金额总和（单位：万元）
 */
function calculateMonthAmount(
  rows: Month2ParsedData['rows'],
  year: number,
  month: number,
  targetIndustry: string
): number {
  const sum = rows
    .filter((row) => {
      // 解析日期
      const date = parseDate(row.实际放款日期)
      if (!date) return false

      // 匹配年月
      const matchDate = date.getFullYear() === year && date.getMonth() + 1 === month

      // 匹配行业
      const industry = String(row.所属行业 || '').trim()
      const matchIndustry = industry === targetIndustry

      return matchDate && matchIndustry
    })
    .reduce((sum, row) => {
      const amount = Number(row.放款金额) || 0
      return sum + amount
    }, 0)

  // 转换为万元
  return sum / 10000
}

/**
 * 获取列字母（A, B, C, ..., Z, AA, AB, ...）
 */
function getColumnLetter(columnIndex: number): string {
  let letter = ''
  while (columnIndex > 0) {
    const remainder = (columnIndex - 1) % 26
    letter = String.fromCharCode(65 + remainder) + letter
    columnIndex = Math.floor((columnIndex - 1) / 26)
  }
  return letter
}

function extractRowData(row: Row): Month2ParsedData['rows'][number] | null {
  if (!row.hasValues) {
    return null
  }

  return {
    实际放款日期: row.getCell(COLUMN_INDEX.actualLoanDate).value,
    所属行业: row.getCell(COLUMN_INDEX.industry).value,
    放款金额: row.getCell(COLUMN_INDEX.amount).value
  }
}

/**
 * 设置单元格样式（背景色 + 文字颜色 + 边框）
 */
function setCellStyle(
  cell: ExcelJS.Cell,
  bgColor: string,
  fontColor: string,
  bold: boolean = false
): void {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: bgColor }
  }
  cell.font = {
    color: { argb: fontColor },
    bold,
    size: 11
  }
  cell.alignment = {
    horizontal: 'center',
    vertical: 'middle'
  }
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  }
}

// ========== ExcelJS 渲染器 ==========

/**
 * 使用 ExcelJS 从零创建行业月度放款统计表
 */
async function renderWithExcelJS(
  parsedData: unknown,
  userInput: Month2UserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput) {
    throw new Error('[month2exceljs] 缺少用户输入参数')
  }

  // 类型断言
  const data = parsedData as Month2ParsedData
  const { queryYear, endMonth } = userInput

  // 验证参数
  if (endMonth < 1 || endMonth > 12) {
    throw new Error(`[month2exceljs] 截止月份必须在 1-12 之间，当前值: ${endMonth}`)
  }

  console.log(`[month2exceljs] 开始生成报表: ${queryYear}年 1-${endMonth}月`)

  // 创建工作簿
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('行业月度放款')

  // 设置列宽
  const columnCount = endMonth + 1 // A列 + 月份列
  worksheet.columns = Array(columnCount)
    .fill(null)
    .map((_, idx) => ({
      width: idx === 0 ? 15 : 18
    }))

  // 行业配置（显示名称 -> 数据源中的行业名称）
  const industries = [
    { display: '基建工程', source: '基建工程' },
    { display: '医药医疗', source: '医药医疗' },
    { display: '再保理', source: '大宗商品' }
  ]

  // ========== 第1行：标题行 ==========
  const row1 = worksheet.getRow(1)
  row1.height = 30

  // A1:A2 合并单元格，显示"行业"
  worksheet.mergeCells('A1:A2')
  const cellA1 = row1.getCell(1)
  cellA1.value = '行业'
  setCellStyle(cellA1, 'FFED7D30', 'FFFFFFFF', true)

  // B1 到最后一列：当月新增放款（万元）
  if (endMonth > 0) {
    worksheet.mergeCells(1, 2, 1, columnCount)
    const cellB1 = row1.getCell(2)
    cellB1.value = '当月新增放款（万元）'
    setCellStyle(cellB1, 'FFED7D30', 'FFFFFFFF', true)
  }

  // ========== 第2行：表头行（月份） ==========
  const row2 = worksheet.getRow(2)
  row2.height = 25

  // A2 已被 A1:A2 合并，这里只处理 B2 开始的月份列

  // B2 到最后：1月、2月、...
  for (let month = 1; month <= endMonth; month++) {
    const cell = row2.getCell(month + 1)
    cell.value = `${month}月`
    setCellStyle(cell, 'FFED7D30', 'FFFFFFFF', true)
  }

  // ========== 第3-5行：行业数据行 ==========
  industries.forEach((industry, idx) => {
    const rowIdx = 3 + idx
    const row = worksheet.getRow(rowIdx)
    row.height = 25

    // A列：行业名称
    const cellA = row.getCell(1)
    cellA.value = industry.display
    setCellStyle(cellA, 'FFFCECE8', 'FF000000', false)

    // 月份列：计算数据
    for (let month = 1; month <= endMonth; month++) {
      const cell = row.getCell(month + 1)
      const amount = calculateMonthAmount(data.rows, queryYear, month, industry.source)
      cell.value = amount
      cell.numFmt = '#,##0.00' // 千分位格式
      setCellStyle(cell, 'FFFCECE8', 'FF000000', false)
    }
  })

  // ========== 第6行：合计行 ==========
  const row6 = worksheet.getRow(6)
  row6.height = 25

  // A6: 合计
  const cellA6 = row6.getCell(1)
  cellA6.value = '合计'
  setCellStyle(cellA6, 'FFED7D30', 'FFFFFFFF', true)

  // B6 到最后：公式求和
  for (let month = 1; month <= endMonth; month++) {
    const cell = row6.getCell(month + 1)
    const colLetter = getColumnLetter(month + 1)
    cell.value = { formula: `SUM(${colLetter}3:${colLetter}5)` }
    cell.numFmt = '#,##0.00'
    setCellStyle(cell, 'FFED7D30', 'FFFFFFFF', true)
  }

  // 写入文件
  await workbook.xlsx.writeFile(outputPath)
  console.log(`[month2exceljs] ExcelJS 渲染完成: ${outputPath}`)
}

// ========== 表单规则 ==========

const inputRules: FormCreateRule[] = [
  {
    type: 'InputNumber',
    field: 'queryYear',
    title: '查询年份',
    value: new Date().getFullYear(),
    props: {
      placeholder: '请输入年份',
      min: 2020,
      max: 2099,
      step: 1
    },
    validate: [
      {
        required: true,
        message: '请输入查询年份',
        trigger: 'blur'
      },
      {
        type: 'number',
        min: 2020,
        max: 2099,
        message: '年份必须在 2020-2099 之间',
        trigger: 'blur'
      }
    ]
  },
  {
    type: 'InputNumber',
    field: 'endMonth',
    title: '截止月份',
    value: new Date().getMonth() + 1,
    props: {
      placeholder: '请输入截止月份（1-12）',
      min: 1,
      max: 12,
      step: 1
    },
    validate: [
      {
        required: true,
        message: '请输入截止月份',
        trigger: 'blur'
      },
      {
        type: 'number',
        min: 1,
        max: 12,
        message: '月份必须在 1-12 之间',
        trigger: 'blur'
      }
    ]
  }
]

// ========== 模板定义与导出 ==========

export const month2exceljsTemplate: TemplateDefinition<Month2UserInput> = {
  meta: {
    id: 'month2exceljs',
    name: '月报2模板',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '统计从1月到指定月份的三个行业（基建工程、医药医疗、再保理）新增放款数据',
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
      queryYear: 2025,
      endMonth: 8
    },
    description: `
### 参数说明
- **查询年份**: 筛选指定年份的放款数据
- **截止月份**: 从1月统计到该月份（1-12）

### 使用示例
选择 "2025年 8月" 将生成一个表格，包含8列（1月到8月）

### 统计行业
- 基建工程：数据源中"所属行业"为"基建工程"
- 医药医疗：数据源中"所属行业"为"医药医疗"
- 再保理：数据源中"所属行业"为"大宗商品"

### 数据源列说明
- P列：实际放款日期（用于月份筛选）
- AA列：所属行业（用于行业分组）
- AW列：放款金额（求和后除以10000转为万元）
    `.trim()
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  excelRenderer: renderWithExcelJS
}
