import ExcelJS from 'exceljs'
import type { Workbook, Cell, Worksheet, DataValidation, Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'
import type { TemplateDefinition, ParseOptions, FormCreateRule } from './types'

interface FactoringParseOptions extends ParseOptions {
  sheet?: string | number
  dataStartRow?: number
  maxRows?: number
}

interface FactoringParsedRow {
  customerName: any
  enterpriseType: any
  actualLoanDate: any
  dueDate: any
  principalAmount: any
  rateAdjusted: any
  rateOriginal: any
  assetId: any
}

interface FactoringParsedData {
  rows: FactoringParsedRow[]
}

interface FactoringUserInput {
  fillMonth: string
}

type WorksheetWithValidation = Worksheet & {
  dataValidations: {
    add(range: string, rules: DataValidation): void
  }
}

const ECONOMIC_INDUSTRY_CLASSIFICATIONS = [
  '农业',
  '林业',
  '畜牧业',
  '渔业',
  '农、林、牧、渔服务业',
  '煤炭开采和洗选业',
  '石油和天然气开采业',
  '黑色金属矿采选业',
  '有色金属矿采选业',
  '非金属矿采选业',
  '开采辅助活动',
  '其他采矿业',
  '农副食品加工业',
  '食品制造业',
  '酒、饮料和精制茶制造业',
  '烟草制品业',
  '纺织业',
  '纺织服装、服饰业',
  '皮革、毛皮、羽毛及其制品和制鞋业',
  '木材加工和木、竹、藤、棕、草制品业',
  '家具制造业',
  '造纸和纸制品业',
  '印刷和记录媒介复制业',
  '文教、工美、体育和娱乐用品制造业',
  '石油加工、炼焦和核燃料加工业',
  '化学原料和化学制品制造业',
  '医药制造业',
  '化学纤维制造业',
  '橡胶和塑料制品业',
  '非金属矿物制品业',
  '黑色金属冶炼和压延加工业',
  '有色金属冶炼和压延加工业',
  '金属制品业',
  '通用设备制造业',
  '专用设备制造业',
  '汽车制造业',
  '铁路、船舶、航空航天和其他运输设备制造业',
  '电气机械和器材制造业',
  '计算机、通信和其他电子设备制造业',
  '仪器仪表制造业',
  '其他制造业',
  '废弃资源综合利用业',
  '金属制品、机械和设备修理业',
  '电力、热力生产和供应业',
  '燃气生产和供应业',
  '水的生产和供应业',
  '房屋建筑业',
  '土木工程建筑业',
  '建筑安装业',
  '建筑装饰和其他建筑业',
  '批发业',
  '零售业',
  '铁路运输业',
  '道路运输业',
  '水上运输业',
  '航空运输业',
  '管道运输业',
  '装卸搬运和运输代理业',
  '仓储业',
  '邮政业',
  '住宿业',
  '餐饮业',
  '电信、广播电视和卫星传输服务',
  '互联网和相关服务',
  '软件和信息技术服务业',
  '货币金融服务',
  '资本市场服务',
  '保险业',
  '其他金融业',
  '房地产业',
  '租赁业',
  '商务服务业',
  '研究和试验发展',
  '专业技术服务业',
  '科技推广和应用服务业',
  '水利管理业',
  '生态保护和环境治理业',
  '公共设施管理业',
  '居民服务业',
  '机动车、电子产品和日用产品修理业',
  '其他服务业',
  '教育',
  '卫生',
  '社会工作',
  '新闻和出版业',
  '广播、电视、电影和影视录音制作业',
  '文化艺术业',
  '体育',
  '娱乐业',
  '中国共产党机关',
  '国家机构',
  '人民政协、民主党派',
  '社会保障',
  '群众团体、社会团体和其他成员组织',
  '基层群众自治组织',
  '国际组织'
]

const HEADER_FILL = 'FFDDEBF7'
const HEADER_FONT_COLOR = 'FF000000'
const HEADER_FONT_SIZE = 10
const STREAM_WORKBOOK_OPTIONS = {
  sharedStrings: 'cache' as const,
  hyperlinks: 'ignore' as const,
  styles: 'ignore' as const,
  worksheets: 'emit' as const
}
const ROW_YIELD_INTERVAL = 2000

export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: FactoringParseOptions
): FactoringParsedData {
  const options = {
    sheet: parseOptions?.sheet ?? 0,
    dataStartRow: parseOptions?.dataStartRow ?? 2,
    maxRows: parseOptions?.maxRows ?? 200000
  }

  const worksheet =
    typeof options.sheet === 'number'
      ? workbook.worksheets[options.sheet]
      : workbook.getWorksheet(options.sheet)

  if (!worksheet) {
    throw new Error(`[f103FactoringDetail] 无法找到工作表: ${options.sheet}`)
  }

  const rows: FactoringParsedData['rows'] = []
  const endRow = Math.min(worksheet.rowCount, options.dataStartRow + options.maxRows - 1)

  for (let rowIndex = options.dataStartRow; rowIndex <= endRow; rowIndex++) {
    const row = worksheet.getRow(rowIndex)
    const parsed = extractFactoringRow(row)
    if (parsed) {
      rows.push(parsed)
    }
  }

  console.log(`[f103FactoringDetail] 解析完成，共获取 ${rows.length} 行放款记录`)
  return { rows }
}

export async function streamParseWorkbook(
  filePath: string,
  parseOptions?: FactoringParseOptions
): Promise<FactoringParsedData> {
  const options = {
    sheet: parseOptions?.sheet ?? 0,
    dataStartRow: parseOptions?.dataStartRow ?? 2,
    maxRows: parseOptions?.maxRows ?? 200000
  }

  console.log('[f103FactoringDetail] 开始流式解析', {
    filePath,
    sheet: options.sheet,
    dataStartRow: options.dataStartRow,
    maxRows: options.maxRows
  })

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_WORKBOOK_OPTIONS)
  const rows: FactoringParsedRow[] = []
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

      const parsed = extractFactoringRow(row as Row)
      if (parsed) {
        rows.push(parsed)
      }

      if (rowIndex % ROW_YIELD_INTERVAL === 0) {
        await setImmediatePromise()
      }
    }

    processedSheet = true
    console.log(
      `[f103FactoringDetail] 流式解析完成，共获取 ${rows.length} 行 (sheet=${sheetName ?? `#${sheetIndex}`})`
    )
    break
  }

  if (!processedSheet) {
    throw new Error(`[f103FactoringDetail] 无法找到工作表: ${options.sheet}`)
  }

  return { rows }
}

function extractFactoringRow(row: Row): FactoringParsedRow | null {
  if (!row.hasValues) {
    return null
  }

  return {
    customerName: row.getCell(3).value,
    enterpriseType: row.getCell(6).value,
    actualLoanDate: row.getCell(16).value,
    dueDate: row.getCell(42).value,
    principalAmount: row.getCell(49).value,
    rateAdjusted: row.getCell(59).value,
    rateOriginal: row.getCell(58).value,
    assetId: row.getCell(11).value
  }
}

function parseExcelDate(value: any): Date | null {
  if (!value) return null
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? value : null
  if (typeof value === 'string') {
    const parsed = new Date(value)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  if (typeof value === 'number') {
    // Excel 序列号纪元：1899-12-30 (UTC)，避免本地时区偏移
    const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30)
    const ms = EXCEL_EPOCH_MS + value * 86400000
    const parsed = new Date(ms)
    return Number.isFinite(parsed.getTime()) ? parsed : null
  }
  return null
}

function parseNumber(value: any): number | null {
  if (value === null || typeof value === 'undefined') return null
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null
  }
  if (typeof value === 'string') {
    const normalized = value.replace(/[^0-9.+-]/g, '').trim()
    if (!normalized) return null
    const parsed = Number(normalized)
    return Number.isFinite(parsed) ? parsed : null
  }
  if (typeof value === 'object' && 'text' in value && typeof value.text === 'string') {
    return parseNumber(value.text)
  }
  return null
}

function deriveRate(row: FactoringParsedRow): number | null {
  const adjusted = parseNumber(row.rateAdjusted)
  const fallback = parseNumber(row.rateOriginal)
  const rate = adjusted ?? fallback
  if (rate === null) return null
  if (rate > 0 && rate <= 1) {
    return rate * 100
  }
  return rate
}

function applyHeaderStyle(cell: Cell): void {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: HEADER_FILL }
  }
  cell.font = {
    bold: true,
    size: HEADER_FONT_SIZE,
    color: { argb: HEADER_FONT_COLOR }
  }
  cell.alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  }
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  }
}

function applyDataStyle(
  cell: Cell,
  align: 'left' | 'center' | 'right' = 'center',
  wrapText = false
): void {
  cell.font = {
    size: 10,
    color: { argb: 'FF000000' }
  }
  cell.alignment = {
    horizontal: align,
    vertical: 'middle',
    wrapText
  }
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  }
}

function ensureValidMonth(month: string): { year: number; month: number; formatted: string } {
  const normalized = month?.trim()
  if (!/^[0-9]{6}$/.test(normalized)) {
    throw new Error('[f103FactoringDetail] 填报年月格式需为 YYYYMM，例如 202501')
  }
  const year = Number(normalized.slice(0, 4))
  const monthValue = Number(normalized.slice(4))
  if (monthValue < 1 || monthValue > 12) {
    throw new Error('[f103FactoringDetail] 月份必须在 01-12 范围内')
  }
  return { year, month: monthValue, formatted: normalized }
}

function filterRowsByMonth(
  rows: FactoringParsedRow[],
  year: number,
  month: number
): FactoringParsedRow[] {
  return rows.filter((row) => {
    const date = parseExcelDate(row.actualLoanDate)
    if (!date) return false
    // 使用 UTC 方法保持与 parseExcelDate 的一致性
    return date.getUTCFullYear() === year && date.getUTCMonth() + 1 === month
  })
}

async function renderWithExcelJS(
  parsedData: unknown,
  userInput: FactoringUserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput) {
    throw new Error('[f103FactoringDetail] 缺少用户输入参数')
  }

  const { year, month, formatted } = ensureValidMonth(userInput.fillMonth)
  const data = parsedData as FactoringParsedData
  const filteredRows = filterRowsByMonth(data.rows, year, month)

  // 按发放日期（E列）升序排序
  filteredRows.sort((a, b) => {
    const dateA = parseExcelDate(a.actualLoanDate)
    const dateB = parseExcelDate(b.actualLoanDate)
    if (!dateA && !dateB) return 0
    if (!dateA) return 1
    if (!dateB) return -1
    return dateA.getTime() - dateB.getTime()
  })

  console.log(`[f103FactoringDetail] 生成 ${formatted} 数据，共 ${filteredRows.length} 条记录`)

  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('F103融资保理业务明细表')
  worksheet.views = [{ state: 'frozen', ySplit: 6 }]
  worksheet.columns = [
    { width: 12 },
    { width: 26 },
    { width: 24 },
    { width: 14 },
    { width: 15 },
    { width: 15 },
    { width: 18 },
    { width: 14 },
    { width: 14 },
    { width: 18 },
    { width: 12 },
    { width: 12 },
    { width: 18 },
    { width: 16 },
    { width: 22 }
  ]

  worksheet.mergeCells('A1:O1')
  const titleCell = worksheet.getCell('A1')
  titleCell.value = '融资保理业务明细表'
  titleCell.font = { size: 13, bold: true }
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  worksheet.getRow(1).height = 30

  worksheet.getCell('A2').value = '填报单位'
  worksheet.getCell('A2').font = { bold: true, size: 10 }
  worksheet.getCell('A2').alignment = { horizontal: 'left', vertical: 'middle' }
  worksheet.mergeCells('B2:O2')
  const unitCell = worksheet.getCell('B2')
  unitCell.value = '宁波国富商业保理有限公司'
  unitCell.alignment = { horizontal: 'left', vertical: 'middle' }
  unitCell.font = { size: 10, bold: true }

  worksheet.getCell('A3').value = '填报年月'
  worksheet.getCell('A3').font = { bold: true, size: 10 }
  worksheet.getCell('A3').alignment = { horizontal: 'left', vertical: 'middle' }
  worksheet.mergeCells('B3:O3')
  const monthCell = worksheet.getCell('B3')
  monthCell.value = formatted
  monthCell.font = { size: 10, bold: true }
  monthCell.alignment = { horizontal: 'left', vertical: 'middle' }

  worksheet.getCell('A4').value = '单位：万元'
  worksheet.getCell('A4').font = { bold: true, size: 10 }
  worksheet.getCell('A4').alignment = { horizontal: 'left', vertical: 'middle' }
  worksheet.mergeCells('B4:O4')
  const formCell = worksheet.getCell('B4')
  formCell.value = '表号：F103'
  formCell.font = { bold: true, size: 10 }
  formCell.alignment = { horizontal: 'right', vertical: 'middle' }

  const mergedColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'M', 'N', 'O']
  mergedColumns.forEach((col) => worksheet.mergeCells(`${col}5:${col}6`))
  worksheet.mergeCells('K5:L5')

  const headerLabels: Record<string, string> = {
    A: '序号',
    B: '融资人名称',
    C: '经济行业分类¹',
    D: '企业类型²',
    E: '发放日期',
    F: '到期日',
    G: '发放保理融资本金',
    H: '利息收入',
    I: '服务费收入',
    J: '保理融资综合费率(%)',
    K: '业务类型',
    M: '五级分类情况³',
    N: '增信方式',
    O: '备注'
  }

  Object.entries(headerLabels).forEach(([column, label]) => {
    const cell = worksheet.getCell(`${column}5`)
    cell.value = label
    applyHeaderStyle(cell)
  })

  const k6 = worksheet.getCell('K6')
  k6.value = '有追索权'
  applyHeaderStyle(k6)

  const l6 = worksheet.getCell('L6')
  l6.value = '无追索权'
  applyHeaderStyle(l6)

  const dataStartRow = 7
  filteredRows.forEach((row, index) => {
    const excelRow = worksheet.getRow(dataStartRow + index)
    const seqCell = excelRow.getCell('A')
    seqCell.value = index + 1
    applyDataStyle(seqCell)

    const nameCell = excelRow.getCell('B')
    nameCell.value = row.customerName ?? ''
    applyDataStyle(nameCell, 'left', true)

    const industryCell = excelRow.getCell('C')
    industryCell.value = ''
    applyDataStyle(industryCell)

    const typeCell = excelRow.getCell('D')
    typeCell.value = row.enterpriseType ?? ''
    applyDataStyle(typeCell)

    const loanDate = parseExcelDate(row.actualLoanDate)
    const loanCell = excelRow.getCell('E')
    loanCell.value = loanDate ?? ''
    if (loanDate) {
      loanCell.numFmt = 'yyyy-mm-dd'
    }
    applyDataStyle(loanCell)

    const dueDate = parseExcelDate(row.dueDate)
    const dueCell = excelRow.getCell('F')
    dueCell.value = dueDate ?? ''
    if (dueDate) {
      dueCell.numFmt = 'yyyy-mm-dd'
    }
    applyDataStyle(dueCell)

    const principal = parseNumber(row.principalAmount)
    const principalCell = excelRow.getCell('G')
    if (principal !== null) {
      principalCell.value = principal / 10000
      principalCell.numFmt = '#,##0.00'
    } else {
      principalCell.value = ''
    }
    applyDataStyle(principalCell)

    const interestCell = excelRow.getCell('H')
    interestCell.value = ''
    applyDataStyle(interestCell)

    const serviceCell = excelRow.getCell('I')
    serviceCell.value = ''
    applyDataStyle(serviceCell)

    const rateValue = deriveRate(row)
    const rateCell = excelRow.getCell('J')
    rateCell.value = rateValue !== null ? rateValue : ''
    if (rateValue !== null) {
      rateCell.numFmt = '0.00"%"'
    }
    applyDataStyle(rateCell)

    const businessTypeYes = excelRow.getCell('K')
    businessTypeYes.value = '√'
    applyDataStyle(businessTypeYes)

    const businessTypeNo = excelRow.getCell('L')
    businessTypeNo.value = ''
    applyDataStyle(businessTypeNo)

    const classificationCell = excelRow.getCell('M')
    classificationCell.value = '正常'
    applyDataStyle(classificationCell)

    const creditCell = excelRow.getCell('N')
    creditCell.value = '无'
    applyDataStyle(creditCell)

    const remarkCell = excelRow.getCell('O')
    remarkCell.value = row.assetId ?? ''
    applyDataStyle(remarkCell, 'left', true)
  })

  const dvEndRow = dataStartRow + Math.max(filteredRows.length, 30)
  const worksheetWithValidation = worksheet as WorksheetWithValidation
  worksheetWithValidation.dataValidations.add(`C${dataStartRow}:C${dvEndRow}`, {
    type: 'list',
    allowBlank: true,
    formulae: [`'经济行业分类'!$A$2:$A$${ECONOMIC_INDUSTRY_CLASSIFICATIONS.length + 1}`]
  })

  const footnoteRow = worksheet.getRow(dataStartRow + filteredRows.length + 1)
  const footnoteCell = footnoteRow.getCell('A')
  footnoteCell.value = `注：
1. 经济行业分类：参照《国民经济行业分类》（GB/T 4754-2011）。
2. 企业类型：标准参照国家统计局印发的《统计上大中小微型企业划分办法（2017）》（国统字〔2017〕213号），企业划型适用行业按企业从事的主要经济活动确定。
3. 五级分类情况：参照中国人民银行关于贷款五级分类的相关要求确定。`
  footnoteCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true }
  footnoteCell.font = { size: 10, color: { argb: 'FF000000' } }
  footnoteRow.height = 120

  const dictionarySheet = workbook.addWorksheet('经济行业分类')
  dictionarySheet.columns = [{ width: 40 }]
  const title = dictionarySheet.getCell('A1')
  title.value = '经济行业分类下拉框选项参考'
  title.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFBDD7EE' }
  }
  title.font = { size: 11, bold: true }
  title.alignment = { horizontal: 'left', vertical: 'middle' }

  ECONOMIC_INDUSTRY_CLASSIFICATIONS.forEach((name, idx) => {
    const cell = dictionarySheet.getCell(`A${idx + 2}`)
    cell.value = name
    cell.font = { size: 11 }
    cell.alignment = { horizontal: 'left', vertical: 'middle' }
  })

  await workbook.xlsx.writeFile(outputPath)
  console.log(`[f103FactoringDetail] ExcelJS 渲染完成: ${outputPath}`)
}

const inputRules: FormCreateRule[] = [
  {
    type: 'Input',
    field: 'fillMonth',
    title: '填报年月',
    value: new Date().toISOString().slice(0, 7).replace('-', ''),
    props: {
      placeholder: '请输入填报年月（YYYYMM，例如 202501）'
    },
    validate: [
      { required: true, message: '请输入填报年月', trigger: 'blur' },
      {
        pattern: '^[0-9]{6}$',
        message: '格式需为 YYYYMM，例如 202501',
        trigger: 'blur'
      }
    ]
  }
]

export const f103FactoringDetailTemplate: TemplateDefinition<FactoringUserInput> = {
  meta: {
    id: 'f103FactoringDetail',
    name: '融资保理业务明细表',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description:
      '基于放款明细中的实际放款日期，生成国富 F103 融资保理业务明细表，并附带经济行业分类下拉选项字典',
    sourceLabel: '放款明细表'
  },
  engine: 'exceljs',
  inputRule: {
    rules: inputRules,
    options: {
      labelWidth: '120px',
      labelPosition: 'left',
      submitBtn: false,
      resetBtn: false
    },
    example: {
      fillMonth: '202501'
    },
    description: `
### 填报说明
- **填报年月**：格式为 YYYYMM，例如 202501，用于筛选 P 列实际放款日期的年份和月份。
- 报表自动引用放款明细的 C/F/P/AP/AW/BG/BF/K 列；部分列按规定填写固定值。
- C 列的经济行业分类可在生成的“经济行业分类”sheet 中进行选择。
    `.trim()
  },
  parser: parseWorkbook,
  streamParser: streamParseWorkbook,
  excelRenderer: renderWithExcelJS
}
