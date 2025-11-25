/**
 * 即期邮件提醒报表模板
 *
 * 数据源:
 *   1. 主数据源: 《即期提醒融资列表》
 *   2. 额外数据源 loanDetail: 《放款明细》
 *   3. 额外数据源 emailCollection: 《通商回款邮件提醒信息采集》
 *
 * 处理逻辑:
 *   1. 筛选《即期提醒融资列表》中 A 列（即期逾期天数）为 0/-1/-5/-10 的行
 *   2. 根据放款流水号到《放款明细》匹配（U列包含"通商"），回填银行/金额信息
 *   3. 从《邮件采集表》按客户名称匹配，回填邮箱字段
 *   4. 按客户名称分组，使用 ExcelJS 生成多 sheet 报表
 *
 * 注意: Carbone 不支持 XLSX 多 sheet，故使用 ExcelJS 引擎
 *
 * 输出: 即期提醒汇总_YYYYMMDD.xlsx
 */

import ExcelJS from 'exceljs'
import type { Workbook } from 'exceljs'
import type { TemplateDefinition, ParseOptions } from './types'

// ========== 常量定义 ==========

/** 额外数据源 ID */
const LOAN_DETAIL_SOURCE_ID = 'loanDetail'
const EMAIL_COLLECTION_SOURCE_ID = 'emailCollection'

/** 默认配置 */
const DEFAULT_DATA_START_ROW = 2 // 数据从第2行开始（第1行为标题）

/** 需要筛选的即期逾期天数值 */
const TARGET_OVERDUE_DAYS = new Set([0, -1, -5, -10])

/** 保理商名称（固定值，用于模板 A7 列） */
const FACTOR_NAME = '国富商业保理有限公司'

// ========== 列索引定义（1-based） ==========

/** 即期提醒融资列表列索引 */
const REMINDER_COLUMNS = {
  overdueDays: 1, // A 列：即期逾期天数（工作日）
  clientName: 3, // C 列：客户名称
  buyerName: 4, // D 列：基础交易对手方
  assetNo: 6, // F 列：资产编号
  loanSerialNo: 8, // H 列：放款流水号
  repaymentAmount: 9, // I 列：融资余额（应还融资款）
  financingDueDate: 10, // J 列：融资到期日
  arDueDate: 12, // L 列：应收账款到期日
  // 以下为新增列（需回填）
  bank: 15, // O 列：出款银行
  arTransferAmount: 16, // P 列：应收账款转让金额
  loanAmount: 17, // Q 列：放款金额
  clientAdminEmail: 18, // R 列：客户管理员邮箱
  managerEmail: 19, // S 列：业务经理邮箱
  headEmail: 20, // T 列：业务负责人邮箱
  contactEmail: 21 // U 列：客关对接人邮箱
} as const

/** 放款明细列索引 */
const LOAN_DETAIL_COLUMNS = {
  loanSerialNo: 13, // M 列：放款流水号
  bankName: 21, // U 列：放款账户开户行
  arTransferAmount: 44, // AR 列：转让总金额
  loanAmount: 49 // AW 列：放款金额
} as const

/** 邮件采集表列索引 */
const EMAIL_COLLECTION_COLUMNS = {
  clientName: 1, // A 列：客户名称
  clientAdminEmail: 2, // B 列：客户管理员邮箱
  managerEmail: 4, // D 列：业务经理邮箱
  headEmail: 6, // F 列：业务负责人邮箱
  contactEmail: 8 // H 列：客关对接人邮箱
} as const

// ========== 类型定义 ==========

/** 解析选项 */
interface EmailNotifyParseOptions extends ParseOptions {
  /** 数据起始行（默认为 2） */
  dataStartRow?: number
}

/** 放款明细匹配结果 */
interface LoanDetailMatch {
  bankName: string
  arTransferAmount: number
  loanAmount: number
}

/** 邮箱信息 */
interface EmailInfo {
  clientAdminEmail: string
  managerEmail: string
  headEmail: string
  contactEmail: string
}

/** 单条交易记录 */
interface Transaction {
  factorName: string // 保理商名称
  buyerName: string // 买方名称（基础交易对手方）
  assetNo: string // 资产编号
  arTransferAmount: number // 应收账款转让金额
  arDueDate: string // 应收账款到期日
  loanSerialNo: string // 放款流水号
  loanAmount: number // 放款金额
  financingDueDate: string // 融资到期日
  repaymentAmount: number // 应还融资款（融资余额）
}

/** 单个报表（按客户分组） */
interface ReportItem {
  clientName: string
  clientAdminEmail: string
  managerEmail: string
  headEmail: string
  contactEmail: string
  transactions: Transaction[]
}

/** 分组后的数据结构 */
interface GroupedReportData {
  reports: ReportItem[]
}

/** 解析后的中间数据 */
interface ParsedRow {
  // 原始字段
  overdueDays: number | null
  clientName: string
  buyerName: string
  assetNo: string
  loanSerialNo: string
  repaymentAmount: number
  financingDueDate: string
  arDueDate: string
  // 回填字段
  bank: string
  arTransferAmount: number
  loanAmount: number
  clientAdminEmail: string
  managerEmail: string
  headEmail: string
  contactEmail: string
}

interface EmailNotifyParsedData {
  rows: ParsedRow[]
}

// ========== 工具函数 ==========

/**
 * 从单元格提取数值
 */
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

/**
 * 从单元格提取字符串
 */
function toString(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : ''
  if (value instanceof Date) return formatDate(value)
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    if ('text' in record) return toString(record.text)
    if ('result' in record) return toString(record.result)
  }
  return String(value).trim()
}

/**
 * 格式化日期为 YYYY-MM-DD
 */
function formatDate(date: Date | null): string {
  if (!date || !Number.isFinite(date.getTime())) return ''
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

/**
 * 解析日期值
 */
function parseDate(value: unknown): Date | null {
  if (!value) return null
  if (value instanceof Date) {
    return Number.isFinite(value.getTime()) ? value : null
  }
  if (typeof value === 'number') {
    // Excel 序列号
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
    if ('result' in record) return parseDate(record.result)
    if ('text' in record) return parseDate(record.text)
  }
  return null
}

/**
 * 判断字符串是否包含"通商"（不区分大小写）
 */
function containsTongShang(value: string): boolean {
  return value.toLowerCase().includes('通商')
}

/**
 * 格式化当前日期为 YYYYMMDD
 */
function formatDateForFilename(): string {
  const now = new Date()
  const year = now.getFullYear()
  const month = String(now.getMonth() + 1).padStart(2, '0')
  const day = String(now.getDate()).padStart(2, '0')
  return `${year}${month}${day}`
}

// ========== 数据提取函数 ==========

/**
 * 构建放款明细匹配表
 * key: 放款流水号, value: 第一条 U 列包含"通商"的记录
 */
function buildLoanDetailMap(workbook: Workbook, startRow: number): Map<string, LoanDetailMatch> {
  const map = new Map<string, LoanDetailMatch>()
  const worksheet = workbook.worksheets[0]

  if (!worksheet) {
    console.warn('[emailNotify] 放款明细工作簿无工作表')
    return map
  }

  const endRow = worksheet.rowCount
  for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
    const row = worksheet.getRow(rowNum)
    if (!row.hasValues) continue

    const loanSerialNo = toString(row.getCell(LOAN_DETAIL_COLUMNS.loanSerialNo).value)
    const bankName = toString(row.getCell(LOAN_DETAIL_COLUMNS.bankName).value)

    // 仅处理 U 列包含"通商"的记录
    if (!loanSerialNo || !containsTongShang(bankName)) continue

    // 如果该放款流水号已有匹配，跳过（只取第一条）
    if (map.has(loanSerialNo)) continue

    map.set(loanSerialNo, {
      bankName,
      arTransferAmount: toNumber(row.getCell(LOAN_DETAIL_COLUMNS.arTransferAmount).value),
      loanAmount: toNumber(row.getCell(LOAN_DETAIL_COLUMNS.loanAmount).value)
    })
  }

  console.log(`[emailNotify] 放款明细匹配表构建完成，共 ${map.size} 条通商记录`)

  // 调试：打印前 3 条放款流水号样例
  const samples = Array.from(map.keys()).slice(0, 3)
  console.log(`[emailNotify] 调试 - 放款明细流水号样例: ${samples.map((s) => `"${s}"`).join(', ')}`)

  return map
}

/**
 * 构建邮箱信息匹配表
 * key: 客户名称, value: 邮箱信息
 */
function buildEmailMap(workbook: Workbook, startRow: number): Map<string, EmailInfo> {
  const map = new Map<string, EmailInfo>()
  const worksheet = workbook.worksheets[0]

  if (!worksheet) {
    console.warn('[emailNotify] 邮件采集表工作簿无工作表')
    return map
  }

  const endRow = worksheet.rowCount
  for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
    const row = worksheet.getRow(rowNum)
    if (!row.hasValues) continue

    const clientName = toString(row.getCell(EMAIL_COLLECTION_COLUMNS.clientName).value)
    if (!clientName) continue

    // 如果该客户名称已存在，跳过（只取第一条）
    if (map.has(clientName)) continue

    map.set(clientName, {
      clientAdminEmail: toString(row.getCell(EMAIL_COLLECTION_COLUMNS.clientAdminEmail).value),
      managerEmail: toString(row.getCell(EMAIL_COLLECTION_COLUMNS.managerEmail).value),
      headEmail: toString(row.getCell(EMAIL_COLLECTION_COLUMNS.headEmail).value),
      contactEmail: toString(row.getCell(EMAIL_COLLECTION_COLUMNS.contactEmail).value)
    })
  }

  console.log(`[emailNotify] 邮箱信息匹配表构建完成，共 ${map.size} 个客户`)
  return map
}

// ========== 解析器实现 ==========

/**
 * 解析即期提醒融资列表，并回填放款明细和邮箱信息
 */
export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: EmailNotifyParseOptions
): EmailNotifyParsedData {
  const startRow = parseOptions?.dataStartRow ?? DEFAULT_DATA_START_ROW

  // 1. 获取额外数据源
  const loanDetailSource = parseOptions?.extraSources?.[LOAN_DETAIL_SOURCE_ID]
  const emailCollectionSource = parseOptions?.extraSources?.[EMAIL_COLLECTION_SOURCE_ID]

  if (!loanDetailSource?.workbook) {
    throw new Error('[emailNotify] 缺少放款明细数据源')
  }
  if (!emailCollectionSource?.workbook) {
    throw new Error('[emailNotify] 缺少邮件采集表数据源')
  }

  // 2. 构建匹配表
  const loanDetailMap = buildLoanDetailMap(loanDetailSource.workbook, startRow)
  const emailMap = buildEmailMap(emailCollectionSource.workbook, startRow)

  // 3. 解析主数据源（即期提醒融资列表）
  const worksheet = workbook.worksheets[0]
  if (!worksheet) {
    throw new Error('[emailNotify] 即期提醒融资列表工作簿无工作表')
  }

  const rows: ParsedRow[] = []
  const endRow = worksheet.rowCount

  // 第一轮：筛选符合条件的行，并回填放款明细信息
  for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
    const row = worksheet.getRow(rowNum)
    if (!row.hasValues) continue

    // 筛选 A 列即期逾期天数
    const overdueDaysRaw = row.getCell(REMINDER_COLUMNS.overdueDays).value
    const overdueDays = toNumber(overdueDaysRaw)

    // 检查是否为目标逾期天数
    if (!TARGET_OVERDUE_DAYS.has(overdueDays)) continue

    const loanSerialNo = toString(row.getCell(REMINDER_COLUMNS.loanSerialNo).value)
    const clientName = toString(row.getCell(REMINDER_COLUMNS.clientName).value)

    // 从放款明细匹配银行和金额信息
    const loanMatch = loanDetailMap.get(loanSerialNo)

    // 调试：打印前几条匹配信息
    if (rows.length < 5) {
      console.log(
        `[emailNotify] 调试 - 行${rowNum}: 逾期天数=${overdueDays}, 放款流水号="${loanSerialNo}", 客户="${clientName}", 匹配结果=${loanMatch ? '找到' : '未找到'}`
      )
    }

    const parsedRow: ParsedRow = {
      overdueDays,
      clientName,
      buyerName: toString(row.getCell(REMINDER_COLUMNS.buyerName).value),
      assetNo: toString(row.getCell(REMINDER_COLUMNS.assetNo).value),
      loanSerialNo,
      repaymentAmount: toNumber(row.getCell(REMINDER_COLUMNS.repaymentAmount).value),
      financingDueDate: formatDate(parseDate(row.getCell(REMINDER_COLUMNS.financingDueDate).value)),
      arDueDate: formatDate(parseDate(row.getCell(REMINDER_COLUMNS.arDueDate).value)),
      // 回填字段
      bank: loanMatch?.bankName ?? '',
      arTransferAmount: loanMatch?.arTransferAmount ?? 0,
      loanAmount: loanMatch?.loanAmount ?? 0,
      // 邮箱字段稍后填充
      clientAdminEmail: '',
      managerEmail: '',
      headEmail: '',
      contactEmail: ''
    }

    rows.push(parsedRow)
  }

  console.log(`[emailNotify] 第一轮筛选完成，共 ${rows.length} 行符合逾期天数条件`)

  // 第二轮：为出款银行不为空的行补齐邮箱信息
  for (const row of rows) {
    if (!row.bank) continue // 出款银行为空则跳过

    const emailInfo = emailMap.get(row.clientName)
    if (emailInfo) {
      row.clientAdminEmail = emailInfo.clientAdminEmail
      row.managerEmail = emailInfo.managerEmail
      row.headEmail = emailInfo.headEmail
      row.contactEmail = emailInfo.contactEmail
    }
  }

  // 统计最终有效行数
  const validRows = rows.filter((r) => r.bank !== '')
  console.log(`[emailNotify] 解析完成，共 ${validRows.length} 行有出款银行信息`)

  return { rows }
}

// ========== 数据分组函数 ==========

/**
 * 将解析后的数据按客户分组
 */
function groupByClient(parsedData: unknown): GroupedReportData {
  const data = parsedData as EmailNotifyParsedData

  if (!data || !Array.isArray(data.rows)) {
    throw new Error('[emailNotify] 解析数据格式不正确')
  }

  // 1. 过滤：仅保留出款银行不为空的行
  const filteredRows = data.rows.filter((row) => row.bank !== '')

  if (filteredRows.length === 0) {
    console.warn('[emailNotify] 没有符合条件的数据行（出款银行均为空）')
    return { reports: [] }
  }

  // 2. 按客户名称分组
  const groupedByClient = new Map<string, ParsedRow[]>()

  for (const row of filteredRows) {
    const clientName = row.clientName
    if (!groupedByClient.has(clientName)) {
      groupedByClient.set(clientName, [])
    }
    groupedByClient.get(clientName)!.push(row)
  }

  console.log(`[emailNotify] 按客户分组完成，共 ${groupedByClient.size} 个客户`)

  // 3. 构建 reports 数组
  const reports: ReportItem[] = []

  for (const [clientName, clientRows] of groupedByClient) {
    // 取第一行的邮箱信息作为头部信息
    const firstRow = clientRows[0]

    const transactions: Transaction[] = clientRows.map((row) => ({
      factorName: FACTOR_NAME,
      buyerName: row.buyerName,
      assetNo: row.assetNo,
      arTransferAmount: row.arTransferAmount,
      arDueDate: row.arDueDate,
      loanSerialNo: row.loanSerialNo,
      loanAmount: row.loanAmount,
      financingDueDate: row.financingDueDate,
      repaymentAmount: row.repaymentAmount
    }))

    reports.push({
      clientName,
      clientAdminEmail: firstRow.clientAdminEmail,
      managerEmail: firstRow.managerEmail,
      headEmail: firstRow.headEmail,
      contactEmail: firstRow.contactEmail,
      transactions
    })
  }

  console.log(
    `[emailNotify] 数据分组完成，共 ${reports.length} 个客户，` +
      `${filteredRows.length} 条交易记录`
  )

  return { reports }
}

// ========== ExcelJS 渲染器实现 ==========

/**
 * 清理 sheet 名称（移除非法字符，限制长度）
 */
function sanitizeSheetName(name: string): string {
  // Excel sheet 名称不允许: \ / ? * [ ] :
  // 最大长度 31 字符
  return name
    .replace(/[\\/?*[\]:]/g, '_')
    .substring(0, 31)
    .trim()
}

/**
 * 设置单元格样式
 */
function setHeaderCellStyle(cell: ExcelJS.Cell, isTitle: boolean = false): void {
  cell.font = {
    bold: true,
    size: isTitle ? 12 : 11,
    color: { argb: 'FF000000' }
  }
  cell.alignment = {
    horizontal: 'left',
    vertical: 'middle'
  }
}

function setTableHeaderStyle(cell: ExcelJS.Cell): void {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFED7D30' } // 橙色背景
  }
  cell.font = {
    bold: true,
    size: 11,
    color: { argb: 'FFFFFFFF' } // 白色文字
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

function setDataCellStyle(cell: ExcelJS.Cell): void {
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

/**
 * 使用 ExcelJS 生成多 sheet 报表
 */
async function renderWithExcelJS(
  parsedData: unknown,
  _userInput: void | undefined,
  outputPath: string
): Promise<void> {
  console.log('[emailNotify] 开始 ExcelJS 渲染')

  // 1. 分组数据
  const { reports } = groupByClient(parsedData)

  if (reports.length === 0) {
    throw new Error('[emailNotify] 没有可生成的报表数据')
  }

  // 2. 创建工作簿
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'GF-Excel-Electron'
  workbook.created = new Date()

  // 3. 为每个客户创建一个 sheet
  for (const report of reports) {
    const sheetName = sanitizeSheetName(report.clientName)
    const worksheet = workbook.addWorksheet(sheetName)

    // 设置列宽
    worksheet.columns = [
      { width: 25 }, // A: 保理商名称
      { width: 24 }, // B: 买方名称
      { width: 19 }, // C: 资产编号
      { width: 18 }, // D: 应收账款转让金额
      { width: 16 }, // E: 应收账款到期日
      { width: 15 }, // F: 放款流水号
      { width: 14 }, // G: 放款金额
      { width: 14 }, // H: 融资到期日
      { width: 14 } // I: 应还融资款
    ]

    // 行 1: 客户名称
    const row1 = worksheet.getRow(1)
    row1.getCell(1).value = '客户名称：'
    row1.getCell(2).value = report.clientName
    setHeaderCellStyle(row1.getCell(1), true)
    setHeaderCellStyle(row1.getCell(2), true)

    // 行 2: 客户管理员邮箱
    const row2 = worksheet.getRow(2)
    row2.getCell(1).value = '客户管理员邮箱：'
    row2.getCell(2).value = report.clientAdminEmail
    setHeaderCellStyle(row2.getCell(1))
    setHeaderCellStyle(row2.getCell(2))

    // 行 3: 业务经理邮箱
    const row3 = worksheet.getRow(3)
    row3.getCell(1).value = '业务经理邮箱：'
    row3.getCell(2).value = report.managerEmail
    setHeaderCellStyle(row3.getCell(1))
    setHeaderCellStyle(row3.getCell(2))

    // 行 4: 业务负责人邮箱
    const row4 = worksheet.getRow(4)
    row4.getCell(1).value = '业务负责人邮箱：'
    row4.getCell(2).value = report.headEmail
    setHeaderCellStyle(row4.getCell(1))
    setHeaderCellStyle(row4.getCell(2))

    // 行 5: 客关对接人邮箱
    const row5 = worksheet.getRow(5)
    row5.getCell(1).value = '客关对接人邮箱：'
    row5.getCell(2).value = report.contactEmail
    setHeaderCellStyle(row5.getCell(1))
    setHeaderCellStyle(row5.getCell(2))

    // 行 6: 表头
    const headerRow = worksheet.getRow(6)
    const headers = [
      '保理商名称',
      '买方名称',
      '资产编号',
      '应收账款转让金额',
      '应收账款到期日',
      '放款流水号',
      '放款金额',
      '融资到期日',
      '应还融资款'
    ]
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1)
      cell.value = header
      setTableHeaderStyle(cell)
    })
    headerRow.height = 25

    // 行 7+: 数据行
    report.transactions.forEach((tx, txIndex) => {
      const dataRow = worksheet.getRow(7 + txIndex)
      const values = [
        tx.factorName,
        tx.buyerName,
        tx.assetNo,
        tx.arTransferAmount,
        tx.arDueDate,
        tx.loanSerialNo,
        tx.loanAmount,
        tx.financingDueDate,
        tx.repaymentAmount
      ]
      values.forEach((value, colIndex) => {
        const cell = dataRow.getCell(colIndex + 1)
        cell.value = value
        setDataCellStyle(cell)
        // 金额列设置数字格式
        if (colIndex === 3 || colIndex === 6 || colIndex === 8) {
          cell.numFmt = '#,##0.00'
        }
      })
    })

    console.log(`[emailNotify] Sheet "${sheetName}" 创建完成，${report.transactions.length} 条记录`)
  }

  // 4. 写入文件
  await workbook.xlsx.writeFile(outputPath)
  console.log(`[emailNotify] ExcelJS 渲染完成: ${outputPath}`)
}

// ========== 模板定义 ==========

export const emailNotifyTemplate: TemplateDefinition<void> = {
  meta: {
    id: 'emailNotify',
    name: '即期邮件提醒',
    // ExcelJS 模式不需要模板文件，由代码动态生成
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description:
      '根据即期提醒融资列表、放款明细、邮件采集表生成按客户分组的即期提醒汇总报表（多 sheet）',
    sourceLabel: '即期提醒融资列表',
    sourceDescription: '请选择《即期提醒融资列表》Excel 文件（第一张工作表）',
    extraSources: [
      {
        id: LOAN_DETAIL_SOURCE_ID,
        label: '放款明细',
        description: '请选择《放款明细》Excel 文件（第一张工作表）',
        required: true,
        supportedExts: ['xlsx'],
        loadStrategy: 'workbook'
      },
      {
        id: EMAIL_COLLECTION_SOURCE_ID,
        label: '通商回款邮件提醒信息采集',
        description: '请选择《通商回款邮件提醒信息采集》Excel 文件（第一张工作表）',
        required: true,
        supportedExts: ['xlsx'],
        loadStrategy: 'workbook'
      }
    ]
  },
  engine: 'exceljs', // 使用 ExcelJS 引擎（Carbone 不支持 XLSX 多 sheet）
  // 此模板不需要用户额外输入参数
  inputRule: undefined,
  parser: parseWorkbook,
  excelRenderer: renderWithExcelJS,
  resolveReportName: () => {
    // 生成文件名：即期提醒汇总_YYYYMMDD.xlsx
    return `即期提醒汇总_${formatDateForFilename()}.xlsx`
  }
}
