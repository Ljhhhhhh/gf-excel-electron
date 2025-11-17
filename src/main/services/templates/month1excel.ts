import ExcelJS from 'exceljs'
import type { TemplateDefinition } from './types'
import {
  parseWorkbook as parseMonth1Workbook,
  streamParseWorkbook as streamMonth1Workbook,
  month1carboneTemplate
} from './month1carbone'

interface Month1ExcelUserInput {
  statYear: number
  statMonth: number
  queryYear: number
  queryMonth: number
  queryDay: number
}

interface Month1SummaryModel {
  monthlySummary: {
    statYear: number
    statMonth: string | number
    newLoanAmount: number
    newInfraRatio: number
    newMedicalRatio: number
    newRefactoringRatio: number
    newArAssignedAmount: number
    newCoopCustomerCount: number
    newAcceptedBusinessCount: number
  }
  asOfDateSummary: {
    queryYear: number
    queryMonth: string | number
    queryDay: string | number
    cumLoanAmount: number
    infraRatio: number
    medicalRatio: number
    refactoringRatio: number
  }
  sinceInceptionSummary: {
    cumAcceptedBusinessCount: number
    cumArAssignedAmount: number
    cumLoanAmount: number
    infraRatio: number
    medicalRatio: number
    refactoringRatio: number
  }
}

const COLUMN_WIDTH = 122
const ROW_HEIGHTS = [211.5, 151, 151] as const

function ensureSummaryModel(
  parsedData: unknown,
  userInput: Month1ExcelUserInput
): Month1SummaryModel {
  const builder = month1carboneTemplate.builder
  if (!builder) {
    throw new Error('[month1excel] month1carboneTemplate 缺少 builder')
  }

  const summary = builder(parsedData, userInput)
  if (!summary || typeof summary !== 'object') {
    throw new Error('[month1excel] 无法从 parser 结果中生成统计摘要')
  }

  return summary as Month1SummaryModel
}

function formatNumber(value: unknown, digits?: number): string {
  const num =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim() !== ''
        ? Number(value)
        : Number.NaN

  if (!Number.isFinite(num)) {
    if (typeof digits === 'number') {
      return (0).toFixed(digits)
    }
    return '0'
  }

  if (typeof digits === 'number') {
    return num.toFixed(digits)
  }

  return `${num}`
}

function formatInteger(value: unknown): string {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.round(value).toString()
  }
  if (typeof value === 'string' && value.trim() !== '') {
    return value.trim()
  }
  return '0'
}

function twoDigits(value: unknown): string {
  if (typeof value === 'string' && value.trim() !== '') {
    return value.padStart(2, '0')
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.trunc(value).toString().padStart(2, '0')
  }
  return '00'
}

function createParagraphs(summary: Month1SummaryModel): [string, string, string] {
  const { monthlySummary: monthly, asOfDateSummary: asOf, sinceInceptionSummary: since } = summary

  const paragraph1 = `${monthly.statYear}年${twoDigits(monthly.statMonth)}月公司新增放款【${formatNumber(
    monthly.newLoanAmount,
    2
  )}】万元（基建占比【${formatNumber(monthly.newInfraRatio, 2)}】%，医药占比【${formatNumber(
    monthly.newMedicalRatio,
    2
  )}】%，再保理占比【${formatNumber(monthly.newRefactoringRatio, 2)}】%），\n受让对应的应收账款【${formatNumber(
    monthly.newArAssignedAmount,
    2
  )}】万元，平均保理利率【】%，新增合作客户【${formatInteger(
    monthly.newCoopCustomerCount
  )}】家，受理业务【${formatInteger(
    monthly.newAcceptedBusinessCount
  )}】笔，已发生实际业务合作的核心企业新增【】家；`

  const paragraph2 = `截止${asOf.queryYear}年${twoDigits(asOf.queryMonth)}月${twoDigits(
    asOf.queryDay
  )}日，公司应收账款余额【】万元，放款余额【】万元，业务不良率为【】%,\n累计受让应收账款金额【】万元，累计放款【${formatNumber(
    asOf.cumLoanAmount,
    2
  )}】万元（基建占比【${formatNumber(asOf.infraRatio, 2)}】% ,医药占比【${formatNumber(
    asOf.medicalRatio,
    2
  )}】%,再保理占比【${formatNumber(asOf.refactoringRatio, 2)}】%）,累计还款【】万元；`

  const paragraph3 = `展业至今，公司合作客户【】家，累计受理业务【${formatInteger(
    since.cumAcceptedBusinessCount
  )}】笔，已发生实际业务合作的核心企业【】家，\n累计受让应收账款【${formatNumber(
    since.cumArAssignedAmount,
    2
  )}】亿元，累计放款【${formatNumber(since.cumLoanAmount, 2)}】亿元（基建占比【${formatNumber(
    since.infraRatio,
    2
  )}】%，医药占比【${formatNumber(since.medicalRatio, 2)}】%，再保理占比【${formatNumber(
    since.refactoringRatio,
    2
  )}】%）；放款余额【】亿元；业务不良率为【】%。`

  return [paragraph1, paragraph2, paragraph3]
}

async function renderWithExcelJS(
  parsedData: unknown,
  userInput: Month1ExcelUserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput) {
    throw new Error('[month1excel] 缺少用户输入参数')
  }

  const summary = ensureSummaryModel(parsedData, userInput)
  const paragraphs = createParagraphs(summary)

  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Sheet1')
  worksheet.columns = [{ width: COLUMN_WIDTH }]

  paragraphs.forEach((text, index) => {
    const rowIndex = index + 1
    const row = worksheet.getRow(rowIndex)
    row.height = ROW_HEIGHTS[index]
    const cell = row.getCell(1)
    cell.value = text
    cell.font = { size: 11 }
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true }
  })

  await workbook.xlsx.writeFile(outputPath)
  console.log(`[month1excel] ExcelJS 渲染完成: ${outputPath}`)
}

export const month1excelTemplate: TemplateDefinition<Month1ExcelUserInput> = {
  meta: {
    id: 'month1excel',
    name: '月报1',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description: '基于月报1统计逻辑的段落式文字模板，由 ExcelJS 直接生成',
    sourceLabel: month1carboneTemplate.meta.sourceLabel,
    sourceDescription: month1carboneTemplate.meta.sourceDescription,
    extraSources: month1carboneTemplate.meta.extraSources?.map((source) => ({ ...source }))
  },
  engine: 'exceljs',
  inputRule: month1carboneTemplate.inputRule,
  parser: parseMonth1Workbook,
  streamParser: streamMonth1Workbook,
  excelRenderer: renderWithExcelJS
}
