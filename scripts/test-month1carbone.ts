/**
 * 测试 month1carbone 模板的数据处理逻辑
 */

import ExcelJS from 'exceljs'
import { join } from 'path'
import {
  parseWorkbook,
  buildReportData,
  type ReportInput
} from '../src/main/services/templates/month1carbone'

async function testMonth1Carbone(): Promise<void> {
  console.log('=== 测试 month1carbone 模板 ===\n')

  // 1. 读取数据源文件
  const sourcePath = join(process.cwd(), 'public/demo/放款明细.xlsx')
  console.log('数据源文件:', sourcePath)

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(sourcePath)

  // 2. 解析 workbook
  console.log('\n--- 解析数据 ---')
  const parsedData = parseWorkbook(workbook)
  console.log('总行数:', parsedData.summary?.totalRows)
  console.log('表头字段数:', parsedData.headers.length)
  console.log('前5个字段:', parsedData.headers.slice(0, 5))

  // 查看有多少条 2025年10月 的数据
  const oct2025Rows = parsedData.rows.filter((row) => {
    const dateStr = row['实际放款日期'] as string | undefined
    if (!dateStr) return false
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/)
    if (!match) return false
    return parseInt(match[1], 10) === 2025 && parseInt(match[2], 10) === 10
  })
  console.log('\n2025年10月的数据行数:', oct2025Rows.length)

  // 3. 构建报表数据
  console.log('\n--- 构建报表数据 (2025年10月) ---')
  const userInput: ReportInput = {
    queryYear: 2025,
    queryMonth: 10
  }

  const reportData = buildReportData(parsedData, userInput)
  console.log('\n生成的报表数据:')
  console.log(JSON.stringify(reportData, null, 2))

  // 4. 验证数据
  console.log('\n--- 数据验证 ---')
  const data = reportData as {
    v: { queryYear: number; queryMonth: number }
    d: {
      total: string
      infrastructurePercentage: string
      medicinePercentage: string
      factoringPercentage: string
      accountsReceivable: string
      customerTotal: number
      businessCount: number
    }
  }

  console.log('查询年月:', `${data.v.queryYear}年${data.v.queryMonth}月`)
  console.log('放款总额:', data.d.total, '万元')
  console.log('基建工程占比:', data.d.infrastructurePercentage, '%')
  console.log('医药医疗占比:', data.d.medicinePercentage, '%')
  console.log('大宗商品占比:', data.d.factoringPercentage, '%')
  console.log('应收账款:', data.d.accountsReceivable, '万元')
  console.log('客户数:', data.d.customerTotal)
  console.log('业务笔数:', data.d.businessCount)

  // 验证百分比总和
  const totalPercentage =
    parseFloat(data.d.infrastructurePercentage) +
    parseFloat(data.d.medicinePercentage) +
    parseFloat(data.d.factoringPercentage)
  console.log('\n三个行业占比总和:', totalPercentage.toFixed(2), '%')

  console.log('\n=== 测试完成 ===')
}

testMonth1Carbone().catch((error) => {
  console.error('测试失败:', error)
  process.exit(1)
})
