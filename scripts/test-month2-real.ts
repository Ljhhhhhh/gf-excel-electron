/**
 * 测试 month2exceljs 模板（使用真实放款明细数据）
 * 使用方法：npx tsx scripts/test-month2-real.ts
 */

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { excelToData } from '../src/main/services/excelToData'
import { dataToReport } from '../src/main/services/dataToReport'
import { initTemplates } from '../src/main/services/templates'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

async function main() {
  console.log('========== 测试 month2exceljs 模板（放款明细数据） ==========\n')

  // 1. 初始化模板系统
  console.log('[1/4] 初始化模板系统...')
  initTemplates()

  // 2. 解析源数据文件
  console.log('[2/4] 解析源数据文件（放款明细.xlsx）...')
  const sourcePath = path.resolve(__dirname, '../public/demo/放款明细.xlsx')
  const parseResult = await excelToData({
    sourcePath,
    templateId: 'month2exceljs',
    parseOptions: {
      dataStartRow: 2 // 第1行是标题，第2行开始是数据
    }
  })

  // 类型断言
  const data = parseResult.data as any
  console.log(`  ✓ 解析完成，共 ${data.rows?.length || 0} 条记录`)

  // 显示前5条数据样本
  if (data.rows && data.rows.length > 0) {
    console.log('\n  数据样本（前5条）：')
    data.rows.slice(0, 5).forEach((row: any, idx: number) => {
      console.log(
        `    ${idx + 1}. 日期: ${row.实际放款日期}, 行业: ${row.所属行业}, 金额: ${row.放款金额}`
      )
    })
  }

  // 3. 生成报表
  console.log('\n[3/4] 生成报表（使用 ExcelJS 渲染）...')
  const outputDir = path.resolve(__dirname, '../output')
  const reportResult = await dataToReport({
    templateId: 'month2exceljs',
    parsedData: parseResult.data,
    outputDir,
    reportName: `month2-放款明细-${Date.now()}.xlsx`,
    userInput: {
      queryYear: 2025,
      endMonth: 3
    }
  })
  console.log(`  ✓ 报表生成成功`)
  console.log(`  ✓ 输出路径: ${reportResult.outputPath}`)
  console.log(`  ✓ 文件大小: ${(reportResult.size / 1024).toFixed(2)} KB`)

  // 4. 完成
  console.log('\n========== 测试完成 ==========')
  console.log(`\n请打开以下文件查看结果：\n  ${reportResult.outputPath}\n`)
  console.log('预期结果：')
  console.log('  - 6行数据（1:标题, 2:月份表头, 3-5:行业数据, 6:合计）')
  console.log('  - 4列（A列行业名 + B-D列1-3月数据）')
  console.log('  - 橙色背景（#ED7D30）：第1、2、6行')
  console.log('  - 浅橙色背景（#FCECE8）：第3、4、5行')
}

main().catch((error) => {
  console.error('❌ 测试失败:', error)
  process.exit(1)
})
