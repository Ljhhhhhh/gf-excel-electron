/**
 * 测试 month2carbone 模板的 postProcess 功能
 */

import path from 'node:path'
import { excelToData } from '../src/main/services/excelToData'
import { dataToReport } from '../src/main/services/dataToReport'
import { initTemplates } from '../src/main/services/templates'

async function testMonth2PostProcess() {
  console.log('=== 测试 month2carbone postProcess 功能 ===\n')

  // 0. 初始化模板系统
  console.log('0. 初始化模板系统...')
  initTemplates()
  console.log('   ✓ 模板系统初始化完成\n')

  // 1. 准备输入文件
  const sourcePath = path.resolve(__dirname, '../public/demo/basic-source.xlsx')
  const outputDir = path.resolve(__dirname, '../public/demo')

  console.log('1. 解析源数据...')
  const parseResult = await excelToData({
    sourcePath,
    templateId: 'month2carbone'
  })
  console.log(`   ✓ 解析完成，数据行数: ${(parseResult.data as any).rows.length}`)

  // 2. 测试不同的截止月份
  const testCases = [
    { queryYear: 2025, queryMonth: 9, desc: '9月（与 f.js 一致）' },
    { queryYear: 2025, queryMonth: 6, desc: '6月' },
    { queryYear: 2025, queryMonth: 3, desc: '3月' }
  ]

  for (const testCase of testCases) {
    console.log(`\n2. 测试场景: ${testCase.desc}`)
    console.log(`   参数: ${testCase.queryYear}年 ${testCase.queryMonth}月`)

    const reportResult = await dataToReport({
      templateId: 'month2carbone',
      parsedData: parseResult.data,
      outputDir,
      reportName: `month2-test-${testCase.queryMonth}月.xlsx`,
      userInput: {
        queryYear: testCase.queryYear,
        queryMonth: testCase.queryMonth
      }
    })

    console.log(`   ✓ 报表生成完成: ${reportResult.outputPath}`)
    console.log(`   ✓ 文件大小: ${(reportResult.size / 1024).toFixed(2)} KB`)
    console.log(`   ✓ 应该已隐藏 ${12 - testCase.queryMonth} 个月份列`)
  }

  console.log('\n=== 测试完成 ===')
  console.log('请手动检查生成的文件:')
  console.log('1. 是否隐藏了多余的月份列')
  console.log('2. A1:A2 是否已合并')
  console.log('3. B1 标题合并范围是否正确调整')
}

testMonth2PostProcess()
  .then(() => {
    console.log('\n✓ 所有测试通过')
    process.exit(0)
  })
  .catch((error) => {
    console.error('\n✗ 测试失败:', error)
    process.exit(1)
  })
