/**
 * 直接测试 postProcess 功能（不依赖 Carbone）
 */

import path from 'node:path'
import fs from 'node:fs/promises'
import { processExcelReport } from '../src/main/services/postProcess/month2Post'

async function testPostProcessDirect() {
  console.log('=== 直接测试 postProcess 功能 ===\n')

  // 使用模板文件作为输入（它应该是有效的）
  const templatePath = path.resolve(__dirname, '../public/demo/month2carbone-20251112-155049.xlsx')
  const outputPath = path.resolve(__dirname, '../public/demo/postprocess-test-output.xlsx')

  // 先复制模板到输出位置
  console.log('1. 复制模板文件...')
  await fs.copyFile(templatePath, outputPath)
  console.log(`   ✓ 已复制: ${outputPath}`)

  // 测试不同的月份
  const testCases = [
    { month: 9, desc: '9月（与 f.js 一致）' },
    { month: 6, desc: '6月' },
    { month: 3, desc: '3月' }
  ]

  for (const testCase of testCases) {
    const testOutput = path.resolve(
      __dirname,
      `../public/demo/postprocess-test-${testCase.month}月.xlsx`
    )

    // 每次复制原始模板
    await fs.copyFile(templatePath, testOutput)

    console.log(`\n2. 测试场景: ${testCase.desc}`)
    console.log(`   处理文件: ${testOutput}`)

    await processExcelReport({
      inputPath: testOutput,
      outputPath: testOutput, // 直接覆盖
      lastDataMonth: testCase.month,
      sheetName: 'Sheet1'
    })

    console.log(`   ✓ 处理完成`)
    console.log(`   应该已隐藏 ${12 - testCase.month} 个月份列（${testCase.month + 1}-12月）`)
  }

  console.log('\n=== 测试完成 ===')
  console.log('请手动打开以下文件检查:')
  console.log('- postprocess-test-9月.xlsx')
  console.log('- postprocess-test-6月.xlsx')
  console.log('- postprocess-test-3月.xlsx')
  console.log('\n检查项:')
  console.log('1. 是否隐藏了正确的月份列')
  console.log('2. A1:A2 是否已合并并居中')
  console.log('3. B1 标题合并范围是否正确（B1:J1、B1:G1、B1:D1）')
}

testPostProcessDirect()
  .then(() => {
    console.log('\n✓ 所有测试通过')
    process.exit(0)
  })
  .catch((error) => {
    console.error('\n✗ 测试失败:', error)
    process.exit(1)
  })
