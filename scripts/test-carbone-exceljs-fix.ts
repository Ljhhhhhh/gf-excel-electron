/**
 * 测试 Carbone 生成的文件经过 xlsx-js-style 规范化后是否可以被 ExcelJS 读取
 */

import { dataToReport } from '../src/main/services/dataToReport'
import { parseWorkbook } from '../src/main/services/templates/month2carbone'
import { initTemplates } from '../src/main/services/templates'
import path from 'node:path'
import fs from 'node:fs'
import ExcelJS from 'exceljs'

async function testCarboneExcelJSFix() {
  // 初始化模板系统
  initTemplates()
  console.log('='.repeat(60))
  console.log('测试：Carbone + xlsx-js-style 规范化 + ExcelJS 读取')
  console.log('='.repeat(60))

  const sourceFile = path.resolve(__dirname, '../public/demo/basic-source.xlsx')
  const outputDir = path.resolve(__dirname, '../output')
  const templateId = 'month2carbone'

  console.log(`\n1. 源数据文件: ${sourceFile}`)
  console.log(`2. 输出目录: ${outputDir}`)
  console.log(`3. 模板ID: ${templateId}`)

  // 确保输出目录存在
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  // 解析源数据
  console.log('\n[步骤1] 解析源数据文件...')
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(sourceFile)
  const parsedData = parseWorkbook(workbook)
  console.log(`✓ 解析完成，数据行数: ${parsedData.rows.length}`)

  // 用户输入参数
  const userInput = {
    queryYear: 2025,
    queryMonth: 9 // 测试9个月的数据
  }

  // 生成报表
  console.log('\n[步骤2] 使用 Carbone 生成报表...')
  const result = await dataToReport({
    templateId,
    parsedData,
    outputDir,
    reportName: `test-fix-${Date.now()}.xlsx`,
    userInput
  })
  console.log(`✓ 报表生成成功: ${result.outputPath}`)
  console.log(`  文件大小: ${(result.size / 1024).toFixed(2)} KB`)

  // 验证：尝试用 ExcelJS 读取生成的文件
  console.log('\n[步骤3] 验证：使用 ExcelJS 读取生成的文件...')
  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(result.outputPath)

    const worksheet = workbook.getWorksheet('Sheet1')
    if (!worksheet) {
      throw new Error('找不到工作表 Sheet1')
    }

    console.log(`✓ ExcelJS 读取成功！`)
    console.log(`  工作表名称: ${worksheet.name}`)
    console.log(`  行数: ${worksheet.rowCount}`)
    console.log(`  列数: ${worksheet.columnCount}`)

    // 检查一些关键单元格
    const cellA1 = worksheet.getCell('A1')
    const cellB1 = worksheet.getCell('B1')
    console.log(`  A1 单元格值: "${cellA1.value}"`)
    console.log(`  B1 单元格值: "${cellB1.value}"`)

    // 验证合并单元格
    const mergedCells =
      worksheet.getCell('A1').master === worksheet.getCell('A1') ? '未合并' : '已合并'
    console.log(`  A1:A2 合并状态: ${mergedCells}`)

    console.log('\n' + '='.repeat(60))
    console.log('✅ 测试通过！Carbone 文件已成功规范化，ExcelJS 可以正常读取')
    console.log('='.repeat(60))
  } catch (error) {
    console.error('\n❌ 测试失败！ExcelJS 无法读取文件')
    console.error('错误信息:', error)
    throw error
  }
}

// 运行测试
testCarboneExcelJSFix().catch((error) => {
  console.error('\n测试执行失败:', error)
  process.exit(1)
})
