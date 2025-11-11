/**
 * 调试 Carbone 渲染
 */

import carbone from 'carbone'
import fs from 'node:fs'
import path from 'node:path'
import { promisify } from 'node:util'

const renderAsync = promisify(carbone.render)

async function debugCarbone(): Promise<void> {
  const projectRoot = process.cwd()
  const templatePath = path.join(projectRoot, 'public/reportTemplates/month1carbone.xlsx')
  const outputPath = path.join(projectRoot, 'output/debug-carbone-test.xlsx')

  console.log('=== Carbone 调试 ===\n')
  console.log('模板路径:', templatePath)
  console.log('输出路径:', outputPath)
  console.log()

  // 简单的测试数据
  const testData = {
    v: {
      queryYear: 2025,
      queryMonth: 10
    },
    d: {
      total: '10485.42',
      infrastructurePercentage: '73.24',
      medicinePercentage: '26.76',
      factoringPercentage: '0.00',
      accountsReceivable: '10997.75',
      customerTotal: 7,
      businessCount: 78,
      factoringRate: '6.5'
    },
    medicinePercentage: '26.76'
  }

  console.log('测试数据:')
  console.log(JSON.stringify(testData, null, 2))
  console.log()

  try {
    console.log('正在渲染...')
    const options = {
      lang: 'zh-cn',
      timezone: 'Asia/Shanghai'
    }

    const result = await renderAsync(templatePath, testData, options)
    const buffer = result as Buffer

    console.log('✅ 渲染成功')
    console.log('Buffer 大小:', buffer.length, 'bytes')

    // 写入文件
    fs.writeFileSync(outputPath, buffer)
    console.log('✅ 文件已写入:', outputPath)

    // 读取生成的文件并检查内容
    console.log('\n正在验证生成的文件...')
    const ExcelJS = (await import('exceljs')).default
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(outputPath)

    const ws = workbook.worksheets[0]
    const cellA1 = ws.getCell('A1')
    console.log('\nA1 单元格内容:')
    console.log(cellA1.value)
    console.log()

    // 检查是否还有占位符
    const content = String(cellA1.value || '')
    const hasPlaceholders = content.includes('{v.') || content.includes('{d.')
    const hasEmptyBraces = content.match(/{\s*}/g)

    if (hasPlaceholders) {
      console.log('⚠️  警告：仍然存在未替换的占位符')
    } else if (hasEmptyBraces) {
      console.log('⚠️  警告：存在空的 {} 标记，数量:', hasEmptyBraces.length)
    } else {
      console.log('✅ 占位符已全部替换')
    }
  } catch (error) {
    console.error('❌ 渲染失败:', error)
    throw error
  }
}

debugCarbone().catch((error) => {
  console.error('调试失败:', error)
  process.exit(1)
})
