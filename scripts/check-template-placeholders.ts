/**
 * 检查模板文件中的占位符
 */

import ExcelJS from 'exceljs'
import path from 'node:path'

async function checkTemplatePlaceholders(): Promise<void> {
  const projectRoot = process.cwd()
  const templatePath = path.join(projectRoot, 'public/reportTemplates/month1carbone.xlsx')

  console.log('📄 读取模板文件:', templatePath)
  console.log()

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(templatePath)

  console.log('📊 模板工作表列表:')
  workbook.worksheets.forEach((sheet, index) => {
    console.log(`  [${index}] ${sheet.name}`)
  })
  console.log()

  // 遍历所有工作表
  workbook.worksheets.forEach((sheet) => {
    console.log(`\n=== 工作表: ${sheet.name} ===\n`)

    const placeholders = new Set<string>()

    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        const value = cell.value

        // 检查字符串类型的单元格
        if (typeof value === 'string') {
          // 匹配 Carbone 占位符 {xxx}
          const matches = value.match(/\{[^}]+\}/g)
          if (matches) {
            matches.forEach((match) => placeholders.add(match))
          }
        }

        // 检查富文本
        if (value && typeof value === 'object' && 'richText' in value) {
          const richText = value as ExcelJS.CellRichTextValue
          richText.richText.forEach((textObj) => {
            const matches = textObj.text.match(/\{[^}]+\}/g)
            if (matches) {
              matches.forEach((match) => placeholders.add(match))
            }
          })
        }
      })
    })

    if (placeholders.size > 0) {
      console.log('找到的占位符:')
      Array.from(placeholders)
        .sort()
        .forEach((placeholder) => {
          console.log(`  ${placeholder}`)
        })
    } else {
      console.log('  (未找到占位符)')
    }
  })

  console.log('\n✅ 检查完成')
}

checkTemplatePlaceholders().catch((error) => {
  console.error('❌ 错误:', error)
  process.exit(1)
})
