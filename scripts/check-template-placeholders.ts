/**
 * 检查模板占位符
 */

import ExcelJS from 'exceljs'
import { join } from 'path'

async function checkPlaceholders(): Promise<void> {
  const templatePath = join(process.cwd(), 'public/reportTemplates/month1carbone.xlsx')
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(templatePath)
  
  const ws = workbook.worksheets[0]
  console.log('工作表名称:', ws.name)
  console.log('\n前10行的内容:')
  
  for (let rowNum = 1; rowNum <= 10; rowNum++) {
    const row = ws.getRow(rowNum)
    let hasContent = false
    let rowContent = `第${rowNum}行: `
    
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const value = cell.value
      if (value) {
        hasContent = true
        rowContent += `[列${colNumber}] ${JSON.stringify(value)} `
      }
    })
    
    if (hasContent) {
      console.log(rowContent)
    }
  }
  
  // 特别检查 A1 单元格
  console.log('\n=== A1 单元格详细信息 ===')
  const cellA1 = ws.getCell('A1')
  console.log('原始值:', cellA1.value)
  console.log('类型:', typeof cellA1.value)
  
  if (cellA1.value && typeof cellA1.value === 'object') {
    console.log('对象键:', Object.keys(cellA1.value))
    console.log('完整对象:', JSON.stringify(cellA1.value, null, 2))
  }
}

checkPlaceholders().catch(console.error)
