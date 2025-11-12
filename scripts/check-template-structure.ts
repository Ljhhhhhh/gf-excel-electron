/**
 * 检查 Carbone 模板的结构
 */

import ExcelJS from 'exceljs'
import path from 'node:path'

async function checkTemplateStructure() {
  const templatePath = path.resolve(__dirname, '../public/reportTemplates/month2carbone.xlsx')
  
  console.log('检查模板文件:', templatePath)
  console.log('='.repeat(60))
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(templatePath)
  
  const sheet = workbook.getWorksheet('Sheet1')
  if (!sheet) {
    console.error('找不到 Sheet1')
    return
  }
  
  console.log(`工作表: ${sheet.name}`)
  console.log(`行数: ${sheet.rowCount}`)
  console.log(`列数: ${sheet.columnCount}`)
  console.log('='.repeat(60))
  
  // 输出每一行的内容
  for (let rowNum = 1; rowNum <= sheet.rowCount; rowNum++) {
    const row = sheet.getRow(rowNum)
    const values: any[] = []
    
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber <= 14) { // 只显示前14列
        values.push(cell.value || '')
      }
    })
    
    console.log(`第${rowNum}行:`, values)
  }
}

checkTemplateStructure().catch(console.error)
