/**
 * 临时脚本：检查数据源文件的列名
 */

import ExcelJS from 'exceljs'
import { join } from 'path'

async function checkColumns() {
  const filePath = join(process.cwd(), 'public/demo/放款明细.xlsx')

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  const ws = workbook.worksheets[0]
  console.log('工作表名称:', ws.name)
  console.log('总行数:', ws.rowCount)
  console.log('总列数:', ws.columnCount)

  // 读取第一行表头
  const headerRow = ws.getRow(1)
  const headers: string[] = []

  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const value = cell.value
    const text = value !== null && value !== undefined ? String(value).trim() : ''
    if (text) {
      // 显示列号（字母）和列索引
      const colLetter = getColumnLetter(colNumber)
      headers.push(`${colLetter}列(${colNumber}): ${text}`)
    }
  })

  console.log('\n表头列表:')
  headers.forEach((h) => console.log(h))

  // 显示第2行的示例数据
  console.log('\n第2行示例数据:')
  const row2 = ws.getRow(2)
  row2.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const colLetter = getColumnLetter(colNumber)
    console.log(`${colLetter}列:`, cell.value)
  })
}

function getColumnLetter(colNumber: number): string {
  let temp = colNumber
  let letter = ''
  while (temp > 0) {
    const mod = (temp - 1) % 26
    letter = String.fromCharCode(65 + mod) + letter
    temp = Math.floor((temp - mod) / 26)
  }
  return letter
}

checkColumns().catch(console.error)
