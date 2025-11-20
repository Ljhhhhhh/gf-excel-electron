const ExcelJS = require('exceljs')

async function analyzeDetailed() {
  const filePath = './public/demo/4.台账模板.xlsx'

  // 使用普通读取来获取列宽等信息
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  const sheet = workbook.getWorksheet('融资及还款明细')

  if (!sheet) {
    console.log('未找到目标工作表')
    return
  }

  console.log('=== 列宽信息（A-AR列）===')
  const columns = [
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'G',
    'H',
    'I',
    'J',
    'K',
    'L',
    'M',
    'N',
    'O',
    'P',
    'Q',
    'R',
    'S',
    'T',
    'U',
    'V',
    'W',
    'X',
    'Y',
    'Z',
    'AA',
    'AB',
    'AC',
    'AD',
    'AE',
    'AF',
    'AG',
    'AH',
    'AI',
    'AJ',
    'AK',
    'AL',
    'AM',
    'AN',
    'AO',
    'AP',
    'AQ',
    'AR'
  ]
  columns.forEach((col) => {
    const column = sheet.getColumn(col)
    if (column.width) {
      console.log(`列${col}: 宽度=${column.width}`)
    }
  })

  console.log('\n=== 第10行（模板行）所有有值的单元格 ===')
  const row10 = sheet.getRow(10)
  row10.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const colLetter = String.fromCharCode(64 + colNumber)
    console.log(`\n${cell.address}:`)

    // 值信息
    if (typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value) {
      console.log('  公式:', cell.value.formula)
      console.log('  结果:', cell.value.result)
    } else {
      console.log('  值:', cell.value)
    }

    // 样式信息
    if (cell.style) {
      if (cell.style.alignment) {
        console.log('  对齐:', JSON.stringify(cell.style.alignment))
      }
      if (cell.style.font) {
        console.log('  字体:', JSON.stringify(cell.style.font))
      }
      if (cell.style.fill) {
        console.log('  填充:', JSON.stringify(cell.style.fill))
      }
      if (cell.style.border) {
        console.log('  边框:', JSON.stringify(cell.style.border))
      }
      if (cell.style.numFmt) {
        console.log('  数字格式:', cell.style.numFmt)
      }
    }
  })

  console.log('\n=== 第3行A列详细信息（公式示例）===')
  const a3 = sheet.getRow(3).getCell('A')
  console.log('A3值类型:', typeof a3.value)
  console.log('A3值:', JSON.stringify(a3.value, null, 2))

  console.log('\n=== 第6行A列详细信息（问题行）===')
  const a6 = sheet.getRow(6).getCell('A')
  console.log('A6值类型:', typeof a6.value)
  console.log('A6值:', JSON.stringify(a6.value, null, 2))

  console.log('\n=== 合并单元格信息 ===')
  if (sheet._merges) {
    Object.keys(sheet._merges).forEach((key) => {
      console.log('合并:', key, '=', sheet._merges[key])
    })
  }

  console.log('\n=== 工作表保护信息 ===')
  if (sheet.properties) {
    console.log('工作表属性:', JSON.stringify(sheet.properties, null, 2))
  }
}

analyzeDetailed().catch((err) => {
  console.error('错误:', err)
  process.exit(1)
})
