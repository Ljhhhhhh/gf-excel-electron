const ExcelJS = require('exceljs')

async function analyzeLedgerTemplate() {
  const filePath = './public/demo/4.台账模板.xlsx'
  const targetSheetName = '融资及还款明细'

  const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    sharedStrings: 'cache',
    hyperlinks: 'cache',
    styles: 'cache',
    worksheets: 'emit'
  })

  console.log('=== 开始流式读取 ===\n')

  let sheetIndex = 0
  const allSheets = []

  for await (const worksheetReader of reader) {
    const sheetName = worksheetReader.name
    allSheets.push(sheetName)

    if (sheetName !== targetSheetName) {
      console.log(`跳过工作表 ${sheetIndex + 1}: ${sheetName}`)
      sheetIndex++
      continue
    }

    console.log(`\n=== 找到目标工作表: ${sheetName} ===\n`)

    let rowNumber = 0
    let row10Data = null

    for await (const row of worksheetReader) {
      rowNumber++

      // 只分析前15行
      if (rowNumber > 15) {
        break
      }

      console.log(`\n--- 第${rowNumber}行 ---`)

      // 保存第10行（模板行）
      if (rowNumber === 10) {
        row10Data = row
      }

      // 分析A-L列
      for (let colIdx = 1; colIdx <= 12; colIdx++) {
        const cell = row.getCell(colIdx)
        const colLetter = String.fromCharCode(64 + colIdx) // A=65

        if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
          const info = [`${colLetter}${rowNumber}:`]

          // 值信息
          if (typeof cell.value === 'object' && cell.value !== null) {
            if (cell.value.formula) {
              info.push(`公式=[${cell.value.formula}]`)
              info.push(`结果=${cell.value.result}`)
            } else {
              info.push(`值=${JSON.stringify(cell.value)}`)
            }
          } else {
            info.push(`值=${cell.value}`)
          }

          // 样式信息（仅第10行显示）
          if (rowNumber === 10) {
            if (cell.style) {
              if (cell.style.alignment) {
                info.push(`对齐=${JSON.stringify(cell.style.alignment)}`)
              }
              if (cell.style.font) {
                info.push(`字体=${JSON.stringify(cell.style.font)}`)
              }
              if (cell.style.fill) {
                info.push(`填充=${JSON.stringify(cell.style.fill)}`)
              }
              if (cell.style.border) {
                info.push(`边框=已设置`)
              }
              if (cell.style.numFmt) {
                info.push(`格式=${cell.style.numFmt}`)
              }
            }
          }

          console.log('  ' + info.join(' | '))
        }
      }

      // 显示行高
      if (row.height) {
        console.log(`  行高: ${row.height}`)
      }
    }

    console.log(`\n=== 列宽信息 ===`)
    const columns = worksheetReader.columns || []
    columns.slice(0, 12).forEach((col, idx) => {
      const colLetter = String.fromCharCode(65 + idx)
      console.log(`列${colLetter}: 宽度=${col?.width || '未设置'}`)
    })

    console.log(`\n总共读取了前 ${rowNumber} 行`)

    sheetIndex++
    break // 找到目标工作表后退出
  }

  console.log('\n=== 所有工作表列表 ===')
  allSheets.forEach((name, idx) => {
    console.log(`${idx + 1}. ${name}`)
  })
}

analyzeLedgerTemplate().catch((err) => {
  console.error('错误:', err)
  process.exit(1)
})
