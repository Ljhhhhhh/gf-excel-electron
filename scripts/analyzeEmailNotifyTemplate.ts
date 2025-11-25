/**
 * 分析 emailNotify.xlsx 模板结构
 * 用于确保开发方案与模板匹配
 */

import ExcelJS from 'exceljs'
import path from 'node:path'

const TEMPLATE_PATH = path.resolve(__dirname, '../public/reportTemplates/emailNotify.xlsx')

interface CellInfo {
  address: string
  value: unknown
  formula?: string
  type: string
}

async function analyzeTemplate() {
  console.log('='.repeat(60))
  console.log('分析模板文件:', TEMPLATE_PATH)
  console.log('='.repeat(60))

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(TEMPLATE_PATH)

  console.log(`\n工作簿信息:`)
  console.log(`  - 工作表数量: ${workbook.worksheets.length}`)
  console.log(`  - 工作表名称: ${workbook.worksheets.map((ws) => ws.name).join(', ')}`)

  for (const worksheet of workbook.worksheets) {
    console.log('\n' + '='.repeat(60))
    console.log(`工作表: "${worksheet.name}"`)
    console.log('='.repeat(60))
    console.log(`  - 行数: ${worksheet.rowCount}`)
    console.log(`  - 列数: ${worksheet.columnCount}`)

    // 分析前 20 行的内容
    console.log('\n前 20 行内容:')
    console.log('-'.repeat(60))

    for (let rowNum = 1; rowNum <= Math.min(20, worksheet.rowCount); rowNum++) {
      const row = worksheet.getRow(rowNum)
      if (!row.hasValues) {
        console.log(`  行 ${rowNum}: (空行)`)
        continue
      }

      const cells: CellInfo[] = []
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const cellInfo: CellInfo = {
          address: cell.address,
          value: cell.value,
          type: typeof cell.value
        }

        // 检查是否为公式
        if (cell.formula) {
          cellInfo.formula = cell.formula
        }

        // 检查 Carbone 标记
        if (typeof cell.value === 'string' && cell.value.includes('{')) {
          cellInfo.type = 'carbone-tag'
        }

        cells.push(cellInfo)
      })

      if (cells.length > 0) {
        console.log(`  行 ${rowNum}:`)
        cells.forEach((c) => {
          const valueStr = typeof c.value === 'string' ? `"${c.value}"` : JSON.stringify(c.value)
          const extra = c.formula ? ` [公式: ${c.formula}]` : ''
          const typeInfo = c.type === 'carbone-tag' ? ' [Carbone标记]' : ''
          console.log(`    ${c.address}: ${valueStr}${extra}${typeInfo}`)
        })
      }
    }

    // 专门提取所有 Carbone 标记
    console.log('\n所有 Carbone 标记:')
    console.log('-'.repeat(60))

    const carboneTags: { address: string; tag: string }[] = []
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const value = cell.value
        if (typeof value === 'string' && value.includes('{')) {
          carboneTags.push({
            address: cell.address,
            tag: value
          })
        }
      })
    })

    if (carboneTags.length > 0) {
      carboneTags.forEach((t) => {
        console.log(`  ${t.address}: ${t.tag}`)
      })
    } else {
      console.log('  (未找到 Carbone 标记)')
    }

    // 分析合并单元格
    console.log('\n合并单元格:')
    console.log('-'.repeat(60))

    const merges = worksheet.model.merges || []
    if (merges.length > 0) {
      merges.forEach((merge) => {
        console.log(`  ${merge}`)
      })
    } else {
      console.log('  (无合并单元格)')
    }

    // 分析列宽
    console.log('\n列宽设置:')
    console.log('-'.repeat(60))
    for (let col = 1; col <= Math.min(10, worksheet.columnCount); col++) {
      const column = worksheet.getColumn(col)
      console.log(`  列 ${col} (${String.fromCharCode(64 + col)}): 宽度=${column.width || '默认'}`)
    }
  }

  console.log('\n' + '='.repeat(60))
  console.log('模板分析完成')
  console.log('='.repeat(60))
}

analyzeTemplate().catch(console.error)
