/**
 * Excel 报表后处理工具函数
 * 提供对 Carbone 生成的 Excel 文件进行额外处理的能力（如隐藏列、合并单元格等）
 */

import ExcelJS from 'exceljs'

/**
 * 处理 Excel 报告的通用配置接口
 */
export interface ProcessExcelReportOptions {
  /** 输入文件路径 */
  inputPath: string
  /** 输出文件路径 */
  outputPath: string
  /** 最后一个有数据的月份 (1-12) */
  lastDataMonth: number
  /** 工作表名称（默认 'Sheet1'） */
  sheetName?: string
}

/**
 * 处理 Excel 报告
 * 功能：
 * 1. 隐藏超出数据范围的月份列（lastDataMonth 之后的列）
 * 2. 合并 A1:A2 单元格并居中
 * 3. 根据实际数据月份调整标题合并范围（B1:XN1，XN 为最后数据月份对应列）
 *
 * @param options 处理配置
 */
export async function processExcelReport(options: ProcessExcelReportOptions): Promise<void> {
  const { inputPath, outputPath, lastDataMonth, sheetName = 'Sheet1' } = options

  try {
    console.log(`[postProcess] 开始处理: ${inputPath} -> ${outputPath}`)
    console.log(`[postProcess] 数据截止月份: ${lastDataMonth}`)

    // 1. 创建工作簿实例并加载文件
    const workbook = new ExcelJS.Workbook()

    // 读取文件（应该已经被 dataToReport 规范化过）
    await workbook.xlsx.readFile(inputPath)
    console.log(`[postProcess] 文件读取成功`)

    // 获取指定工作表
    const sheet = workbook.getWorksheet(sheetName)

    if (!sheet) {
      const sheetNames = workbook.worksheets.map((ws) => ws.name).join(', ')
      throw new Error(`找不到名为 "${sheetName}" 的工作表。可用工作表: ${sheetNames}`)
    }

    console.log(`[postProcess] 工作表 "${sheet.name}" 加载成功，行数: ${sheet.rowCount}`)

    // 定义月份对应的列索引 (ExcelJS 使用 1-based 索引)
    // 列2=1月(B), 列3=2月(C), ..., 列10=9月(J), ..., 列13=12月(M)
    const allMonthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13] // B-M列
    const originalLastCol = 13 // 原始模板的最后一列（12月=M列）

    // --- 任务1: 隐藏多余月份列 ---
    // 从 "最后一个数据月份 + 1" 开始隐藏
    for (let i = lastDataMonth; i < allMonthCols.length; i++) {
      const colToHide = allMonthCols[i] // 列索引
      const column = sheet.getColumn(colToHide)
      column.hidden = true
      console.log(`[postProcess] 隐藏列 ${colToHide} (${i + 1}月)`)
    }

    // --- 任务2: 合并 A1:A2 ---
    const cellA1 = sheet.getCell('A1')
    const a1OriginalStyle = {
      font: cellA1.font ? { ...cellA1.font } : undefined,
      fill: cellA1.fill ? { ...cellA1.fill } : undefined,
      border: cellA1.border ? { ...cellA1.border } : undefined,
      alignment: cellA1.alignment ? { ...cellA1.alignment } : undefined,
      numFmt: cellA1.numFmt
    }

    // 先取消可能存在的合并
    try {
      sheet.unMergeCells('A1:A2')
    } catch {
      // 如果没有合并，忽略错误
    }

    // 合并 A1:A2
    sheet.mergeCells('A1:A2')
    console.log(`[postProcess] 合并单元格 A1:A2`)

    // 恢复样式并确保垂直居中和水平居中
    if (a1OriginalStyle.font) cellA1.font = a1OriginalStyle.font
    if (a1OriginalStyle.fill) cellA1.fill = a1OriginalStyle.fill
    if (a1OriginalStyle.border) cellA1.border = a1OriginalStyle.border
    if (a1OriginalStyle.numFmt) cellA1.numFmt = a1OriginalStyle.numFmt
    cellA1.alignment = {
      ...a1OriginalStyle.alignment,
      vertical: 'middle',
      horizontal: 'center'
    }

    // --- 任务3: 调整标题合并范围 (B1:XN1) ---
    const cellB1 = sheet.getCell('B1')
    const b1OriginalStyle = {
      font: cellB1.font ? { ...cellB1.font } : undefined,
      fill: cellB1.fill ? { ...cellB1.fill } : undefined,
      border: cellB1.border ? { ...cellB1.border } : undefined,
      alignment: cellB1.alignment ? { ...cellB1.alignment } : undefined,
      numFmt: cellB1.numFmt
    }

    // 1. 先取消原始的合并（B1:M1）
    const originalLastColLetter = String.fromCharCode(64 + originalLastCol) // 将13转换为'M'
    try {
      sheet.unMergeCells(`B1:${originalLastColLetter}1`)
    } catch {
      // 如果单元格未合并，忽略错误
      console.warn(`[postProcess] 原始标题范围可能未合并，继续处理...`)
    }

    // 2. 计算新的合并范围
    // lastDataMonth = 9 (9月), 对应的列索引是 allMonthCols[9-1] = 10 (J列)
    const lastDataColIndex = allMonthCols[lastDataMonth - 1]
    const lastDataColLetter = String.fromCharCode(64 + lastDataColIndex)
    const newTitleRange = `B1:${lastDataColLetter}1` // 例如 'B1:J1'

    // 3. 重新合并到新的范围
    sheet.mergeCells(newTitleRange)
    console.log(`[postProcess] 重新合并标题范围: ${newTitleRange}`)

    // 恢复样式并确保标题居中
    if (b1OriginalStyle.font) cellB1.font = b1OriginalStyle.font
    if (b1OriginalStyle.fill) cellB1.fill = b1OriginalStyle.fill
    if (b1OriginalStyle.border) cellB1.border = b1OriginalStyle.border
    if (b1OriginalStyle.numFmt) cellB1.numFmt = b1OriginalStyle.numFmt
    cellB1.alignment = {
      ...b1OriginalStyle.alignment,
      horizontal: 'center',
      vertical: 'middle'
    }

    // --- 任务4: 应用样式 ---
    console.log(`[postProcess] 开始应用样式...`)

    // 样式定义
    const orangeHeaderStyle = {
      fill: {
        type: 'pattern' as const,
        pattern: 'solid' as const,
        fgColor: { argb: 'FFED7D31' } // #ED7D31
      },
      font: {
        color: { argb: 'FFFFFFFF' }, // #FFFFFF
        bold: true
      },
      alignment: {
        horizontal: 'center' as const,
        vertical: 'middle' as const
      }
    }

    const lightOrangeDataStyle = {
      fill: {
        type: 'pattern' as const,
        pattern: 'solid' as const,
        fgColor: { argb: 'FFFCECE8' } // #FCECE8
      },
      font: {
        color: { argb: 'FF000000' } // #000000
      },
      alignment: {
        horizontal: 'center' as const,
        vertical: 'middle' as const
      }
    }

    // 应用第1行样式（标题行：行业、当月新增放款）
    const row1 = sheet.getRow(1)
    row1.eachCell({ includeEmpty: true }, (cell) => {
      cell.style = {
        fill: orangeHeaderStyle.fill,
        font: orangeHeaderStyle.font,
        alignment: orangeHeaderStyle.alignment,
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      }
    })
    console.log(`[postProcess] 第1行样式已应用`)

    // 应用第2行样式（月份标题：1月、2月...）
    const row2 = sheet.getRow(2)
    row2.eachCell({ includeEmpty: true }, (cell) => {
      cell.style = {
        fill: orangeHeaderStyle.fill,
        font: orangeHeaderStyle.font,
        alignment: orangeHeaderStyle.alignment,
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      }
    })
    console.log(`[postProcess] 第2行样式已应用`)

    // 应用第3行及以后的样式（数据行和合计行）
    // 遍历所有数据行，通过检查第一列内容判断是数据行还是合计行
    let totalRowNumber: number | null = null

    for (let rowNum = 3; rowNum <= sheet.rowCount; rowNum++) {
      const row = sheet.getRow(rowNum)
      const firstCell = row.getCell(1) // A列
      const firstCellValue = String(firstCell.value || '').trim()

      // 判断是否是合计行
      if (firstCellValue === '合计' || firstCellValue === '总计') {
        totalRowNumber = rowNum
        // 应用橙色合计行样式
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.style = {
            fill: orangeHeaderStyle.fill,
            font: orangeHeaderStyle.font,
            alignment: orangeHeaderStyle.alignment,
            border: {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            }
          }
        })
        console.log(`[postProcess] 第${rowNum}行样式已应用（合计行）`)
      } else {
        // 应用浅橙色数据行样式
        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.style = {
            fill: lightOrangeDataStyle.fill,
            font: lightOrangeDataStyle.font,
            alignment: lightOrangeDataStyle.alignment,
            border: {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            }
          }
        })
      }
    }

    if (totalRowNumber) {
      console.log(`[postProcess] 数据行样式已应用（第3-${totalRowNumber - 1}行）`)
    } else {
      console.log(`[postProcess] 数据行样式已应用（第3-${sheet.rowCount}行），未找到合计行`)
    }

    // --- 保存文件 ---
    await workbook.xlsx.writeFile(outputPath)
    console.log(`[postProcess] 处理完成! 已保存至: ${outputPath}`)
  } catch (error) {
    console.error('[postProcess] 处理Excel时出错:', error)
    throw error
  }
}
