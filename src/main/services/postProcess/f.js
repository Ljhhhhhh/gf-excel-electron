const ExcelJS = require('exceljs')
const path = require('path')

/**
 * 处理Excel报告
 * @param {string} inputPath - 模板文件路径
 * @param {string} outputPath - 输出文件路径
 * @param {number} lastDataMonth - 最后一个有数据的月份 (例如, 9月就传入 9)
 */
async function processExcelReport(inputPath, outputPath, lastDataMonth) {
  try {
    console.log(inputPath, outputPath, lastDataMonth)

    // 检查文件是否被占用
    const fs = require('fs')
    try {
      const stats = fs.statSync(inputPath)
      console.log(`文件大小: ${stats.size} 字节`)
    } catch (err) {
      throw new Error(`无法访问文件: ${err.message}`)
    }

    // 1. 创建工作簿实例并加载文件
    const workbook = new ExcelJS.Workbook()

    console.log('正在读取文件...')
    try {
      await workbook.xlsx.readFile(inputPath)
      console.log('文件读取成功')
    } catch (readError) {
      console.error('读取文件时出错:', readError.message)
      throw new Error(
        `Excel 文件读取失败: ${readError.message}\n建议: 1) 确保文件未被 Excel 打开; 2) 尝试另存为新文件后再处理`
      )
    }

    // 假设您的工作表名为 'Sheet1'，请根据实际情况修改
    const sheet = workbook.getWorksheet('Sheet1')

    if (!sheet) {
      // 列出所有可用的工作表
      const sheetNames = workbook.worksheets.map((ws) => ws.name).join(', ')
      throw new Error(`找不到名为 "Sheet1" 的工作表。可用工作表: ${sheetNames}`)
    }

    console.log(`工作表 "${sheet.name}" 加载成功，行数: ${sheet.rowCount}`)

    // 定义月份对应的列索引 (ExcelJS 使用 1-based 索引)
    // 列2=1月(B), 列3=2月(C), ..., 列10=9月(J), ..., 列13=12月(M)
    const allMonthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13] // B-M列
    const originalLastCol = 13 // 原始模板的最后一列（12月=M列）

    // --- 任务1: 移除（隐藏）多余月份 ---
    // 从 "最后一个数据月份 + 1" 开始隐藏
    for (let i = lastDataMonth; i < allMonthCols.length; i++) {
      const colToHide = allMonthCols[i] // 列索引
      const column = sheet.getColumn(colToHide)
      column.hidden = true
    }

    // --- 任务2: 合并 A1、A2 ---
    // 保存 A1 的原有样式
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
    } catch (e) {
      // 如果没有合并，忽略错误
    }

    // 合并 A1:A2
    sheet.mergeCells('A1:A2')

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

    // --- 任务3: 合并标题 (调整B1的合并范围) ---
    // 保存 B1 的原有样式
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
    } catch (e) {
      // 如果单元格未合并，忽略错误
      console.warn('Original title range might not be merged, proceeding...')
    }

    // 2. 计算新的合并范围
    // lastDataMonth = 9 (9月), 对应的列索引是 allMonthCols[9-1] = 10 (J列)
    const lastDataColIndex = allMonthCols[lastDataMonth - 1]
    const lastDataColLetter = String.fromCharCode(64 + lastDataColIndex)
    const newTitleRange = `B1:${lastDataColLetter}1` // 例如 'B1:J1'

    // 3. 重新合并到新的范围
    sheet.mergeCells(newTitleRange)

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

    // --- 保存文件 ---
    await workbook.xlsx.writeFile(outputPath)
    console.log(`报告处理完成! 已保存至: ${outputPath}`)
  } catch (error) {
    console.error('处理Excel时出错:', error)
  }
}

// --- 如何使用 ---

// 假设您的原文件叫 template.xlsx
const inputFile = path.resolve(__dirname, '2.xlsx')
console.log(inputFile, 'inputFile')
// 处理后的文件叫 output.xlsx
const outputFile = path.resolve(__dirname, 'output.xlsx')
console.log(outputFile)

// 根据您的图片，数据到9月 (Column J)
// 所以我们传入 9
processExcelReport(inputFile, outputFile, 9)

// 如果您下次只想生成 1-8月 (Column I) 的报告
// 您只需要调用: processExcelReport(inputFile, outputFile, 8);
