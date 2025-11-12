/**
 * 完整测试脚本：生成 month1carbone 报表
 * 包含数据解析 + Carbone 渲染 + 文件输出
 */

import ExcelJS from 'exceljs'
import carbone from 'carbone'
import fs from 'node:fs'
import path from 'node:path'
import { promisify } from 'node:util'
import {
  parseWorkbook,
  buildReportData,
  type Month1CarboneInput
} from '../src/main/services/templates/month1carbone'

// 将 carbone.render 转为 Promise
const renderAsync = promisify(carbone.render)

interface TestOptions {
  /** 数据源文件路径 */
  sourcePath: string
  /** 模板文件路径 */
  templatePath: string
  /** 输出目录 */
  outputDir: string
  /** 查询年份 */
  queryYear: number
  /** 查询月份 */
  queryMonth: number
  /** 输出文件名（可选） */
  outputFileName?: string
}

async function generateReport(options: TestOptions): Promise<void> {
  const { sourcePath, templatePath, outputDir, queryYear, queryMonth, outputFileName } = options

  console.log('=== 生成 month1carbone 报表 ===\n')
  console.log(`📋 查询条件: ${queryYear}年${queryMonth}月`)
  console.log(`📂 数据源: ${sourcePath}`)
  console.log(`📄 模板: ${templatePath}`)
  console.log(`📁 输出目录: ${outputDir}\n`)

  // ========== 步骤 1: 读取数据源文件 ==========
  console.log('⏳ [1/5] 读取数据源文件...')
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(sourcePath)
  console.log('✅ 数据源文件读取完成\n')

  // ========== 步骤 2: 解析数据 ==========
  console.log('⏳ [2/5] 解析数据...')
  const parsedData = parseWorkbook(workbook)
  console.log(`✅ 解析完成 - 总行数: ${parsedData.summary?.totalRows}\n`)

  // ========== 步骤 3: 构建报表数据 ==========
  console.log('⏳ [3/5] 构建报表数据...')
  const userInput: Month1CarboneInput = {
    queryYear,
    queryMonth
  }

  const reportData = buildReportData(parsedData, userInput)
  console.log('✅ 报表数据构建完成')
  console.log('数据预览:', JSON.stringify(reportData as Record<string, unknown>, null, 2))
  console.log()

  // ========== 步骤 4: Carbone 渲染 ==========
  console.log('⏳ [4/5] 使用 Carbone 渲染报表...')
  const carboneOptions = {
    lang: 'zh-cn',
    timezone: 'Asia/Shanghai'
  }

  let resultBuffer: Buffer
  try {
    console.log(reportData, 'reportData')
    const result = await renderAsync(templatePath, reportData as object, carboneOptions)
    resultBuffer = result as Buffer
    console.log(`✅ 渲染完成 - 文件大小: ${(resultBuffer.length / 1024).toFixed(2)} KB\n`)
  } catch (error) {
    console.error('❌ Carbone 渲染失败:', error)
    throw error
  }

  // ========== 步骤 5: 写入文件 ==========
  console.log('⏳ [5/5] 写入输出文件...')

  // 确保输出目录存在
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  // 生成输出文件名
  const fileName =
    outputFileName ||
    `month1carbone-${queryYear}年${queryMonth}月-${formatDateTime(new Date())}.xlsx`

  const outputPath = path.join(outputDir, fileName)

  // 写入文件
  fs.writeFileSync(outputPath, resultBuffer)
  console.log(`✅ 报表已生成\n`)

  // ========== 完成 ==========
  console.log('🎉 ===== 报表生成成功 ===== 🎉\n')
  console.log(`📄 输出文件: ${outputPath}`)
  console.log(`📊 文件大小: ${(fs.statSync(outputPath).size / 1024).toFixed(2)} KB`)
  console.log()
  console.log('💡 使用以下命令打开输出文件夹:')
  console.log(`   explorer "${outputDir}"`)
  console.log()
}

/**
 * 格式化日期时间为文件名友好格式
 */
function formatDateTime(date: Date): string {
  const pad = (n: number): string => n.toString().padStart(2, '0')
  const year = date.getFullYear()
  const month = pad(date.getMonth() + 1)
  const day = pad(date.getDate())
  const hour = pad(date.getHours())
  const minute = pad(date.getMinutes())
  const second = pad(date.getSeconds())
  return `${year}${month}${day}-${hour}${minute}${second}`
}

// ========== 主程序 ==========
async function main(): Promise<void> {
  const projectRoot = process.cwd()

  const options: TestOptions = {
    sourcePath: path.join(projectRoot, 'public/demo/放款明细.xlsx'),
    templatePath: path.join(projectRoot, 'public/reportTemplates/month1carbone.xlsx'),
    outputDir: path.join(projectRoot, 'output'),
    queryYear: 2025,
    queryMonth: 10,
    outputFileName: `month1carbone-${Date.now()}.xlsx`
  }

  try {
    await generateReport(options)
  } catch (error) {
    console.error('\n❌ 生成失败:', error)
    if (error instanceof Error) {
      console.error('错误信息:', error.message)
      console.error('错误堆栈:', error.stack)
    }
    process.exit(1)
  }
}

main()
