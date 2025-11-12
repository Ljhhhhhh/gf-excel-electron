/**
 * 数据 → 报表服务
 * 负责基于解析后的数据，使用 Carbone 渲染生成报表文件
 */

import carbone from 'carbone'
import fs from 'node:fs'
import path from 'node:path'
import { promisify } from 'node:util'
import type { DataToReportInput, DataToReportResult } from './templates/types'
import { getTemplate } from './templates/registry'
import { ReportRenderError, OutputWriteError, OutputDirNotSelectedError } from './errors'
import { getTemplatePath, ensureOutputDir } from './utils/filePaths'
import { generateReportName, sanitizeFilename } from './utils/naming'
import { getFileSize, deleteFileIfExists } from './utils/fileOps'

// 将 carbone.render 转为 Promise 形式
const renderAsync = promisify(carbone.render)

/**
 * 默认 Carbone 选项
 */
const DEFAULT_CARBONE_OPTIONS = {
  lang: 'zh-cn',
  timezone: 'Asia/Shanghai'
}

/**
 * 数据 → 报表转换
 * @param input 输入参数
 * @returns 生成结果
 */
export async function dataToReport(input: DataToReportInput): Promise<DataToReportResult> {
  const { templateId, parsedData, outputDir, reportName, renderOptions, userInput } = input

  console.log(`[dataToReport] 开始生成报表: ${templateId}`)

  // 1. 校验输出目录
  if (!outputDir) {
    throw new OutputDirNotSelectedError()
  }

  // 2. 确保输出目录存在
  try {
    ensureOutputDir(outputDir)
  } catch (error) {
    throw new OutputWriteError(outputDir, error)
  }

  // 3. 获取模板定义
  const template = getTemplate(templateId)
  console.log(`[dataToReport] 模板已加载: ${template.meta.name}`)

  // 4. 调用模板的 buildReportData 构建 Carbone 数据
  let reportData
  try {
    reportData = template.builder(parsedData, userInput)
    console.log(`[dataToReport] 报表数据已构建`)
    console.log('[dataToReport] 传递给 Carbone 的数据:', JSON.stringify(reportData, null, 2))
  } catch (error) {
    throw new ReportRenderError(templateId, error)
  }

  // 5. 合并 Carbone 选项
  const carboneOptions = {
    ...DEFAULT_CARBONE_OPTIONS,
    ...template.carboneOptions,
    ...renderOptions
  }

  // 6. 获取模板文件路径
  const templatePath = getTemplatePath(template.meta.filename)
  console.log(`[dataToReport] 模板文件路径: ${templatePath}`)

  // 7. 使用 Carbone 渲染
  let resultBuffer: Buffer
  try {
    const result = await renderAsync(templatePath, reportData, carboneOptions)
    resultBuffer = result as Buffer
    console.log(`[dataToReport] Carbone 渲染完成，大小: ${resultBuffer.length} 字节`)
  } catch (error) {
    throw new ReportRenderError(templateId, error)
  }

  // 8. 确定输出文件名
  const finalReportName = reportName
    ? sanitizeFilename(reportName)
    : generateReportName(templateId, template.meta.ext)

  const outputPath = path.join(outputDir, finalReportName)
  console.log(`[dataToReport] 输出路径: ${outputPath}`)

  // 9. 写入文件
  try {
    // 如果文件已存在，先删除
    deleteFileIfExists(outputPath)
    fs.writeFileSync(outputPath, resultBuffer)
    console.log(`[dataToReport] 报表已写入`)
  } catch (error) {
    throw new OutputWriteError(outputPath, error)
  }

  // 10. 获取文件大小并返回结果
  const size = getFileSize(outputPath)
  const generatedAt = new Date()

  return {
    outputPath,
    size,
    generatedAt
  }
}
