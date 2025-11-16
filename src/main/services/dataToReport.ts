/**
 * 数据 → 报表服务
 * 负责基于解析后的数据，使用 Carbone 渲染生成报表文件
 */

import carbone from 'carbone'
import fs from 'node:fs'
import path from 'node:path'
import { promisify } from 'node:util'
import * as XLSX from 'xlsx-js-style'
import type { DataToReportInput, DataToReportResult } from './templates/types'
import { getTemplate } from './templates/registry'
import { ReportRenderError, OutputWriteError, OutputDirNotSelectedError } from './errors'
import { getTemplatePath, ensureOutputDir } from './utils/filePaths'
import { generateReportName, sanitizeFilename } from './utils/naming'
import { getFileSize, deleteFileIfExists } from './utils/fileOps'
import { createLogger } from './logger'

// 将 carbone.render 转为 Promise 形式
const renderAsync = promisify(carbone.render)
const log = createLogger('dataToReport')

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

  log.info('开始生成报表', { templateId, outputDir })

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
  log.info('模板已加载', {
    templateId,
    templateName: template.meta.name,
    engine: template.engine
  })

  // 4. 确定输出文件名和路径
  const templateDisplayName = template.meta.name || template.meta.id || templateId
  const finalReportName = reportName
    ? sanitizeFilename(reportName)
    : generateReportName(templateDisplayName, template.meta.ext)
  const outputPath = path.join(outputDir, finalReportName)
  log.info('输出路径已确定', { outputPath })

  // 5. 根据引擎类型分支处理
  if (template.engine === 'exceljs') {
    // ========== ExcelJS 模式 ==========
    log.info('使用 ExcelJS 引擎渲染', { templateId })

    if (!template.excelRenderer) {
      throw new ReportRenderError(
        templateId,
        new Error(`模板 ${templateId} 声明为 exceljs 引擎但未提供 excelRenderer`)
      )
    }

    try {
      // 删除已存在的文件
      deleteFileIfExists(outputPath)

      // 调用 ExcelJS 渲染器，直接生成文件
      await template.excelRenderer(parsedData, userInput, outputPath)
      log.info('ExcelJS 渲染完成', { templateId, outputPath })
    } catch (error) {
      throw new ReportRenderError(templateId, error)
    }
  } else {
    // ========== Carbone 模式 ==========
    log.info('使用 Carbone 引擎渲染', { templateId })

    if (!template.builder) {
      throw new ReportRenderError(
        templateId,
        new Error(`模板 ${templateId} 声明为 carbone 引擎但未提供 builder`)
      )
    }

    // 4.1. 调用模板的 buildReportData 构建 Carbone 数据
    let reportData
    try {
      reportData = template.builder(parsedData, userInput)
      log.info('报表数据已构建', { templateId })
    } catch (error) {
      throw new ReportRenderError(templateId, error)
    }

    // 4.2. 合并 Carbone 选项
    const carboneOptions = {
      ...DEFAULT_CARBONE_OPTIONS,
      ...template.carboneOptions,
      ...renderOptions
    }

    // 4.3. 获取模板文件路径
    if (!template.meta.filename) {
      throw new ReportRenderError(
        templateId,
        new Error(`Carbone 模板 ${templateId} 缺少 filename 字段`)
      )
    }
    const templatePath = getTemplatePath(template.meta.filename)
    log.info('Carbone 模板路径已解析', { templatePath })

    // 4.4. 使用 Carbone 渲染
    let resultBuffer: Buffer
    try {
      const result = await renderAsync(templatePath, reportData, carboneOptions)
      resultBuffer = result as Buffer
      log.info('Carbone 渲染完成', {
        templateId,
        bufferSize: resultBuffer.length
      })
    } catch (error) {
      throw new ReportRenderError(templateId, error)
    }

    // 4.5. 写入文件
    try {
      deleteFileIfExists(outputPath)
      fs.writeFileSync(outputPath, resultBuffer)
      log.info('报表文件写入完成', { outputPath })
    } catch (error) {
      throw new OutputWriteError(outputPath, error)
    }

    // 4.6. 如果有 postProcess 钩子，先用 xlsx-js-style 规范化 Carbone 生成的文件
    if (template.postProcess) {
      log.info('检测到 postProcess 钩子，准备规范化', { templateId })
      try {
        const tempPath = `${outputPath}.temp.xlsx`

        // 使用 xlsx-js-style 读取 Carbone 生成的文件（对格式更宽容）
        const workbook = XLSX.readFile(outputPath)
        log.info('xlsx-js-style 读取成功', { outputPath })

        // 重新写入文件（这会规范化文件结构，使 ExcelJS 能读取）
        XLSX.writeFile(workbook, tempPath)
        log.info('规范化写入临时文件完成', { tempPath })

        // 删除原 Carbone 文件
        deleteFileIfExists(outputPath)

        // 将临时文件重命名为原文件名
        fs.renameSync(tempPath, outputPath)
        log.info('规范化完成并替换原文件', { outputPath })
      } catch (normalizeError) {
        log.warn('文件规范化失败，尝试继续执行 postProcess', { error: normalizeError })
      }
    }
  }

  // 6. 调用后处理钩子（两种引擎共用）
  if (template.postProcess) {
    log.info('开始执行 postProcess', { templateId, outputPath })
    try {
      await template.postProcess(outputPath, userInput)
      log.info('postProcess 执行完成', { templateId, outputPath })
    } catch (error) {
      log.error('postProcess 执行失败', { templateId, outputPath, error })
      throw new ReportRenderError(templateId, error)
    }
  }

  // 11. 获取文件大小并返回结果
  const size = getFileSize(outputPath)
  const generatedAt = new Date()

  return {
    outputPath,
    size,
    generatedAt
  }
}
