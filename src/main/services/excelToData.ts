/**
 * Excel → 数据服务
 * 负责读取源 Excel 文件，使用模板解析器提取结构化数据
 */

import ExcelJS from 'exceljs'
import fs from 'node:fs'
import type {
  ExcelToDataInput,
  ExcelToDataResult,
  ParseOptions,
  SourceMeta,
  ExtraSourceContext
} from './templates/types'
import type { Warning } from './errors'
import { getTemplate } from './templates/registry'
import {
  UnsupportedFileError,
  ExcelFileTooLargeError,
  ExcelParseError,
  MissingSourceError
} from './errors'
import { getFileSize, getFileExtension } from './utils/fileOps'

/**
 * 文件大小限制（MB）
 */
const FILE_SIZE_LIMIT_MB = 100

/**
 * Excel → 数据转换
 * @param input 输入参数
 * @returns 解析结果
 */
export async function excelToData(input: ExcelToDataInput): Promise<ExcelToDataResult> {
  const { sourcePath, templateId, parseOptions, extraSources } = input
  const warnings: Warning[] = []

  console.log(`[excelToData] 开始处理: ${sourcePath} with template: ${templateId}`)

  // 1. 获取模板定义
  const template = getTemplate(templateId)
  console.log(`[excelToData] 模板已加载: ${template.meta.name}`)

  // 2. 校验文件是否存在
  if (!fs.existsSync(sourcePath)) {
    throw new UnsupportedFileError(sourcePath, '文件不存在')
  }

  // 3. 校验文件扩展名
  const ext = getFileExtension(sourcePath)
  if (!template.meta.supportedSourceExts.includes(ext)) {
    throw new UnsupportedFileError(
      sourcePath,
      `不支持的扩展名: ${ext}，模板仅支持: ${template.meta.supportedSourceExts.join(', ')}`
    )
  }

  // 4. 校验文件大小
  const fileSize = getFileSize(sourcePath)
  if (fileSize > FILE_SIZE_LIMIT_MB * 1024 * 1024) {
    throw new ExcelFileTooLargeError(sourcePath, fileSize, FILE_SIZE_LIMIT_MB)
  }

  // 5. 处理额外数据源
  // 如果模板声明了额外数据源，这里逐个解析与校验，并将 workbook 句柄注入 parseOptions
  const extraSourceConfigs = template.meta.extraSources ?? []
  const resolvedExtraSources: Record<string, ExtraSourceContext> = {}
  for (const requirement of extraSourceConfigs) {
    const providedPath = extraSources?.[requirement.id]
    const isRequired = requirement.required !== false
    if (!providedPath) {
      if (isRequired) {
        throw new MissingSourceError(requirement.id, requirement.label)
      }
      continue
    }

    if (!fs.existsSync(providedPath)) {
      throw new UnsupportedFileError(providedPath, '文件不存在')
    }

    const ext = getFileExtension(providedPath)
    const allowedExts = requirement.supportedExts ?? template.meta.supportedSourceExts
    if (!allowedExts.includes(ext)) {
      throw new UnsupportedFileError(
        providedPath,
        `不支持的扩展名: ${ext}，仅支持: ${allowedExts.join(', ')}`
      )
    }

    const size = getFileSize(providedPath)
    if (size > FILE_SIZE_LIMIT_MB * 1024 * 1024) {
      throw new ExcelFileTooLargeError(providedPath, size, FILE_SIZE_LIMIT_MB)
    }

    const workbook = new ExcelJS.Workbook()
    try {
      await workbook.xlsx.readFile(providedPath)
      console.log(`[excelToData] 额外数据源已加载: ${requirement.label || requirement.id}`)
    } catch (error) {
      throw new ExcelParseError(providedPath, error)
    }

    resolvedExtraSources[requirement.id] = {
      path: providedPath,
      workbook,
      size,
      sheets: workbook.worksheets.map((ws) => ws.name)
    }
  }

  // 6. 检查模板是否支持流式解析
  const supportsStreaming =
    'streamParser' in template && typeof template.streamParser === 'function'

  let workbook: ExcelJS.Workbook
  let parsedData: unknown
  let sourceMeta: SourceMeta

  const parserOptions: ParseOptions = {
    ...((parseOptions as ParseOptions) || {}),
    extraSources: Object.keys(resolvedExtraSources).length ? resolvedExtraSources : undefined
  }

  if (supportsStreaming) {
    // 使用流式解析(避免全量加载)
    console.log(`[excelToData] 使用流式解析模式`)
    try {
      parsedData = await template.streamParser!(sourcePath, parserOptions)
      console.log(`[excelToData] 流式解析完成`)
    } catch (error) {
      throw new ExcelParseError(sourcePath, error)
    }

    // 流式模式下无法获取完整sheet列表,使用空数组
    sourceMeta = {
      path: sourcePath,
      size: fileSize,
      sheets: [] // 流式模式下不提供sheet列表
    }
  } else {
    // 使用传统完整加载模式
    console.log(`[excelToData] 使用完整加载模式`)
    try {
      workbook = new ExcelJS.Workbook()
      await workbook.xlsx.readFile(sourcePath)
      console.log(`[excelToData] Workbook 已加载,共 ${workbook.worksheets.length} 个 sheet`)
    } catch (error) {
      throw new ExcelParseError(sourcePath, error)
    }

    // 6. 提取源文件元信息
    sourceMeta = {
      path: sourcePath,
      size: fileSize,
      sheets: workbook.worksheets.map((ws) => ws.name)
    }

    // 7. 调用模板解析器
    try {
      parsedData = await template.parser(workbook, parserOptions)
      console.log(`[excelToData] 解析完成`)
    } catch (error) {
      throw new ExcelParseError(sourcePath, error)
    }
  }

  // 8. 检查解析结果是否有效
  if (!parsedData) {
    warnings.push({
      code: 'PARSE_NO_DATA',
      message: '解析结果为空',
      level: 'warn'
    })
  }

  return {
    templateId,
    data: parsedData,
    warnings,
    sourceMeta
  }
}
