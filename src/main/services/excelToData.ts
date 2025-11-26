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
  ExtraSourceContext,
  TemplateSourceRequirement
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
import { createLogger } from './logger'

const log = createLogger('excelToData')

/**
 * 文件大小限制（MB）
 */
const FILE_SIZE_LIMIT_MB = 100

/**
 * 流式工作簿选项
 */
const STREAM_WORKBOOK_OPTIONS: ExcelJS.stream.xlsx.WorkbookStreamReaderOptions = {
  sharedStrings: 'cache',
  hyperlinks: 'ignore',
  styles: 'ignore',
  worksheets: 'emit'
}

interface ResolveExtraSourceContextInput {
  requirement: TemplateSourceRequirement
  providedPath: string
  templateExts: string[]
  supportsStreaming: boolean
}

async function resolveExtraSourceContext(
  input: ResolveExtraSourceContextInput
): Promise<ExtraSourceContext> {
  const { requirement, providedPath, templateExts, supportsStreaming } = input
  if (!fs.existsSync(providedPath)) {
    throw new UnsupportedFileError(providedPath, '文件不存在')
  }

  const ext = getFileExtension(providedPath)
  const allowedExts = requirement.supportedExts ?? templateExts
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

  const desiredStrategy = requirement.loadStrategy ?? 'auto'
  let shouldLoadWorkbook = desiredStrategy === 'workbook'
  let shouldProvideStream = desiredStrategy === 'stream'

  if (desiredStrategy === 'auto') {
    shouldProvideStream = supportsStreaming
    shouldLoadWorkbook = !supportsStreaming
  } else if (desiredStrategy === 'stream' && !supportsStreaming) {
    log.warn('额外数据源请求流式解析但模板未启用流式解析，回退为 workbook 模式', {
      requirementId: requirement.id
    })
    shouldProvideStream = false
    shouldLoadWorkbook = true
  }

  if (!shouldProvideStream && !shouldLoadWorkbook) {
    shouldLoadWorkbook = true
  }

  let workbook: ExcelJS.Workbook | undefined
  let sheets: string[] = []
  if (shouldLoadWorkbook) {
    workbook = new ExcelJS.Workbook()
    try {
      // 使用流式加载（与主数据源保持一致）
      const stream = fs.createReadStream(providedPath)
      await workbook.xlsx.read(stream)
      sheets = workbook.worksheets.map((ws) => ws.name)
    } catch (error) {
      // 增强错误日志：记录完整的原始错误信息以便诊断
      log.error('额外数据源 Excel 读取失败', { data: JSON.stringify(error, null, 2) })
      throw new ExcelParseError(providedPath, error)
    }
    log.info('额外数据源已加载 Workbook（流式）', {
      requirementId: requirement.id,
      label: requirement.label,
      sheets
    })
  }

  const createReader = shouldProvideStream
    ? () => new ExcelJS.stream.xlsx.WorkbookReader(providedPath, STREAM_WORKBOOK_OPTIONS)
    : undefined

  const loadMode =
    shouldLoadWorkbook && shouldProvideStream
      ? 'hybrid'
      : shouldLoadWorkbook
        ? 'workbook'
        : 'stream'

  return {
    path: providedPath,
    size,
    sheets,
    workbook,
    createReader,
    loadMode
  }
}

/**
 * Excel → 数据转换
 * @param input 输入参数
 * @returns 解析结果
 */
export async function excelToData(input: ExcelToDataInput): Promise<ExcelToDataResult> {
  const { sourcePath, templateId, parseOptions, extraSources } = input
  const warnings: Warning[] = []

  log.info('开始解析 Excel', { sourcePath, templateId })

  // 1. 获取模板定义
  const template = getTemplate(templateId)
  log.info('模板已加载', { templateId, templateName: template.meta.name })

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
  // 6. 检查模板是否支持流式解析
  const supportsStreaming =
    'streamParser' in template && typeof template.streamParser === 'function'

  // 7. 处理额外数据源（根据流式模式和加载策略灵活加载）
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

    resolvedExtraSources[requirement.id] = await resolveExtraSourceContext({
      requirement,
      providedPath,
      templateExts: template.meta.supportedSourceExts,
      supportsStreaming
    })
  }

  let workbook: ExcelJS.Workbook
  let parsedData: unknown
  let sourceMeta: SourceMeta

  const parserOptions: ParseOptions = {
    ...((parseOptions as ParseOptions) || {}),
    extraSources: Object.keys(resolvedExtraSources).length ? resolvedExtraSources : undefined
  }

  if (supportsStreaming) {
    // 使用流式解析(避免全量加载)
    log.info('使用流式解析模式', { templateId })
    try {
      parsedData = await template.streamParser!(sourcePath, parserOptions)
      log.info('流式解析完成', { templateId })
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
    // 使用完整加载模式（流式读取）
    log.info('使用完整加载模式（流式）', { templateId })
    try {
      workbook = new ExcelJS.Workbook()
      // 使用流式加载以优化内存使用
      const stream = fs.createReadStream(sourcePath)
      await workbook.xlsx.read(stream)
      log.info('Workbook 已加载（流式）', {
        templateId,
        sheetCount: workbook.worksheets.length,
        sheets: workbook.worksheets.map((ws) => ws.name)
      })
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
      log.info('解析完成', {
        templateId,
        warnings: warnings.length
      })
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

  log.info('excelToData 完成', {
    templateId,
    sourcePath,
    warningCount: warnings.length,
    sheetCount: sourceMeta.sheets.length
  })

  return {
    templateId,
    data: parsedData,
    warnings,
    sourceMeta
  }
}
