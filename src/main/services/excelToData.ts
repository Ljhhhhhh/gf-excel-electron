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
  SourceMeta
} from './templates/types'
import type { Warning } from './errors'
import { getTemplate } from './templates/registry'
import { UnsupportedFileError, ExcelFileTooLargeError, ExcelParseError } from './errors'
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
  const { sourcePath, templateId, parseOptions } = input
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

  // 5. 使用 ExcelJS 读取 workbook
  let workbook: ExcelJS.Workbook
  try {
    workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(sourcePath)
    console.log(`[excelToData] Workbook 已加载，共 ${workbook.worksheets.length} 个 sheet`)
  } catch (error) {
    throw new ExcelParseError(sourcePath, error)
  }

  // 6. 提取源文件元信息
  const sourceMeta: SourceMeta = {
    path: sourcePath,
    size: fileSize,
    sheets: workbook.worksheets.map((ws) => ws.name)
  }

  // 7. 调用模板解析器
  let parsedData
  try {
    parsedData = template.parser(workbook, parseOptions as ParseOptions)
    console.log(`[excelToData] 解析完成`)
  } catch (error) {
    throw new ExcelParseError(sourcePath, error)
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
