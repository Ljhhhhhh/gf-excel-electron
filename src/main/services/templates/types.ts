/**
 * 模板系统核心类型定义
 */

import type { Workbook } from 'exceljs'
import type { Warning } from '../errors'

// ========== 模板标识与元信息 ==========

/**
 * 模板唯一标识符（简单字符串，不考虑版本）
 */
export type TemplateId = string

/**
 * 模板元信息
 */
export interface TemplateMeta {
  /** 模板唯一 ID */
  id: TemplateId
  /** 模板显示名称 */
  name: string
  /** 模板文件名（如 'month1carbone.xlsx'） */
  filename: string
  /** 模板文件扩展名（当前仅 xlsx） */
  ext: 'xlsx'
  /** 支持的源文件扩展名 */
  supportedSourceExts: string[]
  /** 描述 */
  description?: string
}

// ========== 模板解析与构建 ==========

/**
 * 模板解析选项（由每个模板自定义）
 */
export type ParseOptions = Record<string, unknown>

/**
 * 解析后的数据（由每个模板自定义结构）
 */
export type ParsedData = unknown

/**
 * 源文件元信息
 */
export interface SourceMeta {
  /** 源文件路径 */
  path: string
  /** 文件大小（字节） */
  size: number
  /** 工作表名称列表 */
  sheets: string[]
  /** 其他元信息 */
  [key: string]: unknown
}

/**
 * 模板解析器函数签名
 * @param workbook ExcelJS Workbook 实例
 * @param parseOptions 模板自定义的解析选项
 * @returns 解析后的结构化数据
 */
export type TemplateParser = (workbook: Workbook, parseOptions?: ParseOptions) => ParsedData

/**
 * 报表数据构建器函数签名
 * @param parsedData 解析后的数据
 * @param userInput 用户输入的额外参数（如查询年月、过滤条件等），可选
 * @returns Carbone 渲染所需的 JSON 数据模型
 */
export type TemplateReportBuilder = (parsedData: ParsedData, userInput?: unknown) => unknown

// ========== Carbone 渲染选项 ==========

/**
 * Carbone 渲染选项（可选覆写项）
 */
export interface CarboneRenderOptions {
  /** 转换格式（当前不使用，保留以便扩展） */
  convertTo?: string
  /** 时区（默认 Asia/Shanghai） */
  timezone?: string
  /** 语言/区域（默认 zh-cn） */
  lang?: string
  /** 报表名称（动态文件名） */
  reportName?: string
  /** 枚举映射 */
  enum?: Record<string, unknown>
  /** 翻译映射 */
  translations?: Record<string, Record<string, string>>
  /** 补充数据（可通过 {c.} 访问） */
  complement?: unknown
  /** 其他 Carbone 选项 */
  [key: string]: unknown
}

// ========== 模板完整定义 ==========

/**
 * 模板完整定义
 */
export interface TemplateDefinition {
  /** 元信息 */
  meta: TemplateMeta
  /** 解析器函数 */
  parser: TemplateParser
  /** 报表数据构建器 */
  builder: TemplateReportBuilder
  /** Carbone 默认选项（可被 renderOptions 覆写） */
  carboneOptions?: CarboneRenderOptions
}

// ========== 服务层接口 ==========

/**
 * excelToData 输入参数
 */
export interface ExcelToDataInput {
  /** 源文件路径 */
  sourcePath: string
  /** 模板 ID */
  templateId: TemplateId
  /** 模板解析选项 */
  parseOptions?: ParseOptions
}

/**
 * excelToData 返回结果
 */
export interface ExcelToDataResult {
  /** 模板 ID */
  templateId: TemplateId
  /** 解析后的数据 */
  data: ParsedData
  /** 警告信息 */
  warnings: Warning[]
  /** 源文件元信息 */
  sourceMeta: SourceMeta
}

/**
 * dataToReport 输入参数
 */
export interface DataToReportInput {
  /** 模板 ID */
  templateId: TemplateId
  /** 解析后的数据 */
  parsedData: ParsedData
  /** 输出目录（用户选择） */
  outputDir: string
  /** 报表名称（用户指定，若无则使用兜底策略） */
  reportName?: string
  /** Carbone 渲染选项（可选覆写） */
  renderOptions?: CarboneRenderOptions
}

/**
 * dataToReport 返回结果
 */
export interface DataToReportResult {
  /** 输出文件路径 */
  outputPath: string
  /** 文件大小（字节） */
  size: number
  /** 生成时间 */
  generatedAt: Date
}
