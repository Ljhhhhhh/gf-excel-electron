/**
 * 模板系统核心类型定义
 */

import type { Workbook } from 'exceljs'
import type { Warning } from '../errors'

// ========== formCreate 规则类型 ==========

/**
 * formCreate Rule 类型（简化版，实际使用时由 @form-create/element-ui 提供完整类型）
 * 这里定义基础结构以支持序列化和类型检查
 */
export interface FormCreateRule {
  /** 组件类型 */
  type: string
  /** 字段名 */
  field: string
  /** 字段标题 */
  title: string
  /** 默认值 */
  value?: any
  /** 组件属性 */
  props?: Record<string, any>
  /** 验证规则 */
  validate?: Array<{
    required?: boolean
    message?: string
    trigger?: string | string[]
    type?: string
    min?: number
    max?: number
    pattern?: string
    validator?: (rule: any, value: any, callback: (error?: Error) => void) => void
    [key: string]: any
  }>
  /** 选项列表（用于 select、radio、checkbox 等） */
  options?: Array<{ label: string; value: any; [key: string]: any }>
  /** 是否显示 */
  display?: boolean
  /** 联动字段 */
  link?: string[]
  /** 更新回调 */
  update?: (value: any, rule: FormCreateRule, fApi: any) => void
  /** 子组件 */
  children?: FormCreateRule[]
  /** 其他配置 */
  [key: string]: any
}

/**
 * 模板输入参数规则（基于 formCreate）
 */
export interface TemplateInputRule {
  /** formCreate 规则数组 */
  rules: FormCreateRule[]
  /** 表单配置（可选） */
  options?: {
    /** 表单标签宽度 */
    labelWidth?: string | number
    /** 表单标签位置 */
    labelPosition?: 'left' | 'right' | 'top'
    /** 是否显示提交按钮 */
    submitBtn?: boolean
    /** 是否显示重置按钮 */
    resetBtn?: boolean
    /** 其他 formCreate option */
    [key: string]: any
  }
  /** 示例数据（帮助文档） */
  example?: Record<string, unknown>
  /** 参数说明（Markdown 格式） */
  description?: string
}

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

/**
 * 报表后处理钩子函数签名
 * 在 Carbone 渲染完成后，对生成的 Excel 文件进行额外处理（如隐藏列、合并单元格等）
 * @param outputPath 生成的报表文件路径
 * @param userInput 用户输入的参数（与 builder 中的 userInput 相同）
 * @returns Promise<void>
 */
export type PostProcessHook = (outputPath: string, userInput?: unknown) => Promise<void>

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
 * 模板完整定义（支持泛型输入类型）
 * @template TInput 用户输入参数类型
 */
export interface TemplateDefinition<TInput = unknown> {
  /** 元信息 */
  meta: TemplateMeta
  /**
   * 用户输入参数规则（基于 formCreate）
   * - undefined: 模板不需要用户输入
   * - TemplateInputRule: 使用 formCreate 动态生成表单
   */
  inputRule?: TemplateInputRule
  /** 解析器函数 */
  parser: TemplateParser
  /** 报表数据构建器（类型安全的 userInput） */
  builder: (parsedData: ParsedData, userInput?: TInput) => unknown
  /** Carbone 默认选项（可被 renderOptions 覆写） */
  carboneOptions?: CarboneRenderOptions
  /** 报表后处理钩子（可选，用于对生成的 Excel 文件进行额外处理） */
  postProcess?: (outputPath: string, userInput?: TInput) => Promise<void>
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
  /** 用户输入参数（模板特定参数，如年月等） */
  userInput?: unknown
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
