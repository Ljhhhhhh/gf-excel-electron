/**
 * 示例模板：basicDemo
 * 月度报表模板，支持多表聚合
 */

import type { Workbook } from 'exceljs'
import type { TemplateDefinition, ParseOptions, ParsedData } from './types'
import { extractCellValues, parseSimpleTable, toISODateString } from '../utils/xlsx'

// ========== 解析选项接口 ==========

interface Month1ParseOptions extends ParseOptions {
  /** 指定要解析的 sheet 名称或索引列表，默认解析首表 */
  sheets?: Array<string | number>
  /** 表头行索引（1-based），默认为 1 */
  headerRow?: number
  /** 数据起始行索引（1-based），默认为表头行 + 1 */
  dataStartRow?: number
  /** 最大行数限制（防止读取过多数据），默认 10000 */
  maxRows?: number
}

// ========== 解析结果接口 ==========

interface Month1ParsedData {
  /** 表格数据（多表聚合后的统一结构） */
  rows: Record<string, unknown>[]
  /** 表头字段列表 */
  headers: string[]
  /** 元数据（标题、日期等非表格内容） */
  metadata?: {
    title?: string
    date?: string
    [key: string]: unknown
  }
  /** 汇总统计信息 */
  summary?: {
    totalRows: number
    totalSheets: number
  }
}

// ========== 解析器实现 ==========

/**
 * 解析 workbook，提取表格数据
 */
export function parseWorkbook(
  workbook: Workbook,
  parseOptions?: Month1ParseOptions
): Month1ParsedData {
  const options: Required<Month1ParseOptions> = {
    sheets: parseOptions?.sheets ?? [0],
    headerRow: parseOptions?.headerRow ?? 4,
    dataStartRow: parseOptions?.dataStartRow ?? 5,
    maxRows: parseOptions?.maxRows ?? 10000
  }

  const firstSheetRef = options.sheets[0]
  const ws =
    typeof firstSheetRef === 'number'
      ? workbook.worksheets[firstSheetRef]
      : workbook.getWorksheet(firstSheetRef)
  const metadata = ws
    ? extractCellValues(
        ws,
        { title: 'A1', date: 'B2' },
        {
          date: (v) => {
            const s = String(v ?? '')
              .replace(/^[^:]*:\s*/, '')
              .trim()
            return s ? toISODateString(s) : ''
          }
        }
      )
    : {}

  const table = parseSimpleTable({
    workbook,
    sheets: options.sheets,
    headerRow: options.headerRow,
    dataStartRow: options.dataStartRow,
    maxRows: options.maxRows
  })

  return {
    rows: table.rows,
    headers: table.headers,
    metadata: metadata as { title?: string; date?: string },
    summary: table.summary
  }
}

// ========== 数据构建器实现 ==========

/**
 * 构建 Carbone 渲染数据
 */
export function buildReportData(parsedData: ParsedData): unknown {
  const data = parsedData as Month1ParsedData

  // 将中文字段映射为英文字段，匹配 Carbone 模板
  const products = data.rows.map((row) => ({
    id: row['型号'],
    name: row['名称'],
    quantity: row['数量'],
    price: row['单价']
  }))

  return {
    // 报表标题（从数据源第1行获取）
    reportTitle: data.metadata?.title || '季度销售报表',
    // 日期（从数据源第2行获取）
    date: data.metadata?.date || new Date().toISOString().split('T')[0],
    // 产品列表（匹配模板中的 {d.products[i].xxx}）
    products: products,
    // 汇总信息
    summary: {
      totalCount: data.rows.length,
      generatedAt: new Date().toISOString().split('T')[0],
      ...data.summary
    },
    // 表头（可选，用于动态表头场景）
    headers: data.headers,
    // 原始元数据（可选）
    metadata: data.metadata
  }
}

// ========== 模板定义与导出 ==========

export const basicTemplate: TemplateDefinition = {
  meta: {
    id: 'basic',
    name: '基础报表模板',
    filename: 'basic.xlsx',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx', 'xls'],
    description: '基础报表模板'
  },
  parser: parseWorkbook,
  builder: buildReportData,
  carboneOptions: {
    lang: 'zh-cn',
    timezone: 'Asia/Shanghai'
  }
}
