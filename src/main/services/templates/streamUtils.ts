import type ExcelJS from 'exceljs'
import type { Row } from 'exceljs'
import { setImmediate as setImmediatePromise } from 'node:timers/promises'

export interface WorksheetStreamOptions {
  /** 工厂函数：每次调用返回一个新的 WorkbookReader 实例 */
  readerFactory: () => ExcelJS.stream.xlsx.WorkbookReader
  /** 目标工作表（名称或索引） */
  sheet: string | number
  /** 起始行（1-based），默认 1 */
  startRow?: number
  /** 最大处理行数（达到后立即停止），默认无限制 */
  maxRows?: number
  /** 每处理多少行主动让出事件循环，默认 2000 */
  rowYieldInterval?: number
}

export interface StreamWorksheetRowsResult {
  /** 实际处理的行数 */
  processedRows: number
  /** 目标工作表名称（如果可用） */
  sheetName?: string
}

const DEFAULT_ROW_YIELD_INTERVAL = 2000

/**
 * 在指定工作表上以流式方式逐行处理数据。
 *
 * @param options 流式读取配置
 * @param rowHandler 行处理函数；返回 false 可提前终止
 */
export async function streamWorksheetRows(
  options: WorksheetStreamOptions,
  rowHandler: (row: Row, ctx: { rowIndex: number }) => void | boolean | Promise<void | boolean>
): Promise<StreamWorksheetRowsResult> {
  const {
    readerFactory,
    sheet,
    startRow = 1,
    maxRows = Number.POSITIVE_INFINITY,
    rowYieldInterval = DEFAULT_ROW_YIELD_INTERVAL
  } = options

  const workbookReader = readerFactory()
  const targetSheetName = typeof sheet === 'string' ? sheet : null
  const targetSheetIndex = typeof sheet === 'number' ? sheet : null

  let sheetIndex = 0
  let processedSheet = false
  let processedRows = 0
  let resolvedSheetName: string | undefined

  for await (const worksheetReader of workbookReader) {
    const currentSheetName = (worksheetReader as any).name as string | undefined
    const isTarget =
      (targetSheetName && currentSheetName === targetSheetName) ||
      (targetSheetIndex !== null && sheetIndex === targetSheetIndex)

    if (!isTarget) {
      sheetIndex++
      continue
    }

    resolvedSheetName = currentSheetName
    let rowIndex = 0
    for await (const row of worksheetReader) {
      rowIndex++
      if (rowIndex < startRow) {
        continue
      }
      if (processedRows >= maxRows) {
        break
      }

      const shouldContinue = await rowHandler(row as Row, { rowIndex })
      processedRows++

      if (rowYieldInterval > 0 && rowIndex % rowYieldInterval === 0) {
        await setImmediatePromise()
      }

      if (shouldContinue === false) {
        break
      }
    }

    processedSheet = true
    break
  }

  if (!processedSheet) {
    throw new Error(`未找到指定工作表: ${typeof sheet === 'string' ? sheet : `#${sheet}`}`)
  }

  return {
    processedRows,
    sheetName: resolvedSheetName
  }
}
