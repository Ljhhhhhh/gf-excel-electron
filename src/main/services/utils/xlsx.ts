import type { Workbook, Worksheet, Row } from 'exceljs'

export function getWorksheet(workbook: Workbook, ref: number | string): Worksheet | undefined {
  return typeof ref === 'number' ? workbook.worksheets[ref] : workbook.getWorksheet(ref)
}

export function readHeaders(ws: Worksheet, headerRow: number): string[] {
  const row = ws.getRow(headerRow)
  const headers: string[] = []
  row.eachCell({ includeEmpty: false }, (cell) => {
    const v = cell.value as unknown
    const text = v !== null && v !== undefined ? String(v).trim() : ''
    if (text) headers.push(text)
  })
  return headers
}

export function isRowEmpty(row: Row): boolean {
  const values = row.values as unknown[]
  if (!values || values.length <= 1) return true
  let allEmpty = true
  row.eachCell({ includeEmpty: true }, (cell) => {
    const v = normalizeCellValue(cell.value as unknown)
    if (v !== null && v !== undefined && String(v).trim() !== '') {
      allEmpty = false
    }
  })
  return allEmpty
}

export function toISODateString(input: unknown): string {
  if (input instanceof Date) return input.toISOString().split('T')[0]
  if (typeof input === 'number') {
    const d = new Date(input)
    return isNaN(d.getTime()) ? String(input) : d.toISOString().split('T')[0]
  }
  if (typeof input === 'string') {
    const s = input.trim()
    if (!s) return ''
    const d1 = new Date(s)
    if (!isNaN(d1.getTime())) return d1.toISOString().split('T')[0]
    const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/)
    if (m) {
      const y = Number(m[1])
      const mm = Number(m[2])
      const dd = Number(m[3])
      const d = new Date(Date.UTC(y, mm - 1, dd))
      return d.toISOString().split('T')[0]
    }
    return s
  }
  return input == null ? '' : String(input)
}

export function normalizeCellValue(value: unknown): unknown {
  if (value == null) return value
  if (value instanceof Date) return toISODateString(value)
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>
    if ('result' in obj) return (obj as { result: unknown }).result
    const rich = (obj as { richText?: Array<{ text?: string }> }).richText
    if (Array.isArray(rich)) {
      return rich.map((t) => (t && t.text != null ? String(t.text) : '')).join('')
    }
    if ('text' in obj) return String((obj as { text?: unknown }).text)
  }
  return value
}

export function readRows(
  ws: Worksheet,
  headers: string[],
  startRow: number,
  maxRows: number
): Array<Record<string, unknown>> {
  const rows: Record<string, unknown>[] = []
  const maxRowNum = Math.min(ws.rowCount, startRow + maxRows - 1)
  for (let rowNum = startRow; rowNum <= maxRowNum; rowNum++) {
    const row = ws.getRow(rowNum)
    if (isRowEmpty(row)) continue
    const rowData: Record<string, unknown> = {}
    let hasData = false
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const headerIndex = colNumber - 1
      if (headerIndex < headers.length) {
        const header = headers[headerIndex]
        const cellValue = normalizeCellValue(cell.value as unknown)
        if (cellValue !== null && cellValue !== undefined) {
          rowData[header] = cellValue
          hasData = true
        }
      }
    })
    if (hasData) rows.push(rowData)
  }
  return rows
}

export function extractCellValues(
  ws: Worksheet,
  picks: Record<string, string>,
  transforms?: Record<string, (v: unknown) => unknown>
): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  for (const k of Object.keys(picks)) {
    const addr = picks[k]
    const cell = ws.getCell(addr)
    const raw = normalizeCellValue(cell.value as unknown)
    const fn = transforms?.[k]
    out[k] = fn ? fn(raw) : raw
  }
  return out
}

export function parseSimpleTable(input: {
  workbook: Workbook
  sheets: Array<number | string>
  headerRow: number
  dataStartRow: number
  maxRows: number
}): {
  rows: Record<string, unknown>[]
  headers: string[]
  summary: { totalRows: number; totalSheets: number }
} {
  const { workbook, sheets, headerRow, dataStartRow, maxRows } = input
  const allRows: Record<string, unknown>[] = []
  let headers: string[] = []
  let processedSheets = 0
  for (const sheetRef of sheets) {
    const ws = getWorksheet(workbook, sheetRef)
    if (!ws) continue
    const currentHeaders = readHeaders(ws, headerRow)
    if (headers.length === 0) headers = currentHeaders
    const rows = readRows(ws, headers, dataStartRow, maxRows)
    allRows.push(...rows)
    processedSheets++
  }
  return {
    rows: allRows,
    headers,
    summary: {
      totalRows: allRows.length,
      totalSheets: processedSheets
    }
  }
}
