/**
 * 读取 XLSX 文件的原始内容（使用JSZip）
 */

import JSZip from 'jszip'
import fs from 'node:fs'
import path from 'node:path'

async function readXlsxRaw(): Promise<void> {
  const filePath = path.join(
    process.cwd(),
    'output/month1carbone-2025年10月-20251111-171030.xlsx'
  )

  console.log('读取文件:', filePath)
  console.log()

  try {
    const data = fs.readFileSync(filePath)
    const zip = await JSZip.loadAsync(data)

    // 查找 sharedStrings.xml
    const sharedStrings = zip.file('xl/sharedStrings.xml')
    if (sharedStrings) {
      console.log('=== xl/sharedStrings.xml ===')
      const content = await sharedStrings.async('string')
      console.log(content)
      console.log()
    } else {
      console.log('⚠️  未找到 sharedStrings.xml')
    }

    // 查找第一个 worksheet
    const worksheet1 = zip.file('xl/worksheets/sheet1.xml')
    if (worksheet1) {
      console.log('=== xl/worksheets/sheet1.xml (前2000字符) ===')
      const content = await worksheet1.async('string')
      console.log(content.substring(0, 2000))
      console.log()

      // 查找第一个单元格的内容
      const cellMatch = content.match(/<c r="A1"[^>]*>.*?<\/c>/s)
      if (cellMatch) {
        console.log('=== A1 单元格 XML ===')
        console.log(cellMatch[0])
        console.log()
      }
    } else {
      console.log('⚠️  未找到 sheet1.xml')
    }
  } catch (error) {
    console.error('读取失败:', error)
  }
}

readXlsxRaw()
