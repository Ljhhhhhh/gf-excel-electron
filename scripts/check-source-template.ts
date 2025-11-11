/**
 * 检查源模板文件的格式
 */

import JSZip from 'jszip'
import fs from 'node:fs'
import path from 'node:path'

async function checkSourceTemplate(): Promise<void> {
  const filePath = path.join(process.cwd(), 'public/reportTemplates/month1carbone.xlsx')

  console.log('检查源模板文件:', filePath)
  console.log()

  try {
    const data = fs.readFileSync(filePath)
    const zip = await JSZip.loadAsync(data)

    const worksheet1 = zip.file('xl/worksheets/sheet1.xml')
    if (worksheet1) {
      const content = await worksheet1.async('string')
      console.log('=== xl/worksheets/sheet1.xml (前3000字符) ===')
      console.log(content.substring(0, 3000))
      console.log()

      // 查找 A1 单元格
      const cellMatch = content.match(/<c r="A1"[^>]*>.*?<\/c>/s)
      if (cellMatch) {
        console.log('=== A1 单元格 XML ===')
        console.log(cellMatch[0])
        console.log()

        // 检查单元格类型
        const typeMatch = cellMatch[0].match(/t="([^"]+)"/)
        if (typeMatch) {
          console.log('单元格类型:', typeMatch[1])
        }
      }
    }

    // 检查是否有 sharedStrings
    const sharedStrings = zip.file('xl/sharedStrings.xml')
    if (sharedStrings) {
      console.log('✅ 模板使用 sharedStrings.xml')
      const content = await sharedStrings.async('string')
      console.log(content)
    } else {
      console.log('⚠️  模板使用 inlineStr（可能导致 Carbone 兼容性问题）')
    }
  } catch (error) {
    console.error('读取失败:', error)
  }
}

checkSourceTemplate()
