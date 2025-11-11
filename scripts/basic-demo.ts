/**
 * 命令行测试脚本：报表生成
 * 用于本地验证 excelToData → dataToReport 完整流程
 *
 * 使用方式:
 * pnpm test:report --templateId=month1carbone --source=./test-data.xlsx --outDir=./output [--reportName=my-report.xlsx]
 */

import { parseArgs } from 'node:util'
import path from 'node:path'
import { initTemplates } from '../src/main/services/templates'
import { excelToData } from '../src/main/services/excelToData'
import { dataToReport } from '../src/main/services/dataToReport'

// 模拟 electron app（仅用于路径解析）
const mockApp = {
  isPackaged: false,
  getPath: (name: string) => {
    if (name === 'documents') {
      return path.join(process.cwd(), 'output')
    }
    return process.cwd()
  }
}

// 注入 mock（仅在脚本环境中）
// @ts-ignore - 模拟 Electron app 全局对象用于测试脚本
global.app = mockApp

async function main(): Promise<void> {
  console.log('=== 报表生成测试 ===\n')

  // 解析命令行参数
  const { values } = parseArgs({
    options: {
      templateId: { type: 'string' },
      source: { type: 'string' },
      outDir: { type: 'string' },
      reportName: { type: 'string' }
    }
  })

  const { templateId, source, outDir, reportName } = values

  // 校验必填参数
  if (!templateId || !source || !outDir) {
    console.error('❌ 缺少必填参数')
    console.log('使用方式:')
    console.log(
      '  pnpm test:report --templateId=month1carbone --source=./test-data.xlsx --outDir=./output [--reportName=my-report.xlsx]'
    )
    process.exit(1)
  }

  console.log(`📋 模板 ID: ${templateId}`)
  console.log(`📂 源文件: ${source}`)
  console.log(`📁 输出目录: ${outDir}`)
  if (reportName) {
    console.log(`📝 报表名称: ${reportName}`)
  }
  console.log()

  try {
    // 1. 初始化模板系统
    console.log('🔧 初始化模板系统...')
    initTemplates()
    console.log()

    // 2. Excel → 数据
    console.log('📊 解析 Excel 数据...')
    const startParse = Date.now()
    const dataResult = await excelToData({
      sourcePath: path.resolve(source),
      templateId,
      parseOptions: {}
    })
    const parseDuration = Date.now() - startParse
    console.log(`✅ 解析完成，耗时: ${parseDuration}ms`)
    console.log(`   - 数据源: ${dataResult.sourceMeta.path}`)
    console.log(`   - 文件大小: ${(dataResult.sourceMeta.size / 1024).toFixed(2)} KB`)
    console.log(`   - 工作表: ${dataResult.sourceMeta.sheets.join(', ')}`)
    if (dataResult.warnings.length > 0) {
      console.log(`   ⚠️  警告: ${dataResult.warnings.length} 条`)
      dataResult.warnings.forEach((w) => console.log(`      - ${w.message}`))
    }
    console.log()

    // 3. 数据 → 报表
    console.log('📄 生成报表...')
    const startRender = Date.now()
    const reportResult = await dataToReport({
      templateId,
      parsedData: dataResult.data,
      outputDir: path.resolve(outDir),
      reportName
    })
    const renderDuration = Date.now() - startRender
    console.log(`✅ 报表生成完成，耗时: ${renderDuration}ms`)
    console.log(`   - 输出路径: ${reportResult.outputPath}`)
    console.log(`   - 文件大小: ${(reportResult.size / 1024).toFixed(2)} KB`)
    console.log(`   - 生成时间: ${reportResult.generatedAt.toISOString()}`)
    console.log()

    // 4. 询问是否打开文件夹
    console.log('🎉 测试完成！')
    console.log(`总耗时: ${parseDuration + renderDuration}ms`)
    console.log()
    console.log('💡 提示: 可以使用以下命令打开输出文件夹:')
    console.log(`   explorer "${path.dirname(reportResult.outputPath)}"`)
    console.log()

    // 可选：直接打开文件夹（取消注释以启用）
    // await openFolder(reportResult.outputPath)
  } catch (error) {
    console.error('\n❌ 生成失败:')
    if (error instanceof Error) {
      console.error(`   错误: ${error.message}`)
      if ('code' in error) {
        console.error(`   错误码: ${(error as { code: string }).code}`)
      }
      if ('details' in error) {
        console.error(
          `   详情: ${JSON.stringify((error as { details: unknown }).details, null, 2)}`
        )
      }
    } else {
      console.error(`   ${String(error)}`)
    }
    process.exit(1)
  }
}

main()
