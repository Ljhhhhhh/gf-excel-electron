/**
 * 创建修正后的模板文件（扁平格式占位符）
 */

import ExcelJS from 'exceljs'
import path from 'path'

async function createFixedTemplate(): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Sheet1')

  // 设置列宽
  worksheet.getColumn(1).width = 122.5

  // 设置行高和样式
  const row1 = worksheet.getRow(1)
  row1.height = 159

  const row2 = worksheet.getRow(2)
  row2.height = 114

  const row3 = worksheet.getRow(3)
  row3.height = 114

  // A1: 第一段内容（修正后的占位符 - 扁平格式）
  worksheet.getCell('A1').value =
    '{queryYear}年{queryMonth}月公司新增放款{total}万元（基建占比{infrastructurePercentage}%，医药占比{medicinePercentage}%，再保理占比{factoringPercentage}%），\n' +
    '受让对应的应收账款{accountsReceivable}万元，平均保理利率{factoringRate}%，新增合作客户{customerTotal}家，受理业务{businessCount}笔，已发生实际业务合作的核心企业新增【】家；'

  worksheet.getCell('A1').alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' }

  // A2: 第二段内容（暂时用【】占位符）
  worksheet.getCell('A2').value =
    '截止【】年【】月【】日，公司应收账款余额【】万元，放款余额【】万元，业务不良率为【】%，\n' +
    '累计受让应收账款金额【】万元，累计放款【】万元（基建占比【】%，医药占比【】%,再保理占比【】%）,累计还款【】万元；'

  worksheet.getCell('A2').alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' }

  // A3: 第三段内容（暂时用【】占位符）
  worksheet.getCell('A3').value =
    '展业至今，公司合作客户【】家，累计受理业务【】笔，已发生实际业务合作的核心企业【】家，\n' +
    '累计受让应收账款【】亿元，累计放款【】亿元（基建占比【】%，医药占比【】%,再保理占比【】%）；放款余额【】亿元；业务不良率为【】%。'

  worksheet.getCell('A3').alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' }

  // 保存文件
  const outputPath = path.join(process.cwd(), 'public/reportTemplates/month1carbone-fixed.xlsx')
  await workbook.xlsx.writeFile(outputPath)

  console.log('✅ 修正后的模板已创建:', outputPath)
  console.log()
  console.log('占位符格式（扁平结构）:')
  console.log('  {queryYear}, {queryMonth}, {total}, {infrastructurePercentage},')
  console.log('  {medicinePercentage}, {factoringPercentage}, {accountsReceivable},')
  console.log('  {factoringRate}, {customerTotal}, {businessCount}')
  console.log()
  console.log('请将此文件重命名为 month1carbone.xlsx 替换原文件')
}

createFixedTemplate().catch((error) => {
  console.error('创建失败:', error)
  process.exit(1)
})
