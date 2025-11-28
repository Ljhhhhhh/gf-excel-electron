import path from 'node:path'
import fs from 'node:fs'
import { execa } from 'execa'
import type { Workbook } from 'exceljs'
import type { FormCreateRule, ParseOptions, TemplateDefinition } from './types'

interface LedgerDailyParsedData {
  ledgerPath: string
  loanPath: string
  factoringRepayPath: string
  refactoringRepayPath: string
  zhongdengPath: string
  customerPath: string
}

interface LedgerDailyUserInput {
  /** 目标日期，格式 YYYYMMDD */
  date: string
}

const EXTRA_SOURCE_IDS = {
  loan: 'loanDetail',
  factoringRepay: 'factoringRepay',
  refactoringRepay: 'refactoringRepay',
  zhongdeng: 'zhongdeng',
  customer: 'customer'
}

function resolvePythonScriptPath(): string {
  const devPath = path.join(process.cwd(), 'resources', 'python', 'ledger_daily.py')
  const prodPath = path.join(process.resourcesPath ?? process.cwd(), 'python', 'ledger_daily.py')
  return fs.existsSync(devPath) ? devPath : prodPath
}

function resolvePythonExecutable(): string {
  // macOS 直接使用系统 Python
  if (process.platform === 'darwin') {
    return process.env.PYTHON_PATH || 'python3'
  }

  // Windows：开发环境使用 resources/python-embed/python.exe
  const devPath = path.join(process.cwd(), 'resources', 'python-embed', 'python.exe')
  // Windows：生产环境使用打包后的 python-embed/python.exe
  const prodPath = path.join(process.resourcesPath ?? process.cwd(), 'python-embed', 'python.exe')

  if (fs.existsSync(devPath)) {
    return devPath
  }
  if (fs.existsSync(prodPath)) {
    return prodPath
  }
  // 回退到系统 Python（兼容旧环境或用户自定义）
  return process.env.PYTHON_PATH || 'python'
}

function normalizeInputDate(value: string): string {
  if (!/^\d{8}$/.test(value)) {
    throw new Error('日期格式需为 YYYYMMDD')
  }
  return value
}

function ensureParsedData(data: Partial<LedgerDailyParsedData>): LedgerDailyParsedData {
  const {
    ledgerPath,
    loanPath,
    factoringRepayPath,
    refactoringRepayPath,
    zhongdengPath,
    customerPath
  } = data
  if (
    !ledgerPath ||
    !loanPath ||
    !factoringRepayPath ||
    !refactoringRepayPath ||
    !zhongdengPath ||
    !customerPath
  ) {
    throw new Error('缺少必需的台账/放款/还款/中登/客户数据源路径')
  }
  return {
    ledgerPath,
    loanPath,
    factoringRepayPath,
    refactoringRepayPath,
    zhongdengPath,
    customerPath
  }
}

function ledgerDailyParser(
  _workbook: Workbook,
  _parseOptions?: ParseOptions
): LedgerDailyParsedData {
  // 防御：强制使用流式解析，避免加载超大台账文件
  throw new Error(
    'ledgerDaily 模板仅支持流式解析，请确保调用 streamParser（无需在 parser 中完整加载台账）'
  )
}

async function ledgerDailyStreamParser(
  filePath: string,
  parseOptions?: ParseOptions
): Promise<LedgerDailyParsedData> {
  const extra = parseOptions?.extraSources || {}
  return ensureParsedData({
    ledgerPath: filePath,
    loanPath: extra[EXTRA_SOURCE_IDS.loan]?.path,
    factoringRepayPath: extra[EXTRA_SOURCE_IDS.factoringRepay]?.path,
    refactoringRepayPath: extra[EXTRA_SOURCE_IDS.refactoringRepay]?.path,
    zhongdengPath: extra[EXTRA_SOURCE_IDS.zhongdeng]?.path,
    customerPath: extra[EXTRA_SOURCE_IDS.customer]?.path
  })
}

async function renderLedgerDailyWithPython(
  parsedData: unknown,
  userInput: LedgerDailyUserInput | undefined,
  outputPath: string
): Promise<void> {
  if (!userInput) {
    throw new Error('缺少用户输入日期')
  }

  const targetDate = normalizeInputDate(userInput.date)
  const data = ensureParsedData(parsedData as Partial<LedgerDailyParsedData>)
  const scriptPath = resolvePythonScriptPath()
  if (!fs.existsSync(scriptPath)) {
    throw new Error(`未找到台账 Python 脚本: ${scriptPath}`)
  }

  const pythonExecutable = resolvePythonExecutable()
  await execa(
    pythonExecutable,
    [
      scriptPath,
      '--ledger',
      path.resolve(data.ledgerPath),
      '--loan',
      path.resolve(data.loanPath),
      '--factoring-repay',
      path.resolve(data.factoringRepayPath),
      '--refactoring-repay',
      path.resolve(data.refactoringRepayPath),
      '--zhongdeng',
      path.resolve(data.zhongdengPath),
      '--customer',
      path.resolve(data.customerPath),
      '--date',
      targetDate,
      '--output',
      outputPath
    ],
    {
      stdio: 'inherit',
      env: {
        ...process.env,
        PYTHONIOENCODING: 'utf-8'
      }
    }
  )
}

const inputRules: FormCreateRule[] = [
  {
    type: 'Input',
    field: 'date',
    title: '目标日期',
    value: (() => {
      const now = new Date()
      const y = now.getFullYear()
      const m = `${now.getMonth() + 1}`.padStart(2, '0')
      const d = `${now.getDate()}`.padStart(2, '0')
      return `${y}${m}${d}`
    })(),
    props: {
      placeholder: '例如：20251029'
    },
    validate: [
      { required: true, message: '请输入日期', trigger: 'blur' },
      { pattern: '^\\d{8}$', message: '格式需为 YYYYMMDD', trigger: 'blur' }
    ]
  }
]

export const ledgerDailyTemplate: TemplateDefinition<LedgerDailyUserInput> = {
  meta: {
    id: 'ledgerDaily',
    name: '台账',
    ext: 'xlsx',
    supportedSourceExts: ['xlsx'],
    description:
      '按指定日期从“放款明细+保理/再保理融资还款明细+客户表”增量更新台账（融资及还款明细/资产明细/中登登记表/客户表/利息缴纳），保持原公式/样式不变（Python openpyxl 实现）',
    sourceLabel: '现有台账（xlsx）',
    extraSources: [
      {
        id: EXTRA_SOURCE_IDS.loan,
        label: '放款明细',
        supportedExts: ['xlsx']
      },
      {
        id: EXTRA_SOURCE_IDS.factoringRepay,
        label: '保理融资还款明细',
        supportedExts: ['xlsx']
      },
      {
        id: EXTRA_SOURCE_IDS.refactoringRepay,
        label: '再保理融资还款明细',
        supportedExts: ['xlsx']
      },
      {
        id: EXTRA_SOURCE_IDS.zhongdeng,
        label: '中登登记表',
        supportedExts: ['xlsx']
      },
      {
        id: EXTRA_SOURCE_IDS.customer,
        label: '客户表',
        supportedExts: ['xlsx']
      }
    ]
  },
  engine: 'exceljs',
  inputRule: {
    rules: inputRules,
    options: {
      submitBtn: false,
      resetBtn: false,
      labelWidth: '140px'
    },
    description: `
### 规则
- 仅更新《融资及还款明细》：先放款，再保理还款，再再保理还款
- 若目标日期小于等于现有表中的最新日期（W/AE 列），则不追加
- AI 列按 AE=目标日期 & 同 AR 的行合并，并写入 AH 求和
- 利息缴纳：筛选资金费（AE=目标日期 & AB=资金费），保理+再保理依次追加，S 列写入 U*360/T/R
- 客户表：取目标日期放款的申请人/买方/确权方，缺失时按模板第 10 行样式补充并带出下载客户表信息

### 数据源
- 主文件：现有台账（xlsx）
- 额外：放款明细、保理融资还款明细、再保理融资还款明细、中登登记表、客户表（均 xlsx）
    `.trim()
  },
  parser: ledgerDailyParser,
  streamParser: ledgerDailyStreamParser,
  excelRenderer: renderLedgerDailyWithPython
}
