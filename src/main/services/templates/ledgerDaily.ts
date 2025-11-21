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
}

interface LedgerDailyUserInput {
  /** 目标日期，格式 YYYYMMDD */
  date: string
}

const EXTRA_SOURCE_IDS = {
  loan: 'loanDetail',
  factoringRepay: 'factoringRepay',
  refactoringRepay: 'refactoringRepay'
}

function resolvePythonScriptPath(): string {
  const devPath = path.join(process.cwd(), 'resources', 'python', 'ledger_daily.py')
  const prodPath = path.join(process.resourcesPath ?? process.cwd(), 'python', 'ledger_daily.py')
  return fs.existsSync(devPath) ? devPath : prodPath
}

function normalizeInputDate(value: string): string {
  if (!/^\d{8}$/.test(value)) {
    throw new Error('日期格式需为 YYYYMMDD')
  }
  return value
}

function ensureParsedData(data: Partial<LedgerDailyParsedData>): LedgerDailyParsedData {
  const { ledgerPath, loanPath, factoringRepayPath, refactoringRepayPath } = data
  if (!ledgerPath || !loanPath || !factoringRepayPath || !refactoringRepayPath) {
    throw new Error('缺少必需的台账/放款/还款数据源路径')
  }
  return {
    ledgerPath,
    loanPath,
    factoringRepayPath,
    refactoringRepayPath
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
    refactoringRepayPath: extra[EXTRA_SOURCE_IDS.refactoringRepay]?.path
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

  const pythonExecutable = process.env.PYTHON_PATH || 'python'
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
    title: '目标日期（YYYYMMDD）',
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
      '按指定日期从“放款明细+保理/再保理融资还款明细”增量追加《融资及还款明细》，保持原公式/样式不变（Python openpyxl 实现）',
    sourceLabel: '现有台账（xlsx）',
    extraSources: [
      {
        id: EXTRA_SOURCE_IDS.loan,
        label: '放款明细',
        supportedExts: ['xlsx'],
        loadStrategy: 'stream',
        description: 'P 列按日期降序，按日期精确匹配'
      },
      {
        id: EXTRA_SOURCE_IDS.factoringRepay,
        label: '保理融资还款明细',
        supportedExts: ['xlsx'],
        loadStrategy: 'stream',
        description: 'AE=日期且 AB=本金，按 AH 交易流水升序'
      },
      {
        id: EXTRA_SOURCE_IDS.refactoringRepay,
        label: '再保理融资还款明细',
        supportedExts: ['xlsx'],
        loadStrategy: 'stream',
        description: 'AE=日期且 AB=本金，按 AH 交易流水升序'
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

### 数据源
- 主文件：现有台账（xlsx）
- 额外：放款明细、保理融资还款明细、再保理融资还款明细（均 xlsx）
    `.trim()
  },
  parser: ledgerDailyParser,
  streamParser: ledgerDailyStreamParser,
  excelRenderer: renderLedgerDailyWithPython
}
