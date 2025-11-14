import { ref } from 'vue'
import { trpc } from '../utils/trpc'

const sourcePath = ref<string | null>(null)
const outputDir = ref<string | null>(null)
const reportName = ref('')
const userInput = ref<Record<string, any>>({})
const generating = ref(false)
const result = ref<any>(null)
// 记录每个模板自定义的额外数据源（如“放款明细”），保持响应式便于多视图共享
const extraSources = ref<Record<string, string | null>>({})

function setSourcePath(path: string | null) {
  sourcePath.value = path
}

function setOutputDir(path: string | null) {
  outputDir.value = path
}

function setUserInput(data: Record<string, any>) {
  userInput.value = data
}

// 单独更新某个额外数据源的路径
function setExtraSource(id: string, path: string | null) {
  extraSources.value = {
    ...extraSources.value,
    [id]: path
  }
}

// 切换模板时清空额外数据源，避免不同模板之间串数据
function resetExtraSources() {
  extraSources.value = {}
}

async function generate(templateId: string) {
  if (!templateId || !sourcePath.value || !outputDir.value) return
  generating.value = true
  result.value = null
  try {
    const extraSourcePayload: Record<string, string> = {}
    Object.entries(extraSources.value).forEach(([key, value]) => {
      if (value) {
        extraSourcePayload[key] = value
      }
    })
    const res = await trpc.report.generate.mutate({
      templateId,
      sourcePath: sourcePath.value!,
      outputDir: outputDir.value!,
      reportName: reportName.value || undefined,
      userInput: Object.keys(userInput.value).length > 0 ? userInput.value : undefined,
      extraSources: Object.keys(extraSourcePayload).length ? extraSourcePayload : undefined
    })
    result.value = { ...res, success: true }
  } catch (error: any) {
    result.value = { success: false, error: error?.message || '未知错误' }
  } finally {
    generating.value = false
  }
}

export function useReport() {
  return {
    sourcePath,
    outputDir,
    reportName,
    userInput,
    generating,
    result,
    setSourcePath,
    setOutputDir,
    setUserInput,
    extraSources,
    setExtraSource,
    resetExtraSources,
    generate
  }
}
