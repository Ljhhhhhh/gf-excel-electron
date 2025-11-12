import { ref } from 'vue'
import { trpc } from '../utils/trpc'

const sourcePath = ref<string | null>(null)
const outputDir = ref<string | null>(null)
const reportName = ref('')
const userInput = ref<Record<string, any>>({})
const generating = ref(false)
const result = ref<any>(null)

function setSourcePath(path: string | null) {
  sourcePath.value = path
}

function setOutputDir(path: string | null) {
  outputDir.value = path
}

function setUserInput(data: Record<string, any>) {
  userInput.value = data
}

async function generate(templateId: string) {
  if (!templateId || !sourcePath.value || !outputDir.value) return
  generating.value = true
  result.value = null
  try {
    const res = await trpc.report.generate.mutate({
      templateId,
      sourcePath: sourcePath.value!,
      outputDir: outputDir.value!,
      reportName: reportName.value || undefined,
      userInput: Object.keys(userInput.value).length > 0 ? userInput.value : undefined
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
    generate
  }
}

