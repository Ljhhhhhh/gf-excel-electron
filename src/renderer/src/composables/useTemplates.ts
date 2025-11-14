import { ref, computed } from 'vue'
import { trpc } from '../utils/trpc'

const templates = ref<any[]>([])
const loadingTemplates = ref(false)
const currentTemplateId = ref<string | null>(null)
const currentTemplateName = ref<string>('')
// 在前端直接拿到当前模板元数据（名称、额外数据源配置等）
const currentTemplateMeta = computed(() => {
  if (!currentTemplateId.value) return null
  return templates.value.find((x) => (x.id || x.meta?.id) === currentTemplateId.value) || null
})

const hasTemplates = computed(() => templates.value.length > 0)

async function loadTemplates() {
  loadingTemplates.value = true
  try {
    templates.value = await trpc.template.list.query()
  } finally {
    loadingTemplates.value = false
  }
}

function selectTemplate(id: string) {
  currentTemplateId.value = id
  const t = templates.value.find((x) => (x.id || x.meta?.id) === id)
  currentTemplateName.value = t ? t.name || t.meta?.name || '' : ''
}

export function useTemplates() {
  return {
    templates,
    loadingTemplates,
    hasTemplates,
    currentTemplateId,
    currentTemplateName,
    currentTemplateMeta,
    loadTemplates,
    selectTemplate
  }
}
