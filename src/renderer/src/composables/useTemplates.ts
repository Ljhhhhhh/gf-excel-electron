import { ref, computed } from 'vue'
import { trpc } from '../utils/trpc'

const templates = ref<any[]>([])
const loadingTemplates = ref(false)
const currentTemplateId = ref<string | null>(null)
const currentTemplateName = ref<string>('')

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
  const t = templates.value.find((x) => x.id === id)
  currentTemplateName.value = t ? t.name : ''
}

export function useTemplates() {
  return {
    templates,
    loadingTemplates,
    hasTemplates,
    currentTemplateId,
    currentTemplateName,
    loadTemplates,
    selectTemplate
  }
}

