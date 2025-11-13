<script setup lang="ts">
import { ref, watch } from 'vue'
import { trpc } from '../utils/trpc'

const props = defineProps<{
  templateId: string
}>()

const emit = defineEmits<{
  change: [value: Record<string, any>]
  ready: [formData: Record<string, any>]
}>()

// formCreate v3 API 实例和数据
const formApi = ref()
const formData = ref<Record<string, any>>({})
const loading = ref(true)
const error = ref<string>()

// 从后端加载 inputRule
const ruleData = ref<{
  templateId: string
  inputRule: {
    rules: any[]
    options?: any
    example?: Record<string, any>
    description?: string
  } | null
} | null>(null)

// 加载规则
const loadRule = async () => {
  try {
    loading.value = true
    error.value = undefined
    const result = await trpc.template.getInputRule.query({
      templateId: props.templateId
    })
    ruleData.value = result
  } catch (err) {
    error.value = err instanceof Error ? err.message : '加载表单规则失败'
    console.error('[TemplateInputForm] 加载规则失败:', err)
  } finally {
    loading.value = false
  }
}

// 监听 templateId 变化
watch(() => props.templateId, loadRule, { immediate: true })

// 监听表单数据变化（v3 使用 v-model 自动同步）
watch(
  formData,
  (newData) => {
    emit('change', newData)
  },
  { deep: true }
)

// 监听 API 就绪
watch(
  () => formApi.value,
  (api) => {
    if (!api) return
    // 表单就绪后发送初始值
    emit('ready', formData.value)
  }
)

// 暴露验证方法给父组件
const validate = async (): Promise<boolean> => {
  if (!formApi.value) return true // 无表单时认为验证通过
  try {
    await formApi.value.validate()
    return true
  } catch {
    return false
  }
}

// 暴露获取表单数据方法
const getFormData = (): Record<string, any> => {
  return formData.value
}

defineExpose({
  validate,
  getFormData
})
</script>

<template>
  <div class="template-input-form">
    <!-- 加载中 -->
    <div v-if="loading" v-loading="true" style="min-height: 200px" />

    <!-- 加载失败 -->
    <el-alert v-else-if="error" type="error" :closable="false" show-icon>
      <template #title>加载表单失败</template>
      {{ error }}
    </el-alert>

    <!-- 无需参数 -->
    <div v-else-if="!ruleData?.inputRule" class="no-input-hint">
      <el-alert type="info" :closable="false" show-icon>
        <template #title>此模板无需额外参数</template>
        该模板不需要用户输入额外参数，可以直接生成报表。
      </el-alert>
    </div>

    <!-- formCreate v3 动态表单 -->
    <div v-else class="form-container">
      <form-create
        v-model="formData"
        v-model:api="formApi"
        :rule="ruleData.inputRule.rules"
        :option="ruleData.inputRule.options || {}"
      />

      <!-- 参数说明（可选） -->
      <!-- <el-alert
        v-if="ruleData.inputRule.description"
        type="info"
        :closable="false"
        style="margin-top: 16px"
      >
        <div class="description-content" v-html="ruleData.inputRule.description" />
      </el-alert> -->
    </div>
  </div>
</template>

<style scoped>
.template-input-form {
  width: 100%;
}

.no-input-hint {
  padding: 20px 0;
}

.form-container {
  width: 100%;
}

.description-content {
  font-size: 14px;
  line-height: 1.6;
}

.description-content :deep(h3) {
  margin: 8px 0;
  font-size: 14px;
  font-weight: 600;
}

.description-content :deep(ul) {
  margin: 4px 0;
  padding-left: 20px;
}

.description-content :deep(li) {
  margin: 4px 0;
}
</style>
