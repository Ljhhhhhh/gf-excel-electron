<script setup lang="ts">
import { ref, watch, computed } from 'vue'
import { trpc } from '../utils/trpc'

const props = defineProps<{ label: string; modelValue: string | null; type: 'file' | 'dir'; disabled?: boolean }>()
const emit = defineEmits<{ 'update:modelValue': [value: string | null] }>()
const localValue = ref<string | null>(props.modelValue || null)

const displayText = computed(() => {
  const v = localValue.value || ''
  if (!v) return ''
  const parts = v.split(/\\|\//)
  return parts[parts.length - 1]
})

watch(
  () => props.modelValue,
  (v) => {
    localValue.value = v || null
  }
)

async function select() {
  if (props.disabled) return
  if (props.type === 'file') {
    const res = await trpc.file.selectSourceFile.query()
    if (!res.canceled && res.filePath) emit('update:modelValue', res.filePath)
  } else {
    const res = await trpc.file.selectOutputDir.query()
    if (!res.canceled && res.dirPath) emit('update:modelValue', res.dirPath)
  }
}
</script>

<template>
  <el-form-item :label="label">
    <div class="field">
      <el-input class="truncate-input" :model-value="displayText" placeholder="未选择" disabled size="small" />
      <el-button :disabled="disabled" size="small" class="picker" @click="select">{{ type==='file' ? '选择文件' : '选择目录' }}</el-button>
    </div>
  </el-form-item>
</template>

<style scoped>
.field { display: grid; grid-template-columns: 1fr; gap: 6px; }
.picker { width: max-content; }
.truncate-input :deep(.el-input__inner) { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
</style>
