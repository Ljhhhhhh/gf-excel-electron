<script setup lang="ts">
import { ref, watch, computed } from 'vue'
import { trpc } from '../utils/trpc'
import { FolderOpened, Document, CircleClose, Upload } from '@element-plus/icons-vue'

const props = defineProps<{
  label: string
  modelValue: string | null
  type: 'file' | 'dir'
  disabled?: boolean
  required?: boolean
}>()
const emit = defineEmits<{ 'update:model-value': [value: string | null] }>()
const localValue = ref<string | null>(props.modelValue || null)

const displayText = computed(() => {
  const v = localValue.value || ''
  if (!v) return '未选择'
  const parts = v.split(/\\|\//)
  return parts[parts.length - 1]
})

const fullPath = computed(() => {
  return localValue.value || '点击按钮选择'
})

// 获取文件图标
const fileIcon = computed(() => {
  if (props.type === 'dir') return FolderOpened
  return Document
})

// 获取文件扩展名
const fileExtension = computed(() => {
  if (!localValue.value) return ''
  const parts = localValue.value.split('.')
  return parts.length > 1 ? parts[parts.length - 1].toUpperCase() : ''
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
    if (!res.canceled && res.filePath) emit('update:model-value', res.filePath)
  } else {
    const res = await trpc.file.selectOutputDir.query()
    if (!res.canceled && res.dirPath) emit('update:model-value', res.dirPath)
  }
}

function clear() {
  if (props.disabled) return
  emit('update:model-value', null)
}
</script>

<template>
  <div class="file-path-field">
    <div class="field-header">
      <label class="field-label">
        <el-icon :size="24" class="label-icon">
          <component :is="fileIcon" />
        </el-icon>
        <span>{{ label }}</span>
        <el-tag v-if="required" size="small" type="danger" effect="plain">必填</el-tag>
      </label>
    </div>

    <!-- 未选择状态 -->
    <div v-if="!localValue" class="empty-state">
      <div class="empty-content">
        <el-icon class="empty-icon" :size="28">
          <Upload />
        </el-icon>
        <div class="empty-text">
          <p class="empty-title">{{ type === 'file' ? '选择源数据文件' : '选择输出目录' }}</p>
          <p class="empty-hint">点击按钮进行选择</p>
        </div>
      </div>
      <el-button :disabled="disabled" type="primary" size="default" :icon="Upload" @click="select">
        {{ type === 'file' ? '选择文件' : '选择目录' }}
      </el-button>
    </div>

    <!-- 已选择状态 -->
    <div v-else class="file-card">
      <div class="file-info">
        <div class="file-icon-wrapper">
          <el-icon class="file-main-icon">
            <component :is="fileIcon" />
          </el-icon>
          <span v-if="fileExtension" class="file-extension">{{ fileExtension }}</span>
        </div>
        <div class="file-details">
          <div class="file-name">{{ displayText }}</div>
          <div class="file-path" :title="fullPath">
            <el-icon :size="12"><FolderOpened /></el-icon>
            <span>{{ fullPath }}</span>
          </div>
        </div>
      </div>
      <div class="file-actions">
        <el-button
          :disabled="disabled"
          type="primary"
          plain
          size="small"
          :icon="fileIcon"
          @click="select"
        >
          重新选择
        </el-button>
        <el-button
          :disabled="disabled"
          type="danger"
          plain
          size="small"
          :icon="CircleClose"
          @click="clear"
        >
          清除
        </el-button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.file-path-field {
  display: flex;
  flex-direction: column;
  gap: 12px;
  width: 100%;
}

/* 字段标题 */
.field-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.field-label {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 14px;
  font-weight: 600;
  color: #303133;
}

.field-label span {
  white-space: nowrap;
}

.label-icon {
  color: #4f46e5;
  flex-shrink: 0;
}

/* 空状态 */
.empty-state {
  display: flex;
  flex-direction: row;
  align-items: center;
  justify-content: space-between;
  gap: 20px;
  padding: 20px 24px;
  background: linear-gradient(135deg, #f9fafb 0%, #ffffff 100%);
  border: 2px dashed #d1d5db;
  border-radius: 12px;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

.empty-state:hover {
  border-color: #4f46e5;
  background: linear-gradient(135deg, #eef2ff 0%, #f9fafb 100%);
  box-shadow: 0 4px 12px rgba(79, 70, 229, 0.1);
}

.empty-content {
  flex: 1;
  display: flex;
  flex-direction: row;
  align-items: center;
  gap: 16px;
}

.empty-icon {
  color: #9ca3af;
  transition: color 0.3s;
  flex-shrink: 0;
}

.empty-state:hover .empty-icon {
  color: #4f46e5;
}

.empty-text {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.empty-title {
  margin: 0;
  font-size: 14px;
  font-weight: 600;
  color: #374151;
  line-height: 1.4;
}

.empty-hint {
  margin: 0;
  font-size: 12px;
  color: #9ca3af;
  line-height: 1.4;
}

.empty-state .el-button {
  min-width: 120px;
  flex-shrink: 0;
  font-weight: 600;
  border-radius: 8px;
  background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
  border: none;
  box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);
  transition: all 0.3s;
}

.empty-state .el-button:hover {
  box-shadow: 0 6px 16px rgba(79, 70, 229, 0.4);
}

/* 文件卡片 */
.file-card {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 16px 20px;
  background: linear-gradient(135deg, #ffffff 0%, #f9fafb 100%);
  border: 2px solid #e5e7eb;
  border-radius: 12px;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

.file-card:hover {
  border-color: #4f46e5;
  box-shadow: 0 4px 16px rgba(79, 70, 229, 0.15);
}

.file-info {
  flex: 1;
  display: flex;
  align-items: center;
  gap: 16px;
  min-width: 0;
}

.file-icon-wrapper {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  width: 48px;
  height: 48px;
  background: linear-gradient(135deg, #eef2ff 0%, #e0e7ff 100%);
  border-radius: 10px;
  flex-shrink: 0;
  box-shadow: 0 2px 8px rgba(79, 70, 229, 0.15);
}

.file-main-icon {
  color: #4f46e5;
  font-size: 24px;
}

.file-extension {
  position: absolute;
  bottom: -4px;
  right: -4px;
  padding: 2px 6px;
  font-size: 10px;
  font-weight: 700;
  color: white;
  background: linear-gradient(135deg, #10b981 0%, #059669 100%);
  border-radius: 4px;
  border: 2px solid white;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.file-details {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 6px;
  min-width: 0;
}

.file-name {
  font-size: 14px;
  font-weight: 600;
  color: #111827;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.file-path {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: #6b7280;
  overflow: hidden;
  line-height: 1.4;
}

.file-path span {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.file-actions {
  display: flex;
  gap: 8px;
  flex-shrink: 0;
}

.file-actions .el-button {
  padding: 0 12px;
  font-weight: 500;
  border-radius: 6px;
  transition: all 0.3s;
}

.file-actions .el-button:hover {
  transform: scale(1.02);
}
</style>
