<script setup lang="ts">
import { trpc } from '../utils/trpc'
import { computed } from 'vue'
import { SuccessFilled, CircleCloseFilled, FolderOpened } from '@element-plus/icons-vue'

const props = defineProps<{ result: any }>()

const fileName = computed(() => {
  const p = props.result?.outputPath || ''
  if (!p) return ''
  const parts = p.split(/\\|\//)
  return parts[parts.length - 1]
})

async function openFolder() {
  if (!props.result?.outputPath) return
  await trpc.file.openInFolder.mutate({ path: props.result.outputPath })
}
</script>

<template>
  <div
    v-if="result"
    class="result-card"
    :class="{ success: result.success, error: !result.success }"
  >
    <!-- 成功状态 -->
    <template v-if="result.success">
      <div class="result-content">
        <div class="result-icon success-icon">
          <el-icon><SuccessFilled /></el-icon>
        </div>
        <div class="result-info">
          <div class="result-title">✓ 报表生成成功</div>
          <div class="result-meta">
            <span class="meta-item">{{ fileName }}</span>
            <span class="meta-divider">·</span>
            <span class="meta-item">{{ (result.size / 1024).toFixed(1) }} KB</span>
            <span class="meta-divider">·</span>
            <span class="meta-item">{{ result.duration }}s</span>
          </div>
        </div>
        <el-button type="primary" :icon="FolderOpened" size="default" @click="openFolder">
          打开文件夹
        </el-button>
      </div>
    </template>

    <!-- 失败状态 -->
    <template v-else>
      <div class="result-content">
        <div class="result-icon error-icon">
          <el-icon><CircleCloseFilled /></el-icon>
        </div>
        <div class="result-info">
          <div class="result-title">✗ 生成失败</div>
          <div class="result-error">{{ result.error || '未知错误' }}</div>
        </div>
      </div>
    </template>
  </div>
</template>

<style scoped>
.result-card {
  background: #ffffff;
  border-radius: 10px;
  padding: 16px;
  border: 1.5px solid #e5e7eb;
  animation: slideIn 0.25s ease-out;
}

@keyframes slideIn {
  from {
    opacity: 0;
    transform: translateY(10px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

.result-card.success {
  border-color: #10b981;
  background: linear-gradient(135deg, #f0fdf4 0%, #ffffff 100%);
}

.result-card.error {
  border-color: #ef4444;
  background: linear-gradient(135deg, #fef2f2 0%, #ffffff 100%);
}

/* 内容区域 */
.result-content {
  display: flex;
  align-items: center;
  gap: 12px;
}

.result-icon {
  flex-shrink: 0;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 50%;
  font-size: 20px;
}

.success-icon {
  background: #d1fae5;
  color: #10b981;
}

.error-icon {
  background: #fee2e2;
  color: #ef4444;
}

.result-info {
  flex: 1;
  min-width: 0;
}

.result-title {
  font-size: 14px;
  font-weight: 700;
  color: #111827;
  margin-bottom: 4px;
}

.result-meta {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: #6b7280;
}

.meta-item {
  white-space: nowrap;
}

.meta-divider {
  color: #d1d5db;
}

.result-error {
  font-size: 13px;
  color: #ef4444;
  line-height: 1.4;
}

/* 按钮样式 */
.result-content .el-button {
  flex-shrink: 0;
  font-weight: 600;
  border-radius: 6px;
  transition: all 0.2s;
}

.result-content .el-button:hover {
  transform: scale(1.02);
}
</style>
