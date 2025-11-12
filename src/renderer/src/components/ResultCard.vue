<script setup lang="ts">
import { trpc } from '../utils/trpc'
import { computed } from 'vue'
import {
  SuccessFilled,
  CircleCloseFilled,
  FolderOpened,
  Document,
  Timer,
  Calendar
} from '@element-plus/icons-vue'

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
      <div class="result-header success-header">
        <div class="status-icon success-icon">
          <el-icon :size="32"><SuccessFilled /></el-icon>
        </div>
        <div class="status-content">
          <h3 class="status-title">报表生成成功</h3>
          <p class="status-desc">文件已保存到指定目录</p>
        </div>
      </div>

      <div class="result-details">
        <div class="detail-item">
          <div class="detail-icon">
            <el-icon :size="16"><Document /></el-icon>
          </div>
          <div class="detail-content">
            <span class="detail-label">文件名称</span>
            <span class="detail-value">{{ fileName }}</span>
          </div>
        </div>

        <div class="detail-item">
          <div class="detail-icon">
            <el-icon :size="16"><Document /></el-icon>
          </div>
          <div class="detail-content">
            <span class="detail-label">文件大小</span>
            <span class="detail-value">{{ (result.size / 1024).toFixed(2) }} KB</span>
          </div>
        </div>

        <div class="detail-item">
          <div class="detail-icon">
            <el-icon :size="16"><Calendar /></el-icon>
          </div>
          <div class="detail-content">
            <span class="detail-label">生成时间</span>
            <span class="detail-value">{{ new Date(result.generatedAt).toLocaleString() }}</span>
          </div>
        </div>

        <div class="detail-item">
          <div class="detail-icon">
            <el-icon :size="16"><Timer /></el-icon>
          </div>
          <div class="detail-content">
            <span class="detail-label">用时</span>
            <span class="detail-value">{{ result.duration }}ms</span>
          </div>
        </div>
      </div>

      <div class="result-actions">
        <el-button type="primary" size="large" :icon="FolderOpened" @click="openFolder">
          打开文件所在文件夹
        </el-button>
      </div>
    </template>

    <!-- 失败状态 -->
    <template v-else>
      <div class="result-header error-header">
        <div class="status-icon error-icon">
          <el-icon :size="32"><CircleCloseFilled /></el-icon>
        </div>
        <div class="status-content">
          <h3 class="status-title">报表生成失败</h3>
          <p class="status-desc">{{ result.error || '未知错误' }}</p>
        </div>
      </div>
    </template>
  </div>
</template>

<style scoped>
.result-card {
  background: white;
  border-radius: 12px;
  padding: 24px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
  animation: slideIn 0.3s ease-out;
}

@keyframes slideIn {
  from {
    opacity: 0;
    transform: translateY(20px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

.result-card.success {
  border: 2px solid #10b981;
}

.result-card.error {
  border: 2px solid #ef4444;
}

/* 头部区域 */
.result-header {
  display: flex;
  align-items: center;
  gap: 16px;
  padding-bottom: 20px;
  border-bottom: 2px solid #f5f7fa;
}

.status-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 64px;
  height: 64px;
  border-radius: 50%;
  animation: iconPulse 0.5s ease-out;
}

@keyframes iconPulse {
  0% {
    transform: scale(0.8);
    opacity: 0;
  }
  50% {
    transform: scale(1.1);
  }
  100% {
    transform: scale(1);
    opacity: 1;
  }
}

.success-icon {
  background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%);
  color: #10b981;
}

.error-icon {
  background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%);
  color: #ef4444;
}

.status-content {
  flex: 1;
}

.status-title {
  font-size: 18px;
  font-weight: 600;
  color: #303133;
  margin: 0 0 6px 0;
}

.status-desc {
  font-size: 14px;
  color: #606266;
  margin: 0;
}

/* 详情区域 */
.result-details {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 16px;
  padding: 20px 0;
}

.detail-item {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px;
  background: #f8f9fa;
  border-radius: 8px;
  transition: all 0.3s;
}

.detail-item:hover {
  background: #e8ecf1;
  transform: translateX(4px);
}

.detail-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  background: white;
  border-radius: 6px;
  color: #3b82f6;
}

.detail-content {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.detail-label {
  font-size: 12px;
  color: #909399;
}

.detail-value {
  font-size: 14px;
  font-weight: 500;
  color: #303133;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* 操作区域 */
.result-actions {
  display: flex;
  justify-content: center;
  padding-top: 20px;
  border-top: 2px solid #f5f7fa;
}

.result-actions .el-button {
  min-width: 240px;
  font-size: 15px;
  font-weight: 600;
  height: 44px;
  border-radius: 8px;
}

.result-actions .el-button:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
}
</style>
