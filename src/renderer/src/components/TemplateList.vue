<script setup lang="ts">
import { onMounted, ref, computed } from 'vue'
import { useTemplates } from '../composables/useTemplates'
import { Document, Refresh, Search } from '@element-plus/icons-vue'

const { templates, loadingTemplates, currentTemplateId, selectTemplate, loadTemplates } =
  useTemplates()
const keyword = ref('')

const filtered = computed(() => {
  const k = keyword.value.trim().toLowerCase()
  if (!k) return templates.value
  return templates.value.filter((t: any) => {
    const name = t.name || t.meta?.name || ''
    const filename = t.filename || t.meta?.filename || ''
    const id = t.id || t.meta?.id || ''
    return `${name} ${filename} ${id}`.toLowerCase().includes(k)
  })
})

function choose(t: any) {
  const id = t.id || t.meta?.id
  if (id) {
    selectTemplate(id)
  }
}
onMounted(() => {
  loadTemplates()
})
</script>

<template>
  <div class="template-list">
    <div class="header">
      <h2 class="title">报表模板</h2>
      <el-button
        :icon="Refresh"
        :loading="loadingTemplates"
        circle
        size="small"
        @click="loadTemplates"
      />
    </div>

    <div class="search-box">
      <el-input
        v-model="keyword"
        placeholder="搜索模板名称、文件名或 ID..."
        clearable
        size="large"
        class="search-input"
      >
        <template #prefix>
          <el-icon class="search-icon"><Search /></el-icon>
        </template>
        <template v-if="keyword" #suffix>
          <span class="search-count">{{ filtered.length }} 个结果</span>
        </template>
      </el-input>
    </div>

    <div v-loading="loadingTemplates" class="cards-container">
      <el-empty
        v-if="!loadingTemplates && filtered.length === 0"
        description="暂无模板"
        :image-size="80"
      />
      <div v-else class="cards">
        <div
          v-for="t in filtered"
          :key="t.id || t.meta?.id"
          class="template-card"
          :class="{
            active: currentTemplateId === (t.id || t.meta?.id)
          }"
          @click="choose(t)"
        >
          <!-- 选中指示器 -->
          <div v-if="currentTemplateId === (t.id || t.meta?.id)" class="card-indicator">
            <div class="indicator-dot"></div>
          </div>

          <div class="card-icon">
            <el-icon :size="28"><Document /></el-icon>
          </div>

          <div class="card-content">
            <div class="card-title">{{ t.name || t.meta?.name || '未命名' }}</div>
            <div class="card-filename" :title="t.filename || t.meta?.filename">
              {{ t.filename || t.meta?.filename || '' }}
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.template-list {
  display: flex;
  flex-direction: column;
  height: 100%;
  background: white;
  border-radius: 12px;
  padding: 20px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
}

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
}

.title {
  font-size: 18px;
  font-weight: 600;
  color: #303133;
  margin: 0;
}

.search-box {
  margin-bottom: 16px;
}

.search-input {
  transition: all 0.3s;
}

.search-input :deep(.el-input__wrapper) {
  border-radius: 10px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
  transition: all 0.3s;
}

.search-input :deep(.el-input__wrapper):hover {
  box-shadow: 0 4px 12px rgba(79, 70, 229, 0.15);
}

.search-input :deep(.el-input__wrapper.is-focus) {
  box-shadow: 0 4px 16px rgba(79, 70, 229, 0.25);
  border-color: #4f46e5;
}

.search-icon {
  color: #4f46e5;
}

.search-count {
  font-size: 12px;
  color: #6b7280;
  font-weight: 500;
  padding-right: 8px;
}

.cards-container {
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  margin: 0 -4px;
  padding: 0 4px;
}

/* 自定义滚动条 */
.cards-container::-webkit-scrollbar {
  width: 6px;
}

.cards-container::-webkit-scrollbar-track {
  background: transparent;
}

.cards-container::-webkit-scrollbar-thumb {
  background: rgba(0, 0, 0, 0.1);
  border-radius: 3px;
}

.cards-container::-webkit-scrollbar-thumb:hover {
  background: rgba(0, 0, 0, 0.2);
}

.cards {
  display: flex;
  flex-direction: column;
  gap: 12px;
  padding: 4px 0;
}

.template-card {
  position: relative;
  display: flex;
  align-items: center;
  gap: 14px;
  padding: 16px;
  background: #ffffff;
  border: 2px solid #e5e7eb;
  border-radius: 12px;
  cursor: pointer;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  overflow: hidden;
}

.template-card:hover {
  border-color: #4f46e5;
  box-shadow: 0 6px 20px rgba(79, 70, 229, 0.15);
  transform: translateY(-3px);
}

.template-card.active {
  background: linear-gradient(135deg, #eef2ff 0%, #e0e7ff 100%);
  border-color: #4f46e5;
  box-shadow: 0 6px 24px rgba(79, 70, 229, 0.25);
  transform: translateY(-2px);
}

.template-card.active .card-icon {
  background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
  color: white;
  box-shadow: 0 4px 12px rgba(79, 70, 229, 0.4);
}

/* 最近使用徽章 */
.recent-badge {
  position: absolute;
  top: 12px;
  right: 12px;
  padding: 4px 10px;
  background: linear-gradient(135deg, #f59e0b 0%, #f97316 100%);
  color: white;
  font-size: 11px;
  font-weight: 600;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgba(245, 158, 11, 0.3);
  z-index: 2;
}

.card-icon {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  width: 52px;
  height: 52px;
  background: linear-gradient(135deg, #f9fafb 0%, #f3f4f6 100%);
  border-radius: 10px;
  color: #4f46e5;
  transition: all 0.3s;
  border: 1px solid #e5e7eb;
}

.template-card:hover .card-icon {
  transform: scale(1.05);
  box-shadow: 0 4px 12px rgba(79, 70, 229, 0.2);
}

.card-content {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.card-title {
  font-size: 15px;
  font-weight: 600;
  color: #111827;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.card-filename {
  font-size: 12px;
  color: #6b7280;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-family: 'Monaco', 'Menlo', monospace;
}

.card-indicator {
  position: absolute;
  top: 16px;
  right: 16px;
}

.indicator-dot {
  width: 8px;
  height: 8px;
  background: #3b82f6;
  border-radius: 50%;
  animation: pulse 2s ease-in-out infinite;
}

@keyframes pulse {
  0%,
  100% {
    opacity: 1;
    transform: scale(1);
  }
  50% {
    opacity: 0.5;
    transform: scale(1.2);
  }
}
</style>
