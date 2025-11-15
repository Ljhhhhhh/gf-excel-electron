<script setup lang="ts">
import { onMounted, ref, computed } from 'vue'
import { useTemplates } from '../composables/useTemplates'
import { Refresh, Search } from '@element-plus/icons-vue'

const { templates, loadingTemplates, currentTemplateId, selectTemplate, loadTemplates } =
  useTemplates()
const keyword = ref('')

// 极简的类型色彩映射（仅用于视觉区分）
const typeColors: Record<string, { bg: string; text: string; emoji: string }> = {
  month: { bg: 'linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%)', text: '#4F46E5', emoji: '📅' },
  factoring: {
    bg: 'linear-gradient(135deg, #E0F2FE 0%, #BAE6FD 100%)',
    text: '#0284C7',
    emoji: '💼'
  },
  bank: { bg: 'linear-gradient(135deg, #F3E8FF 0%, #E9D5FF 100%)', text: '#9333EA', emoji: '🏦' },
  client: { bg: 'linear-gradient(135deg, #FCE7F3 0%, #FBCFE8 100%)', text: '#DB2777', emoji: '👥' },
  default: { bg: 'linear-gradient(135deg, #F9FAFB 0%, #F3F4F6 100%)', text: '#6B7280', emoji: '📄' }
}

// 根据模板名称推断类型（用于视觉着色）
function getTemplateType(name: string): keyof typeof typeColors {
  const lowerName = name.toLowerCase()
  if (lowerName.includes('月报') || lowerName.includes('月')) return 'month'
  if (lowerName.includes('保理') || lowerName.includes('f103')) return 'factoring'
  if (lowerName.includes('银行')) return 'bank'
  if (lowerName.includes('客户')) return 'client'
  return 'default'
}

const filtered = computed(() => {
  const k = keyword.value.trim().toLowerCase()
  if (!k) return templates.value
  return templates.value.filter((t: any) => {
    const name = t.name || t.meta?.name || ''
    return name.toLowerCase().includes(k)
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
        placeholder="搜索模板名称"
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
          <!-- 动态着色的图标区域 -->
          <div
            class="card-icon"
            :style="{
              background: typeColors[getTemplateType(t.name || t.meta?.name || '')].bg,
              color: typeColors[getTemplateType(t.name || t.meta?.name || '')].text
            }"
          >
            <span class="icon-emoji">
              {{ typeColors[getTemplateType(t.name || t.meta?.name || '')].emoji }}
            </span>
          </div>

          <!-- 模板名称 -->
          <div class="card-content">
            <div class="card-title">{{ t.name || t.meta?.name || '未命名' }}</div>
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
  background: linear-gradient(to bottom, #ffffff 0%, #fafbfc 100%);
  border-radius: 16px;
  padding: 24px;
  border: 1px solid rgba(0, 0, 0, 0.03);
}

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 20px;
  padding-bottom: 16px;
  border-bottom: 1px solid rgba(0, 0, 0, 0.06);
}

.title {
  font-size: 20px;
  font-weight: 700;
  background: linear-gradient(135deg, #1f2937 0%, #4b5563 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  margin: 0;
  letter-spacing: -0.02em;
}

.search-box {
  margin: 0 8px 20px;
}

.search-input {
  transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
}

.search-input :deep(.el-input__wrapper) {
  border-radius: 12px;
  box-shadow:
    0 1px 3px rgba(0, 0, 0, 0.05),
    0 4px 12px rgba(0, 0, 0, 0.03);
  transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
  border: 1.5px solid rgba(0, 0, 0, 0.05);
  background: #ffffff;
}

.search-input :deep(.el-input__wrapper):hover {
  box-shadow:
    0 2px 6px rgba(79, 70, 229, 0.08),
    0 8px 24px rgba(79, 70, 229, 0.12);
  border-color: rgba(79, 70, 229, 0.2);
  transform: translateY(-1px);
}

.search-input :deep(.el-input__wrapper.is-focus) {
  box-shadow:
    0 2px 8px rgba(79, 70, 229, 0.12),
    0 12px 32px rgba(79, 70, 229, 0.18);
  border-color: #4f46e5;
  transform: translateY(-1px);
}

.search-icon {
  color: #4f46e5;
  transition: transform 0.3s;
}

.search-input:focus-within .search-icon {
  transform: scale(1.1);
}

.search-count {
  font-size: 12px;
  color: #6b7280;
  font-weight: 600;
  padding: 2px 10px;
  background: linear-gradient(135deg, #f3f4f6 0%, #e5e7eb 100%);
  border-radius: 6px;
  margin-right: 4px;
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
  gap: 10px;
  padding: 4px 0;
  margin-top: 8px;
}

.template-card {
  position: relative;
  display: flex;
  align-items: center;
  gap: 16px;
  padding: 12px 16px;
  background: #ffffff;
  border: 1.5px solid rgba(0, 0, 0, 0.06);
  border-radius: 14px;
  cursor: pointer;
  transition: all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
  overflow: hidden;
  margin: 0 8px;
}

.template-card::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: linear-gradient(135deg, rgba(79, 70, 229, 0.03) 0%, rgba(99, 102, 241, 0.02) 100%);
  opacity: 0;
  transition: opacity 0.4s;
}

.template-card:hover {
  border-color: rgba(79, 70, 229, 0.4);
  box-shadow:
    0 2px 6px rgba(79, 70, 229, 0.08),
    0 8px 24px rgba(79, 70, 229, 0.12);
  transform: translateY(-4px) scale(1.01);
}

.template-card:hover::before {
  opacity: 1;
}

.template-card.active {
  background: linear-gradient(135deg, #f0f3ff 0%, #e8edff 100%);
  border-color: #4f46e5;
  box-shadow:
    0 2px 6px rgba(79, 70, 229, 0.15),
    0 8px 24px rgba(79, 70, 229, 0.2),
    inset 0 1px 0 rgba(255, 255, 255, 0.8);
  transform: translateY(-2px) scale(1.005);
}

.template-card.active::before {
  opacity: 0;
}

.card-icon {
  position: relative;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  width: 40px;
  height: 40px;
  border-radius: 12px;
  transition: all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
  border: 1px solid rgba(255, 255, 255, 0.5);
  box-shadow:
    0 2px 8px rgba(0, 0, 0, 0.06),
    inset 0 1px 0 rgba(255, 255, 255, 0.6);
  z-index: 1;
}

.icon-emoji {
  font-size: 28px;
  transition: all 0.3s;
  filter: drop-shadow(0 1px 2px rgba(0, 0, 0, 0.1));
}

.template-card:hover .card-icon {
  transform: scale(1.08) rotate(3deg);
  box-shadow:
    0 4px 16px rgba(0, 0, 0, 0.1),
    0 8px 32px rgba(0, 0, 0, 0.08),
    inset 0 1px 0 rgba(255, 255, 255, 0.8);
}

.template-card:hover .icon-emoji {
  transform: scale(1.1);
  filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.15));
}

.template-card.active .card-icon {
  transform: scale(1.05);
  box-shadow:
    0 6px 20px rgba(0, 0, 0, 0.12),
    0 10px 40px rgba(79, 70, 229, 0.2),
    inset 0 1px 0 rgba(255, 255, 255, 0.9);
  border-color: rgba(255, 255, 255, 0.8);
}

.card-content {
  position: relative;
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 4px;
  z-index: 1;
}

.card-title {
  font-size: 15px;
  font-weight: 600;
  color: #111827;
  line-height: 1.4;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  letter-spacing: -0.01em;
  transition: color 0.3s;
}

.template-card:hover .card-title {
  color: #4f46e5;
}

.template-card.active .card-title {
  color: #3730a3;
  font-weight: 700;
}
</style>
