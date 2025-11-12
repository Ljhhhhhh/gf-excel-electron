<script setup lang="ts">
import { onMounted, ref, computed } from 'vue'
import { useTemplates } from '../composables/useTemplates'

const { templates, loadingTemplates, currentTemplateId, selectTemplate, loadTemplates } = useTemplates()
const keyword = ref('')

const filtered = computed(() => {
  const k = keyword.value.trim().toLowerCase()
  if (!k) return templates.value
  return templates.value.filter((t: any) => `${t.name} ${t.filename} ${t.id}`.toLowerCase().includes(k))
})

function choose(t: any) {
  selectTemplate(t.id)
}

onMounted(() => {
  loadTemplates()
})
</script>

<template>
  <div class="template-list" v-loading="loadingTemplates">
    <div class="toolbar">
      <el-input v-model="keyword" placeholder="搜索模板" clearable class="search" />
      <el-button :loading="loadingTemplates" type="primary" @click="loadTemplates">刷新列表</el-button>
    </div>
    <el-empty v-if="!loadingTemplates && filtered.length===0" description="暂无模板" />
    <div v-else class="cards">
      <el-card v-for="t in filtered" :key="t.id" class="card" :class="{ active: currentTemplateId===t.id }" shadow="hover" @click="choose(t)">
        <div class="title">{{ t.name }}</div>
        <div class="meta">ID：{{ t.id }}</div>
        <div class="meta">文件：{{ t.filename }}</div>
      </el-card>
    </div>
  </div>
</template>

<style scoped>
.template-list { display:flex; flex-direction: column; gap: 12px; }
.toolbar { display:flex; gap:8px; }
.search { flex:1; }
.cards { display:flex; flex-direction: column; gap: 12px; }
.card { border: 1px solid rgba(255,255,255,0.08); cursor: pointer; }
.active { border-color: #409eff; }
.title { font-weight:600; }
.meta { color:#909399; font-size:12px; margin-top:4px; }
</style>

