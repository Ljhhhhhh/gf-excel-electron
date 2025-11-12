<script setup lang="ts">
import { onMounted, ref, computed } from 'vue'
import { useRouter } from 'vue-router'
import { useTemplates } from '../composables/useTemplates'
import { trpc } from '../utils/trpc'
import { ElMessage } from 'element-plus'

const router = useRouter()
const { templates, loadingTemplates, currentTemplateId, selectTemplate, loadTemplates } = useTemplates()
const keyword = ref('')

const filtered = computed(() => {
  const k = keyword.value.trim().toLowerCase()
  if (!k) return templates.value
  return templates.value.filter((t: any) => `${t.name} ${t.filename} ${t.id}`.toLowerCase().includes(k))
})

async function validateTemplate(id: string) {
  try {
    const res = await trpc.template.validate.query({ id })
    if (res.valid) {
      ElMessage.success('模板校验通过')
    } else {
      ElMessage.error('模板校验失败')
    }
  } catch (e: any) {
    ElMessage.error(e?.message || '校验失败')
  }
}

function choose(t: any) {
  selectTemplate(t.id)
}

function goReport() {
  if (!currentTemplateId.value) return
  router.push('/report')
}

onMounted(() => {
  loadTemplates()
})
</script>

<template>
  <div class="templates-view" v-loading="loadingTemplates">
    <div class="header">
      <el-input v-model="keyword" placeholder="搜索模板" clearable class="search" />
      <div class="header-actions">
        <el-button :loading="loadingTemplates" type="primary" @click="loadTemplates">刷新列表</el-button>
        <el-button :disabled="!currentTemplateId" @click="goReport">前往生成</el-button>
      </div>
    </div>
    <el-empty v-if="!loadingTemplates && filtered.length===0" description="暂无模板" />
    <el-row v-else :gutter="16" class="grid">
      <el-col v-for="t in filtered" :key="t.id" :xs="24" :sm="12" :md="8" :lg="6">
        <el-card class="card" :class="{ active: currentTemplateId===t.id }" shadow="hover" @click="choose(t)">
          <div class="title">{{ t.name }}</div>
          <div class="meta">ID：{{ t.id }}</div>
          <div class="meta">文件：{{ t.filename }}</div>
          <div class="card-actions">
            <el-button size="small" @click.stop="validateTemplate(t.id)">校验模板</el-button>
            <el-button size="small" type="primary" @click.stop="choose(t)">设为当前</el-button>
          </div>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<style scoped>
.templates-view { padding: 8px; }
.header { display:flex; align-items:center; justify-content: space-between; gap:12px; }
.header-actions { display:flex; gap:8px; }
.search { max-width: 360px; }
.grid { margin-top: 12px; }
.card { border: 1px solid rgba(255,255,255,0.08); }
.title { font-weight:600; }
.meta { color:#909399; font-size:12px; margin-top:4px; }
.card-actions { margin-top:8px; display:flex; gap:8px; }
.active { border-color: #409eff; }
</style>
