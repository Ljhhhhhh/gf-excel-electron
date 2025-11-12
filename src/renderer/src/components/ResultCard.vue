<script setup lang="ts">
import { trpc } from '../utils/trpc'
import { computed } from 'vue'

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
  <el-card v-if="result" class="result-card" shadow="never">
    <template v-if="result.success">
      <el-alert type="success" :closable="false" show-icon title="报表已生成" />
      <div class="details">
        <div class="item"><span class="label">文件</span><span class="value">{{ fileName }}</span></div>
        <div class="item"><span class="label">大小</span><span class="value">{{ (result.size / 1024).toFixed(2) }} KB</span></div>
        <div class="item"><span class="label">时间</span><span class="value">{{ new Date(result.generatedAt).toLocaleString() }}</span></div>
        <div class="item"><span class="label">耗时</span><span class="value">{{ result.duration }}ms</span></div>
      </div>
      <div class="ops">
        <el-button type="primary" size="small" @click="openFolder">打开所在文件夹</el-button>
      </div>
    </template>
    <template v-else>
      <el-alert type="error" :closable="false" show-icon :title="'生成失败'" :description="result.error" />
    </template>
  </el-card>
</template>

<style scoped>
.details { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; margin-top: 8px; }
.item { display: flex; gap: 8px; }
.label { color: #909399; }
.value { color: #303133; }
.ops { margin-top: 8px; }
</style>
