<template>
  <div class="report-test">
    <h1>tRPC 报表生成测试</h1>

    <!-- 模板列表 -->
    <section class="section">
      <h2>模板列表</h2>
      <button :disabled="loading.templates" @click="loadTemplates">
        {{ loading.templates ? '加载中...' : '刷新模板列表' }}
      </button>
      <div v-if="templates.length > 0" class="template-list">
        <div
          v-for="template in templates"
          :key="template.id"
          class="template-item"
          :class="{ active: selectedTemplateId === template.id }"
          @click="selectTemplate(template.id)"
        >
          <h3>{{ template.name }}</h3>
          <p>ID: {{ template.id }}</p>
          <p>文件: {{ template.filename }}</p>
        </div>
      </div>
      <p v-else class="empty">暂无模板</p>
    </section>

    <!-- 报表生成 -->
    <section v-if="selectedTemplateId" class="section">
      <h2>生成报表</h2>
      <div class="form">
        <div class="form-item">
          <label>模板ID:</label>
          <input type="text" :value="selectedTemplateId" disabled />
        </div>
        <div class="form-item">
          <label>源文件:</label>
          <input type="text" :value="sourcePath || '未选择'" disabled />
          <button :disabled="generating" @click="selectSourceFile">选择文件</button>
        </div>
        <div class="form-item">
          <label>输出目录:</label>
          <input type="text" :value="outputDir || '未选择'" disabled />
          <button :disabled="generating" @click="selectOutputDir">选择目录</button>
        </div>
        <div class="form-item">
          <label>报表名称:</label>
          <input v-model="reportName" type="text" placeholder="可选，留空使用默认名称" />
        </div>
        <div class="form-actions">
          <button :disabled="!canGenerate || generating" class="primary" @click="generateReport">
            {{ generating ? '生成中...' : '生成报表' }}
          </button>
        </div>
      </div>
    </section>

    <!-- 生成结果 -->
    <section v-if="result" class="section">
      <h2>生成结果</h2>
      <div class="result" :class="{ success: result.success, error: !result.success }">
        <template v-if="result.success">
          <p><strong>✅ 生成成功</strong></p>
          <p>输出路径: {{ result.outputPath }}</p>
          <p>文件大小: {{ (result.size / 1024).toFixed(2) }} KB</p>
          <p>生成时间: {{ new Date(result.generatedAt).toLocaleString() }}</p>
          <p>耗时: {{ result.duration }}ms</p>
          <button class="primary" @click="openResultFolder">打开文件所在文件夹</button>
        </template>
        <template v-else>
          <p><strong>❌ 生成失败</strong></p>
          <p>{{ result.error }}</p>
        </template>
      </div>
    </section>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import { trpc } from '../utils/trpc'

// 状态
const templates = ref<any[]>([])
const selectedTemplateId = ref<string | null>(null)
const sourcePath = ref<string | null>(null)
const outputDir = ref<string | null>(null)
const reportName = ref('')
const result = ref<any>(null)
const loading = ref({
  templates: false
})
const generating = ref(false)

// 计算属性
const canGenerate = computed(() => {
  return selectedTemplateId.value && sourcePath.value && outputDir.value
})

// 加载模板列表
async function loadTemplates() {
  loading.value.templates = true
  try {
    templates.value = await trpc.template.list.query()
    console.log('[ReportTest] 模板列表:', templates.value)
  } catch (error) {
    console.error('[ReportTest] 加载模板失败:', error)
    alert('加载模板失败: ' + (error as any).message)
  } finally {
    loading.value.templates = false
  }
}

// 选择模板
function selectTemplate(id: string) {
  selectedTemplateId.value = id
  result.value = null
}

// 选择源文件
async function selectSourceFile() {
  try {
    const res = await trpc.file.selectSourceFile.query()
    if (!res.canceled && res.filePath) {
      sourcePath.value = res.filePath
    }
  } catch (error) {
    console.error('[ReportTest] 选择源文件失败:', error)
    alert('选择源文件失败: ' + (error as any).message)
  }
}

// 选择输出目录
async function selectOutputDir() {
  try {
    const res = await trpc.file.selectOutputDir.query()
    if (!res.canceled && res.dirPath) {
      outputDir.value = res.dirPath
    }
  } catch (error) {
    console.error('[ReportTest] 选择输出目录失败:', error)
    alert('选择输出目录失败: ' + (error as any).message)
  }
}

// 生成报表
async function generateReport() {
  if (!canGenerate.value) return

  generating.value = true
  result.value = null

  try {
    const res = await trpc.report.generate.mutate({
      templateId: selectedTemplateId.value!,
      sourcePath: sourcePath.value!,
      outputDir: outputDir.value!,
      reportName: reportName.value || undefined
    })

    console.log('[ReportTest] 生成结果:', res)
    result.value = { ...res, success: true }
  } catch (error) {
    console.error('[ReportTest] 生成报表失败:', error)
    result.value = {
      success: false,
      error: (error as any).message || '未知错误'
    }
  } finally {
    generating.value = false
  }
}

// 打开结果文件夹
async function openResultFolder() {
  if (!result.value?.outputPath) return

  try {
    await trpc.file.openInFolder.mutate({ path: result.value.outputPath })
  } catch (error) {
    console.error('[ReportTest] 打开文件夹失败:', error)
    alert('打开文件夹失败: ' + (error as any).message)
  }
}

// 页面加载时自动加载模板
loadTemplates()
</script>

<style scoped>
.report-test {
  padding: 20px;
  max-width: 1200px;
  margin: 0 auto;
}

h1 {
  font-size: 28px;
  margin-bottom: 30px;
  color: #333;
}

h2 {
  font-size: 20px;
  margin-bottom: 15px;
  color: #555;
}

.section {
  background: #f5f5f5;
  padding: 20px;
  border-radius: 8px;
  margin-bottom: 20px;
}

button {
  padding: 8px 16px;
  border: none;
  border-radius: 4px;
  background: #42b983;
  color: white;
  cursor: pointer;
  font-size: 14px;
}

button:hover:not(:disabled) {
  background: #359268;
}

button:disabled {
  background: #ccc;
  cursor: not-allowed;
}

button.primary {
  background: #409eff;
  padding: 10px 20px;
  font-size: 16px;
}

button.primary:hover:not(:disabled) {
  background: #66b1ff;
}

.template-list {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 15px;
  margin-top: 15px;
}

.template-item {
  background: white;
  padding: 15px;
  border-radius: 6px;
  border: 2px solid transparent;
  cursor: pointer;
  transition: all 0.2s;
}

.template-item:hover {
  border-color: #42b983;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.template-item.active {
  border-color: #409eff;
  background: #ecf5ff;
}

.template-item h3 {
  margin: 0 0 10px 0;
  font-size: 16px;
  color: #333;
}

.template-item p {
  margin: 5px 0;
  font-size: 13px;
  color: #666;
}

.empty {
  color: #999;
  text-align: center;
  padding: 20px;
}

.form {
  background: white;
  padding: 20px;
  border-radius: 6px;
}

.form-item {
  margin-bottom: 15px;
  display: flex;
  align-items: center;
  gap: 10px;
}

.form-item label {
  width: 100px;
  font-weight: 500;
  color: #555;
}

.form-item input {
  flex: 1;
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px;
}

.form-item input:disabled {
  background: #f5f5f5;
  color: #999;
}

.form-actions {
  margin-top: 20px;
  text-align: center;
}

.result {
  background: white;
  padding: 20px;
  border-radius: 6px;
  border-left: 4px solid #67c23a;
}

.result.error {
  border-left-color: #f56c6c;
}

.result p {
  margin: 10px 0;
  font-size: 14px;
  color: #333;
}

.result strong {
  font-size: 16px;
}

.result button {
  margin-top: 15px;
}
</style>
