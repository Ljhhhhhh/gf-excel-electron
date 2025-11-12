<script setup lang="ts">
import { computed, ref } from 'vue'
import { ElMessage } from 'element-plus'
import { useTemplates } from '../composables/useTemplates'
import { useReport } from '../composables/useReport'
import TemplateInputForm from './TemplateInputForm.vue'
import FilePathField from './FilePathField.vue'
import ResultCard from './ResultCard.vue'

const { currentTemplateId, currentTemplateName } = useTemplates()
const {
  sourcePath,
  outputDir,
  reportName,
  userInput,
  generating,
  result,
  setSourcePath,
  setOutputDir,
  setUserInput,
  generate
} = useReport()

const canGenerate = computed(() => !!(currentTemplateId.value && sourcePath.value && outputDir.value))
const templateFormRef = ref<InstanceType<typeof TemplateInputForm>>()

function onParamsChange(data: Record<string, any>) { setUserInput(data) }
function onParamsReady(data: Record<string, any>) { setUserInput(data) }

async function submit() {
  if (!canGenerate.value) { ElMessage.warning('请完善必填项'); return }
  if (templateFormRef.value) {
    const ok = await templateFormRef.value.validate()
    if (!ok) { ElMessage.error('请完善模板参数'); return }
  }
  await generate(currentTemplateId.value!)
}
</script>

<template>
  <el-card class="generate-form" shadow="never">
    <div class="header">
      <div class="title">{{ currentTemplateName || '请选择左侧模板' }}</div>
    </div>
    <el-form label-width="100px">
      <FilePathField label="源文件" :model-value="sourcePath" type="file" :disabled="generating" @update:modelValue="setSourcePath" />
      <FilePathField label="输出目录" :model-value="outputDir" type="dir" :disabled="generating" @update:modelValue="setOutputDir" />
      <el-form-item label="报表名称"><el-input v-model="reportName" placeholder="可选，留空使用默认名称" /></el-form-item>
      <el-form-item label="模板参数">
        <TemplateInputForm ref="templateFormRef" :template-id="currentTemplateId || ''" @change="onParamsChange" @ready="onParamsReady" />
      </el-form-item>
      <div class="actions">
        <el-button type="primary" :loading="generating" :disabled="!canGenerate || generating" @click="submit">生成报表</el-button>
      </div>
      <ResultCard :result="result" />
    </el-form>
  </el-card>
 </template>

<style scoped>
.generate-form { border: 1px solid rgba(255,255,255,0.08); }
.header { margin-bottom: 8px; }
.title { font-weight: 600; }
.actions { margin-top: 12px; }
</style>

