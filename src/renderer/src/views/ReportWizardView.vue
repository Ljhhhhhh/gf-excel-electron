<script setup lang="ts">
import { ref, computed } from 'vue'
import { useTemplates } from '../composables/useTemplates'
import { useReport } from '../composables/useReport'
import TemplateInputForm from '../components/TemplateInputForm.vue'
import FilePathField from '../components/FilePathField.vue'
import ResultCard from '../components/ResultCard.vue'
import { ElMessage } from 'element-plus'

const { currentTemplateId, currentTemplateName } = useTemplates()
const {
  sourcePath,
  outputDir,
  reportName,
  generating,
  result,
  setSourcePath,
  setOutputDir,
  setUserInput,
  generate
} = useReport()

const active = ref(0)
const canGenerate = computed(
  () => !!(currentTemplateId.value && sourcePath.value && outputDir.value)
)
const templateFormRef = ref<InstanceType<typeof TemplateInputForm>>()

function onParamsChange(data: Record<string, any>) {
  setUserInput(data)
}

function onParamsReady(data: Record<string, any>) {
  setUserInput(data)
}

async function submit() {
  if (!canGenerate.value) {
    ElMessage.warning('请完善必填项')
    return
  }
  if (templateFormRef.value) {
    const ok = await templateFormRef.value.validate()
    if (!ok) {
      ElMessage.error('请完善模板参数')
      return
    }
  }
  await generate(currentTemplateId.value!)
  active.value = 3
}
</script>

<template>
  <div class="report-wizard-view">
    <div class="summary">
      <el-card shadow="never">
        <div class="summary-title">当前模板</div>
        <div class="summary-name">{{ currentTemplateName || '未选择' }}</div>
      </el-card>
    </div>
    <div class="wizard">
      <el-card shadow="never">
        <el-steps :active="active" finish-status="success" class="steps">
          <el-step title="选择源文件" />
          <el-step title="选择输出目录" />
          <el-step title="填写参数" />
          <el-step title="生成结果" />
        </el-steps>

        <el-form label-width="100px">
          <div v-show="active === 0">
            <FilePathField
              label="源文件"
              :model-value="sourcePath"
              type="file"
              :disabled="generating"
              @update:model-value="setSourcePath"
            />
            <div class="form-actions">
              <el-button type="primary" :disabled="!sourcePath" @click="active = 1"
                >下一步</el-button
              >
            </div>
          </div>
          <div v-show="active === 1">
            <FilePathField
              label="输出目录"
              :model-value="outputDir"
              type="dir"
              :disabled="generating"
              @update:model-value="setOutputDir"
            />
            <el-form-item label="报表名称"
              ><el-input v-model="reportName" placeholder="可选，留空使用默认名称"
            /></el-form-item>
            <div class="form-actions">
              <el-button @click="active = 0">上一步</el-button
              ><el-button type="primary" :disabled="!outputDir" @click="active = 2"
                >下一步</el-button
              >
            </div>
          </div>
          <div v-show="active === 2">
            <el-form-item label="模板参数">
              <TemplateInputForm
                ref="templateFormRef"
                :template-id="currentTemplateId || ''"
                @change="onParamsChange"
                @ready="onParamsReady"
              />
            </el-form-item>
            <div class="form-actions">
              <el-button @click="active = 1">上一步</el-button>
              <el-button
                type="primary"
                :loading="generating"
                :disabled="!canGenerate || generating"
                @click="submit"
                >生成报表</el-button
              >
            </div>
          </div>
          <div v-show="active === 3">
            <ResultCard :result="result" />
            <div class="form-actions"><el-button @click="active = 0">重新生成</el-button></div>
          </div>
        </el-form>
      </el-card>
    </div>
  </div>
</template>

<style scoped>
.report-wizard-view {
  display: grid;
  grid-template-columns: 280px 1fr;
  gap: 16px;
  padding: 8px;
}
.summary-title {
  font-size: 12px;
  color: #909399;
}
.summary-name {
  margin-top: 6px;
  font-weight: 600;
}
.steps {
  margin: 12px 0;
}
.form-actions {
  margin-top: 12px;
  display: flex;
  gap: 8px;
}
</style>
