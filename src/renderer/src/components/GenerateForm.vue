<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { ElMessage } from 'element-plus'
import { useTemplates } from '../composables/useTemplates'
import { useReport } from '../composables/useReport'
import { trpc } from '../utils/trpc'
import TemplateInputForm from './TemplateInputForm.vue'
import FilePathField from './FilePathField.vue'
import ResultCard from './ResultCard.vue'
import { Checked } from '@element-plus/icons-vue'

const { currentTemplateId, currentTemplateName, currentTemplateMeta } = useTemplates()
const {
  sourcePath,
  outputDir,
  reportName,
  generating,
  result,
  setSourcePath,
  setOutputDir,
  setUserInput,
  extraSources,
  setExtraSource,
  resetReportState,
  generate
} = useReport()

// 单独查询当前模板的 inputRule（用于判断是否需要参数）
const currentInputRule = ref<any>(null)
const loadingInputRule = ref(false)

watch(
  () => currentTemplateId.value,
  async (templateId, oldTemplateId) => {
    // 切换模板时重置状态（首次加载时 oldTemplateId 为 undefined，不重置）
    if (oldTemplateId !== undefined) {
      resetReportState()
    }

    if (!templateId) {
      currentInputRule.value = null
      return
    }
    try {
      loadingInputRule.value = true
      const result = await trpc.template.getInputRule.query({ templateId })
      currentInputRule.value = result?.inputRule || null
    } catch (err) {
      console.error('[GenerateForm] 加载 inputRule 失败:', err)
      currentInputRule.value = null
    } finally {
      loadingInputRule.value = false
    }
  },
  { immediate: true }
)

// 由模板元信息决定是否需要额外的数据源输入（例如“放款明细表”）
const extraSourceRequirements = computed(() => currentTemplateMeta.value?.extraSources ?? [])

const primarySourceLabel = computed(() => currentTemplateMeta.value?.sourceLabel || '源数据文件')

// 所有必填的额外数据源必须选择完毕才能发起生成
const requiredExtraReady = computed(() => {
  if (!extraSourceRequirements.value.length) return true
  return extraSourceRequirements.value.every((req) => {
    if (req.required === false) return true
    return !!extraSources.value[req.id]
  })
})

const canGenerate = computed(
  () =>
    !!(currentTemplateId.value && sourcePath.value && outputDir.value && requiredExtraReady.value)
)
const templateFormRef = ref<InstanceType<typeof TemplateInputForm>>()

// 各区块完成状态
const sectionCompleted = computed(() => ({
  files: !!(sourcePath.value && outputDir.value && requiredExtraReady.value),
  // 如果当前模板需要 inputRule，则需要表单验证通过；否则默认完成
  params: !currentInputRule.value
}))

// 进度计算
const progressPercent = computed(() => {
  let completed = 0
  if (sectionCompleted.value.params) completed++
  if (sectionCompleted.value.files) completed++
  const total = 2
  return Math.round((completed / total) * 100)
})

const progressText = computed(() => {
  const params = sectionCompleted.value.params ? '✓' : '○'
  const files = sectionCompleted.value.files ? '✓' : '○'
  return `${params} 模板参数  ${files} 文件配置`
})

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
}
</script>

<template>
  <div class="generate-form">
    <!-- 头部：模板信息 + 报表名称 -->
    <div v-if="currentTemplateId" class="form-header">
      <div class="header-left">
        <div class="header-badge">
          <span class="badge-emoji">📄</span>
          <span class="badge-text">当前模板</span>
        </div>
        <h2 class="header-title">{{ currentTemplateName }}</h2>
      </div>
      <div class="header-right">
        <el-input
          v-model="reportName"
          placeholder="报表名称（可选）"
          clearable
          size="default"
          class="report-name-input"
        >
          <template #prefix>
            <span class="input-prefix">📝</span>
          </template>
        </el-input>
      </div>
    </div>

    <!-- 表单填写进度 -->
    <div v-if="currentTemplateId" class="form-progress">
      <div class="progress-bar">
        <div class="progress-fill" :style="{ width: progressPercent + '%' }" />
      </div>
      <span class="progress-text">{{ progressText }}</span>
    </div>

    <!-- 配置表单 -->
    <div v-if="currentTemplateId" class="form-sections">
      <!-- 步骤1：模板参数 -->
      <div
        class="form-section"
        :class="{ completed: sectionCompleted.params, 'no-params': sectionCompleted.params }"
      >
        <div class="section-header">
          <div
            class="step-number"
            :class="{ active: !sectionCompleted.params, completed: sectionCompleted.params }"
          >
            <span v-if="!sectionCompleted.params">1</span>
            <el-icon v-else><Checked /></el-icon>
          </div>
          <div class="section-info">
            <h3 class="section-title">模板参数</h3>
            <p class="section-subtitle">
              {{ sectionCompleted.params ? '无需额外参数' : '填写报表所需的额外参数' }}
            </p>
          </div>
        </div>
        <div v-if="!sectionCompleted.params" class="section-content">
          <TemplateInputForm
            ref="templateFormRef"
            :template-id="currentTemplateId || ''"
            @change="onParamsChange"
            @ready="onParamsReady"
          />
        </div>
      </div>

      <!-- 步骤2：文件配置 -->
      <div class="form-section" :class="{ completed: sectionCompleted.files }">
        <div class="section-header">
          <div
            class="step-number"
            :class="{ active: !sectionCompleted.files, completed: sectionCompleted.files }"
          >
            <span v-if="!sectionCompleted.files">2</span>
            <el-icon v-else><Checked /></el-icon>
          </div>
          <div class="section-info">
            <h3 class="section-title">文件配置</h3>
            <p class="section-subtitle">选择数据文件和输出位置</p>
          </div>
          <el-tag size="small" type="danger" effect="plain">必填</el-tag>
        </div>
        <div class="section-content">
          <FilePathField
            :label="primarySourceLabel"
            :model-value="sourcePath"
            type="file"
            :disabled="generating"
            :required="true"
            @update:model-value="setSourcePath"
          />
          <p v-if="currentTemplateMeta?.sourceDescription" class="field-hint">
            {{ currentTemplateMeta.sourceDescription }}
          </p>
          <div v-for="extra in extraSourceRequirements" :key="extra.id" class="extra-source">
            <FilePathField
              :label="extra.label"
              :model-value="extraSources[extra.id] || null"
              type="file"
              :disabled="generating"
              :required="extra.required !== false"
              @update:model-value="(val) => setExtraSource(extra.id, val)"
            />
            <p v-if="extra.description" class="field-hint">{{ extra.description }}</p>
          </div>
          <FilePathField
            label="输出目录"
            :model-value="outputDir"
            type="dir"
            :disabled="generating"
            :required="true"
            @update:model-value="setOutputDir"
          />
        </div>
      </div>

      <!-- 固定生成按钮 -->
      <div class="form-actions-fixed">
        <el-button
          class="generate-button"
          type="primary"
          size="large"
          :loading="generating"
          :disabled="!canGenerate || generating"
          @click="submit"
        >
          <el-icon v-if="!generating" class="button-icon"><Checked /></el-icon>
          <span v-if="generating">生成中...</span>
          <span v-else>生成报表</span>
        </el-button>
      </div>

      <!-- 结果展示 -->
      <ResultCard :result="result" />
    </div>

    <!-- 空状态 -->
    <div v-else class="empty-state">
      <el-empty description="请先从左侧选择一个模板" :image-size="160" />
    </div>
  </div>
</template>

<style scoped>
.generate-form {
  display: flex;
  flex-direction: column;
  gap: 16px;
  padding-bottom: 80px; /* 为固定按钮留空间 */
}

.field-hint {
  margin: 4px 0 12px 2px;
  font-size: 13px;
  color: #6b7280;
}

/* 头部：模板信息 + 报表名称 */
.form-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 20px;
  padding: 16px 20px;
  background: #ffffff;
  border-radius: 10px;
  border: 1px solid #e5e7eb;
}

.header-left {
  display: flex;
  align-items: center;
  gap: 12px;
  flex: 1;
  min-width: 0;
}

.header-badge {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 6px 12px;
  background: linear-gradient(135deg, #eef2ff 0%, #e0e7ff 100%);
  border-radius: 6px;
  flex-shrink: 0;
}

.badge-emoji {
  font-size: 16px;
}

.badge-text {
  font-size: 12px;
  font-weight: 600;
  color: #4f46e5;
}

.header-title {
  font-size: 16px;
  font-weight: 700;
  color: #111827;
  margin: 0;
  letter-spacing: -0.01em;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.header-right {
  flex-shrink: 0;
  min-width: 280px;
}

.report-name-input {
  width: 100%;
}

.report-name-input :deep(.el-input__wrapper) {
  border-radius: 8px;
  border: 1.5px solid #e5e7eb;
  box-shadow: none;
  transition: all 0.3s;
}

.report-name-input :deep(.el-input__wrapper):hover {
  border-color: #d1d5db;
}

.report-name-input :deep(.el-input__wrapper.is-focus) {
  border-color: #4f46e5;
  box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.1);
}

.input-prefix {
  font-size: 16px;
  margin-right: 4px;
}

/* 进度条 */
.form-progress {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 16px;
  background: #ffffff;
  border-radius: 10px;
  border: 1px solid #e5e7eb;
}

.progress-bar {
  flex: 1;
  height: 6px;
  background: #f3f4f6;
  border-radius: 3px;
  overflow: hidden;
}

.progress-fill {
  height: 100%;
  background: linear-gradient(90deg, #4f46e5 0%, #6366f1 100%);
  transition: width 0.4s cubic-bezier(0.4, 0, 0.2, 1);
}

.progress-text {
  font-size: 12px;
  font-weight: 600;
  color: #6b7280;
  white-space: nowrap;
}

/* 表单区域 */
.form-sections {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.form-section {
  background: #ffffff;
  border-radius: 12px;
  padding: 20px;
  border: 1.5px solid #e5e7eb;
  transition: all 0.3s;
}

.form-section:hover {
  border-color: #d1d5db;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
}

.form-section.completed {
  border-color: #10b981;
  background: linear-gradient(135deg, #f0fdf4 0%, #ffffff 100%);
}

/* 无参数时的紧凑样式 */
.form-section.no-params {
  padding: 16px 20px;
}

.form-section.no-params .section-header {
  margin-bottom: 0;
  padding-bottom: 0;
  border-bottom: none;
}

.form-section.no-params .section-subtitle {
  color: #10b981;
  font-weight: 600;
}

.section-header {
  display: flex;
  align-items: flex-start;
  gap: 12px;
  margin-bottom: 16px;
  padding-bottom: 12px;
  border-bottom: 1px solid #f3f4f6;
}

/* 步骤编号 */
.step-number {
  flex-shrink: 0;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 8px;
  background: #f3f4f6;
  color: #9ca3af;
  font-size: 14px;
  font-weight: 700;
  transition: all 0.3s;
}

.step-number.active {
  background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
  color: white;
  box-shadow: 0 2px 8px rgba(79, 70, 229, 0.3);
}

.step-number.completed {
  background: linear-gradient(135deg, #10b981 0%, #059669 100%);
  color: white;
  box-shadow: 0 2px 8px rgba(16, 185, 129, 0.3);
}

.section-info {
  flex: 1;
  min-width: 0;
}

.section-title {
  font-size: 15px;
  font-weight: 700;
  color: #111827;
  margin: 0 0 4px 0;
  letter-spacing: -0.01em;
}

.section-subtitle {
  font-size: 13px;
  color: #6b7280;
  margin: 0;
  line-height: 1.4;
}

.section-content {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

/* 固定生成按钮 */
.form-actions-fixed {
  position: fixed;
  bottom: 20px;
  right: 20px;
  z-index: 100;
}

.generate-button {
  min-width: 160px;
  height: 48px;
  font-size: 15px;
  font-weight: 700;
  border-radius: 24px;
  background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
  border: none;
  box-shadow:
    0 4px 12px rgba(79, 70, 229, 0.3),
    0 8px 24px rgba(0, 0, 0, 0.15);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  letter-spacing: 0.02em;
}

.generate-button .button-icon {
  margin-right: 6px;
  font-size: 18px;
}

.generate-button:hover:not(:disabled) {
  transform: translateY(-2px) scale(1.05);
  box-shadow:
    0 6px 16px rgba(79, 70, 229, 0.4),
    0 12px 32px rgba(0, 0, 0, 0.2);
}

.generate-button:active:not(:disabled) {
  transform: translateY(0) scale(1.02);
  box-shadow:
    0 3px 8px rgba(79, 70, 229, 0.35),
    0 6px 16px rgba(0, 0, 0, 0.15);
}

.generate-button:disabled {
  background: #9ca3af;
  opacity: 0.5;
  cursor: not-allowed;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

/* 空状态 */
.empty-state {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 480px;
  background: linear-gradient(135deg, #ffffff 0%, #fafbfc 100%);
  border-radius: 16px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06);
  border: 2px dashed #e5e7eb;
}
</style>
