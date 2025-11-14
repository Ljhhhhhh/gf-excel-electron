<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { ElMessage } from 'element-plus'
import { useTemplates } from '../composables/useTemplates'
import { useReport } from '../composables/useReport'
import TemplateInputForm from './TemplateInputForm.vue'
import FilePathField from './FilePathField.vue'
import ResultCard from './ResultCard.vue'
import { Document, Folder, Edit, Checked } from '@element-plus/icons-vue'

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
  resetExtraSources,
  generate
} = useReport()

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

function onParamsChange(data: Record<string, any>) {
  setUserInput(data)
}
function onParamsReady(data: Record<string, any>) {
  setUserInput(data)
}

watch(
  () => currentTemplateId.value,
  () => {
    resetExtraSources()
  }
)

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
    <!-- 头部区域 -->
    <div class="form-header">
      <div class="header-icon">
        <el-icon :size="28"><Document /></el-icon>
      </div>
      <div class="header-content">
        <h2 class="header-title">{{ currentTemplateName || '请先选择模板' }}</h2>
        <p v-if="currentTemplateId" class="header-subtitle">配置参数并生成报表</p>
        <p v-else class="header-subtitle">从左侧列表中选择一个报表模板开始</p>
      </div>
    </div>

    <!-- 配置表单 -->
    <div v-if="currentTemplateId" class="form-sections">
      <!-- 文件配置 -->
      <div class="form-section">
        <div class="section-header">
          <el-icon class="section-icon" :size="28"><Folder /></el-icon>
          <h3 class="section-title">文件配置</h3>
          <el-tag size="small" type="danger">必填</el-tag>
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

      <!-- 报表配置 -->
      <div class="form-section">
        <div class="section-header">
          <el-icon class="section-icon" :size="28"><Edit /></el-icon>
          <h3 class="section-title">报表配置</h3>
          <el-tag size="small">可选</el-tag>
        </div>
        <div class="section-content">
          <el-form label-width="100px" label-position="left">
            <el-form-item label="报表名称">
              <el-input v-model="reportName" placeholder="留空则使用默认名称" clearable />
            </el-form-item>
          </el-form>
        </div>
      </div>

      <!-- 模板参数 -->
      <div class="form-section">
        <div class="section-header">
          <el-icon class="section-icon" :size="28"><Checked /></el-icon>
          <h3 class="section-title">模板参数</h3>
        </div>
        <div class="section-content">
          <TemplateInputForm
            ref="templateFormRef"
            :template-id="currentTemplateId || ''"
            @change="onParamsChange"
            @ready="onParamsReady"
          />
        </div>
      </div>

      <!-- 操作按钮 -->
      <div class="form-actions">
        <el-button
          type="primary"
          size="large"
          :loading="generating"
          :disabled="!canGenerate || generating"
          @click="submit"
        >
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
  gap: 24px;
}

.field-hint {
  margin: 4px 0 12px 2px;
  font-size: 13px;
  color: #6b7280;
}

/* 头部区域 */
.form-header {
  display: flex;
  align-items: center;
  gap: 20px;
  padding: 28px 32px;
  background: linear-gradient(135deg, #ffffff 0%, #fafbfc 100%);
  border-radius: 16px;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
  border: 2px solid #f3f4f6;
  transition: all 0.3s;
}

.form-header:hover {
  box-shadow: 0 6px 28px rgba(79, 70, 229, 0.12);
  border-color: #e5e7eb;
}

.header-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 64px;
  height: 64px;
  background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
  color: white;
  border-radius: 16px;
  box-shadow: 0 8px 20px rgba(79, 70, 229, 0.35);
  transition: all 0.3s;
}

.form-header:hover .header-icon {
  transform: scale(1.05) rotate(2deg);
  box-shadow: 0 10px 28px rgba(79, 70, 229, 0.45);
}

.header-content {
  flex: 1;
}

.header-title {
  font-size: 22px;
  font-weight: 700;
  color: #111827;
  margin: 0 0 6px 0;
  letter-spacing: -0.02em;
}

.header-subtitle {
  font-size: 14px;
  color: #6b7280;
  margin: 0;
  line-height: 1.5;
}

/* 表单区域 */
.form-sections {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.form-section {
  background: linear-gradient(135deg, #ffffff 0%, #fafbfc 100%);
  border-radius: 16px;
  padding: 28px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06);
  border: 2px solid #f3f4f6;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  position: relative;
  overflow: hidden;
}

.form-section::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  width: 4px;
  height: 100%;
  background: linear-gradient(180deg, #4f46e5 0%, #6366f1 100%);
  opacity: 0;
  transition: opacity 0.3s;
}

.form-section:hover {
  box-shadow: 0 6px 24px rgba(79, 70, 229, 0.12);
  border-color: #e5e7eb;
  transform: translateY(-2px);
}

.form-section:hover::before {
  opacity: 1;
}

.section-header {
  display: flex;
  align-items: center;
  gap: 14px;
  margin-bottom: 20px;
  padding-bottom: 16px;
  border-bottom: 2px solid #e5e7eb;
}

.section-icon {
  color: #4f46e5;
  filter: drop-shadow(0 2px 4px rgba(79, 70, 229, 0.2));
}

.section-title {
  flex: 1;
  font-size: 17px;
  font-weight: 700;
  color: #111827;
  margin: 0;
  letter-spacing: -0.01em;
}

.section-content {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

/* 操作按钮 */
.form-actions {
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 32px;
  background: linear-gradient(135deg, #fafbfc 0%, #ffffff 100%);
  border-radius: 16px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06);
  border: 2px solid #f3f4f6;
  position: relative;
  overflow: hidden;
}

.form-actions::before {
  content: '';
  position: absolute;
  inset: 0;
  background: linear-gradient(135deg, rgba(79, 70, 229, 0.03) 0%, rgba(99, 102, 241, 0.03) 100%);
  opacity: 0;
  transition: opacity 0.3s;
}

.form-actions:hover::before {
  opacity: 1;
}

.form-actions .el-button {
  min-width: 240px;
  font-size: 17px;
  font-weight: 700;
  height: 56px;
  border-radius: 12px;
  background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
  border: none;
  box-shadow: 0 8px 20px rgba(79, 70, 229, 0.35);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  position: relative;
  z-index: 1;
  letter-spacing: 0.02em;
}

.form-actions .el-button:hover:not(:disabled) {
  transform: translateY(-3px) scale(1.02);
  box-shadow: 0 12px 28px rgba(79, 70, 229, 0.45);
}

.form-actions .el-button:active:not(:disabled) {
  transform: translateY(-1px) scale(1.01);
  box-shadow: 0 6px 16px rgba(79, 70, 229, 0.4);
}

.form-actions .el-button:disabled {
  background: linear-gradient(135deg, #9ca3af 0%, #6b7280 100%);
  opacity: 0.6;
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
