<script setup lang="ts">
import TemplateList from '../components/TemplateList.vue'
import GenerateForm from '../components/GenerateForm.vue'
</script>

<template>
  <div class="workbench">
    <div class="left-panel">
      <TemplateList />
    </div>
    <div class="right-panel">
      <GenerateForm />
    </div>
  </div>
</template>

<style scoped>
.workbench {
  display: grid;
  grid-template-columns: 360px 1fr;
  gap: 32px;
  padding: 32px;
  min-height: 100vh;
  background: linear-gradient(135deg, #f8fafc 0%, #eef2f7 50%, #e5e9f0 100%);
  position: relative;
}

/* 添加背景装饰 */
.workbench::before {
  content: '';
  position: fixed;
  top: -50%;
  left: -50%;
  width: 200%;
  height: 200%;
  background: radial-gradient(circle at 30% 20%, rgba(79, 70, 229, 0.08) 0%, transparent 50%);
  pointer-events: none;
  z-index: 0;
}

.workbench::after {
  content: '';
  position: fixed;
  bottom: -50%;
  right: -50%;
  width: 200%;
  height: 200%;
  background: radial-gradient(circle at 70% 80%, rgba(99, 102, 241, 0.06) 0%, transparent 50%);
  pointer-events: none;
  z-index: 0;
}

.left-panel {
  height: calc(100vh - 64px);
  overflow: hidden;
  display: flex;
  flex-direction: column;
  position: relative;
  z-index: 1;
  animation: slideInLeft 0.4s cubic-bezier(0.4, 0, 0.2, 1);
}

@keyframes slideInLeft {
  from {
    opacity: 0;
    transform: translateX(-20px);
  }
  to {
    opacity: 1;
    transform: translateX(0);
  }
}

.right-panel {
  height: calc(100vh - 64px);
  overflow-y: auto;
  overflow-x: hidden;
  position: relative;
  z-index: 1;
  animation: slideInRight 0.4s cubic-bezier(0.4, 0, 0.2, 1);
}

@keyframes slideInRight {
  from {
    opacity: 0;
    transform: translateX(20px);
  }
  to {
    opacity: 1;
    transform: translateX(0);
  }
}

/* 自定义滚动条样式 - 更现代化 */
.right-panel::-webkit-scrollbar {
  width: 10px;
}

.right-panel::-webkit-scrollbar-track {
  background: rgba(0, 0, 0, 0.02);
  border-radius: 10px;
}

.right-panel::-webkit-scrollbar-thumb {
  background: linear-gradient(180deg, #4f46e5 0%, #6366f1 100%);
  border-radius: 10px;
  border: 2px solid transparent;
  background-clip: padding-box;
  transition: all 0.3s;
}

.right-panel::-webkit-scrollbar-thumb:hover {
  background: linear-gradient(180deg, #4338ca 0%, #4f46e5 100%);
  background-clip: padding-box;
}

/* 响应式调整 */
@media (max-width: 1280px) {
  .workbench {
    grid-template-columns: 300px 1fr;
    gap: 24px;
    padding: 24px;
  }
}
</style>
