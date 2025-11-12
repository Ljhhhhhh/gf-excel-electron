<script setup lang="ts">
import TemplateList from '../components/TemplateList.vue'
import GenerateForm from '../components/GenerateForm.vue'
</script>

<template>
  <div class="workbench">
    <div class="left">
      <TemplateList />
    </div>
    <div class="right">
      <GenerateForm />
    </div>
  </div>
</template>

<style scoped>
.workbench { display: grid; grid-template-columns: 320px 1fr; gap: 16px; padding: 12px; }
.left { height: calc(100vh - 24px); overflow: auto; }
.right { height: calc(100vh - 24px); overflow: auto; }
</style>

