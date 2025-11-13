import { createRouter, createWebHashHistory, RouteRecordRaw } from 'vue-router'
import Workbench from '../views/Workbench.vue'

const routes: RouteRecordRaw[] = [{ path: '/', name: 'Workbench', component: Workbench }]

const router = createRouter({
  history: createWebHashHistory(),
  routes
})

export default router
