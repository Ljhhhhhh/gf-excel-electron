import './assets/main.css'

import { createApp } from 'vue'
import App from './App.vue'
import FormCreate from '@form-create/element-ui'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'

const app = createApp(App)

// 注册 Element Plus
app.use(ElementPlus)

// 注册 formCreate
app.use(FormCreate)

app.mount('#app')
