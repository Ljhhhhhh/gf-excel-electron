import { createApp } from 'vue'
import App from './App.vue'
import router from './router'
import FormCreate from '@form-create/element-ui'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import './assets/design-system.css'
import './assets/main.css'

const app = createApp(App)

// 注册 Element Plus
app.use(ElementPlus)

// 注册 formCreate
app.use(FormCreate)
app.use(router)

app.mount('#app')
