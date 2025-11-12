import { resolve } from 'path'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  main: {
    resolve: {
      alias: {
        '@shared': resolve('src/shared')
      }
    },
    plugins: [externalizeDepsPlugin()]
  },
  preload: {
    resolve: {
      alias: {
        '@shared': resolve('src/shared')
      }
    },
    plugins: [externalizeDepsPlugin()]
  },
  renderer: {
    resolve: {
      alias: {
        '@renderer': resolve('src/renderer/src'),
        '@shared': resolve('src/shared'),
        // 修复 @form-create/element-ui 的 Vue 3 兼容性问题
        vue: resolve('node_modules/vue/dist/vue.esm-bundler.js')
      }
    },
    plugins: [vue()],
    define: {
      // 确保 Vue 功能标志正确设置
      __VUE_OPTIONS_API__: true,
      __VUE_PROD_DEVTOOLS__: false
    }
  }
})
