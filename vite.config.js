import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import path from 'path'
path.resolve(__dirname, './src')

// https://vitejs.dev/config/
export default defineConfig({
  base:"./",
  plugins: [vue()],
  resolve: {
    alias: {
      // src 读取电脑src地址
      '@': path.resolve(__dirname, './src')
    }
  },
  css: {
    preprocessorOptions: {
      scss: {
        additionalData: `@import "@/style/minxin.scss";`,
      }
    },
  },
  build: {
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
      }
    }
  },
})
