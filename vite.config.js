import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import electron from 'vite-plugin-electron';
// https://vite.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    electron({
      entry: {
        main: 'electron/main.js',
        preload: 'electron/preload.js',
      },
    }),
  ],
  // 性能优化配置
  build: {
    // 代码分割
    rollupOptions: {
      output: {
        manualChunks: {
          // 将大型依赖单独打包
          'element-plus': ['element-plus'],
          'xlsx': ['xlsx'],
          'docxtemplater': ['docxtemplater', 'pizzip'],
          'jszip': ['jszip', 'file-saver'],
        },
      },
    },
    // 减小打包体积
    minify: 'terser',
    terserOptions: {
      compress: {
        drop_console: true,
        drop_debugger: true,
        pure_funcs: ['console.log'],
      },
    },
    // 生成 sourcemap，便于调试
    sourcemap: false,
    // 启用 CSS 代码分割
    cssCodeSplit: true,
  },
  // 优化依赖预构建
  optimizeDeps: {
    include: [
      'vue',
      'element-plus',
      'xlsx',
      'docxtemplater',
      'pizzip',
      'jszip',
      'file-saver',
    ],
    exclude: ['electron'],
  },
  // 开发服务器优化
  server: {
    port: 5173,
    host: '0.0.0.0',
    open: false,
    // 增加服务器响应超时
    timeout: 60000,
  },
  // 预览服务器优化
  preview: {
    port: 5174,
    host: '0.0.0.0',
    open: false,
  },
});