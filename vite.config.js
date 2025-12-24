import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import electron from 'vite-plugin-electron';
import AutoImport from 'unplugin-auto-import/vite';
import Components from 'unplugin-vue-components/vite';
import { ElementPlusResolver } from 'unplugin-vue-components/resolvers';
// https://vite.dev/config/
export default defineConfig({
  plugins: [
    vue(),
    AutoImport({
      resolvers: [ElementPlusResolver()],
    }),
    Components({
      resolvers: [ElementPlusResolver()],
    }),
    electron({
      entry: {
        main: 'electron/main.js',
        preload: 'electron/preload.js',
      },
    }),
  ],
  build: {
    target: 'esnext',
    rollupOptions: {
      output: {
        manualChunks: {
          'element-plus': ['element-plus'],
          xlsx: ['xlsx'],
          docxtemplater: ['docxtemplater', 'pizzip'],
          jszip: ['jszip', 'file-saver'],
        },
      },
    },
    minify: 'terser',
    terserOptions: {
      compress: {
        drop_console: true,
        drop_debugger: true,
        pure_funcs: ['console.log'],
      },
    },
    cssCodeSplit: true,
    chunkSizeWarningLimit: 1000,
  },
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
  server: {
    port: 5173,
    host: '0.0.0.0',
    open: false,
    timeout: 60000,
  },
});
