import 'element-plus/dist/index.css';
import { createApp } from 'vue';

import App from './App.vue';
import { i18n } from './i18n';

const app = createApp(App);
app.use(i18n);
app.mount('#app');

const loadingScreen = document.getElementById('loading-screen');
if (loadingScreen) {
  setTimeout(() => {
    loadingScreen.classList.add('hidden');
    setTimeout(() => {
      loadingScreen.remove();
    }, 500);
  }, 100);
}
