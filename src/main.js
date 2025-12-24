import 'element-plus/dist/index.css';
import { createApp } from 'vue';

import App from './App.vue';

const app = createApp(App);
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
