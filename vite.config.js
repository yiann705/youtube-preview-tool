// vite.config.js
import { defineConfig } from 'vite';
import { execSync } from 'child_process';

const commitHash = execSync('git rev-parse --short HEAD').toString().trim();
const buildTime = new Date().toLocaleString('zh-TW', {
  timeZone: 'Asia/Taipei',
  hour12: false,
});

export default defineConfig({
  base: './', // 保持相對路徑，適用 Netlify
  define: {
    __APP_VERSION__: JSON.stringify(`v${commitHash} - ${buildTime}`)
  }
});
