// vite.config.js
import { defineConfig } from 'vite';

export default defineConfig({
  base: './', // 保持相對路徑（這對 Netlify 很重要）
});
