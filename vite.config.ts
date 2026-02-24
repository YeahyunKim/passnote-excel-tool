import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  // GitHub Pages(프로젝트 페이지)에서도 경로가 깨지지 않도록 상대 경로로 빌드합니다.
  base: './',
});
