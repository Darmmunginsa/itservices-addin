import { defineConfig } from 'vite'
import { resolve } from 'path'

export default defineConfig({
  base: '/itservices-addin/',
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'index.html'),
        auth: resolve(__dirname, 'auth.html'),
        // PhishGuard = อีกหน้าหนึ่งใน build เดียวกัน (คนละ task pane แต่ deploy พร้อมกัน)
        phish: resolve(__dirname, 'phish.html'),
      },
    },
  },
})
