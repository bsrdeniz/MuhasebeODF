import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
// Test push for Railway deployment
export default defineConfig({
  plugins: [react()],
  preview: {
    allowedHosts: true
  }
})
