import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// GitHub Pages serves a project repo at /<repo-name>/, not /, so asset
// paths need that prefix in production. The deploy workflow sets
// BASE_PATH to the real repo name at build time - locally (no env var
// set) this defaults to / so `npm run dev` is unaffected.
// https://vite.dev/config/
export default defineConfig({
  base: process.env.BASE_PATH || '/',
  plugins: [react()],
})
