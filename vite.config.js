export default {
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    watch: {
      ignored: ['**/src-tauri/**'],
    },
  },
  envPrefix: ['VITE_', 'TAURI_'],
  optimizeDeps: {
    include: ['exceljs'],
  },
  build: {
    target: 'es2021',
    minify: 'esbuild',
    sourcemap: false,
    rollupOptions: {
      output: {
        manualChunks: {
          exceljs: ['exceljs'],
        },
      },
    },
  },
};
