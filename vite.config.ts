import { defineConfig } from 'vite';
import { resolve } from 'path';
import fs from 'fs';
import path from 'path';
import { homedir } from 'os';

function getOfficeAddinCertificates() {
  try {
    const certDir = path.join(homedir(), '.office-addin-dev-certs');
    const certPath = path.join(certDir, 'localhost.crt');
    const keyPath = path.join(certDir, 'localhost.key');
    if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
      return { cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) };
    }
  } catch (error) {
    console.warn('Could not load Office Add-in certificates:', error);
  }
  return undefined;
}

export default defineConfig({
  root: '.',
  server: {
    port: 3000,
    https: getOfficeAddinCertificates() || {},
    strictPort: true,
    cors: true
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: resolve(__dirname, 'index.html'),
      output: {
        entryFileNames: 'taskpane.js',
        chunkFileNames: '[name].js',
        assetFileNames: '[name].[ext]'
      }
    },
    target: 'es2015'
  },
  esbuild: {
    target: 'es2015'
  }
});
