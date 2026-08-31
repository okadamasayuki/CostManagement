import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';

// 単一 HTML ファイルとして出力する。
// 社内 PC でファイルをダブルクリックするだけで動き、サーバも
// インストールも不要にするため。外部 CDN は一切参照しない。
export default defineConfig({
  plugins: [react(), viteSingleFile({ removeViteModuleLoader: true })],
  build: {
    target: 'es2020',
    assetsInlineLimit: 100000000,
    chunkSizeWarningLimit: 100000,
    cssCodeSplit: false,
    reportCompressedSize: false,
    rollupOptions: { output: { inlineDynamicImports: true } },
  },
});
