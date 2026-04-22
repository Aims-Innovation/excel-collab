import { defineConfig } from 'vite';
import dts from 'vite-plugin-dts';

export default defineConfig({
  plugins: [dts()],
  build: {
    sourcemap: true,
    outDir: 'lib',
    lib: {
      entry: './src/index.ts',
      cssFileName: 'style',
      name: 'Excel',
      formats: ['es', 'umd'],
      fileName: (format, entryName) => {
        return `${entryName}.${format}.js`;
      },
    },

    rollupOptions: {
      external: ['react', 'react-dom'],
      output: {
        globals: {
          react: 'React',
          'react-dom': 'ReactDOM',
        },
      },
    },
  },
});
