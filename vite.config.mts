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
      // External must cover subpath imports like `react-dom/client`,
      // `react/jsx-runtime`, etc. — otherwise rollup bundles the whole
      // react-dom package and ships its "Incompatible React versions"
      // check compiled with hardcoded versions from the build machine.
      external: (id) =>
        id === 'react' ||
        id === 'react-dom' ||
        id.startsWith('react/') ||
        id.startsWith('react-dom/'),
      output: {
        globals: {
          react: 'React',
          'react-dom': 'ReactDOM',
        },
      },
    },
  },
});
