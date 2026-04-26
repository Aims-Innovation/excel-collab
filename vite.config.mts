import { defineConfig } from 'vite';
import dts from 'vite-plugin-dts';

export default defineConfig({
  plugins: [dts()],
  build: {
    sourcemap: true,
    outDir: 'lib',
    // Don't base64-inline anything large into CSS / JS. With the
    // bundled fonts (Archivo + IBM Plex Sans Arabic woff2 -- ~50 KB
    // each before gzip) inlining would blow style.css from ~22 KB
    // to ~700 KB. Setting the limit to 0 forces every asset to be
    // emitted as a separate file in lib/ with rewritten URLs --
    // browsers cache them independently and parallel-download.
    assetsInlineLimit: 0,
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
      //
      // yjs is externalized so the consumer provides a single copy and
      // we don't ship the ~200 KB library twice. All our yjs imports are
      // bare 'yjs' (no subpaths in use; verified via `grep -r "from 'yjs/"`)
      // — the id === 'yjs' match is sufficient but we still allow subpaths
      // defensively for future-proofing.
      external: (id) =>
        id === 'react' ||
        id === 'react-dom' ||
        id.startsWith('react/') ||
        id.startsWith('react-dom/') ||
        id === 'yjs' ||
        id.startsWith('yjs/'),
      output: {
        globals: {
          react: 'React',
          'react-dom': 'ReactDOM',
          yjs: 'Y',
        },
      },
    },
  },
});
