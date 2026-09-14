import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';

export default {
  input: {
    core: 'src/core/index.ts',
    pro: 'src/pro/index.ts'
  },
  output: [
    {
      dir: 'dist',
      format: 'es',
      entryFileNames: '[name].mjs',
      chunkFileNames: '[name]-[hash].mjs'
    },
    {
      dir: 'dist',
      format: 'cjs',
      entryFileNames: '[name].js',
      chunkFileNames: '[name]-[hash].js'
    }
  ],
  plugins: [
    typescript(),
    // In a real build, we would have a custom encryption plugin here before terser.
    terser({
      compress: {
        passes: 2,
        drop_console: true,
      },
      mangle: {
        toplevel: true,
      },
    })
  ]
};
