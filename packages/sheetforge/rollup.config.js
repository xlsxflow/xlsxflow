import typescript from '@rollup/plugin-typescript';
import { dts } from 'rollup-plugin-dts';

// Node built-ins are only loaded lazily by createFileReader()
const external = ['fs', 'fs/promises', 'stream'];

export default [
  {
    input: 'src/index.ts',
    output: [
      { file: 'dist/index.mjs', format: 'es', sourcemap: true },
      { file: 'dist/index.cjs', format: 'cjs', sourcemap: true },
    ],
    external,
    plugins: [typescript({ declaration: false, declarationDir: undefined, exclude: ['test/**', 'scripts/**'] })],
  },
  {
    // Single bundled declaration file: no extensionless relative imports for nodenext consumers.
    input: 'src/index.ts',
    output: [
      { file: 'dist/index.d.ts', format: 'es' },
      { file: 'dist/index.d.cts', format: 'es' },
    ],
    external,
    plugins: [dts()],
  },
];
