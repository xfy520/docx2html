import { defineConfig, Options } from 'tsup';
import { BuildOptions } from 'esbuild';

export default defineConfig((options): Options => ({
  entry: ['src/index.ts'],
  splitting: false,
  sourcemap: false,
  clean: true,
  outExtension({ format }) {
    return {
      js: `.${format}.js`,
    };
  },
  esbuildOptions(opts: BuildOptions) {
    if (!options.watch) {
      opts.external = ['jszip'];
    }
  },
  watch: options.watch,
  minify: !options.watch,
}));
