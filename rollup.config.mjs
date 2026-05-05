import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';

const output = {
	banner: `/*
 * @license
 * xlsx-preview
 * Released under Apache License 2.0
 */`,
	sourcemap: true,
}

const umdOutput = {
	...output,
	name: "xlsx",
	file: 'dist/xlsx-preview.js',
	format: 'umd',
	globals: {
		jszip: 'JSZip'
	},
};

export default args => {
	const prod = args.environment == 'BUILD:production';

	// In prod, emit .d.ts alongside the bundle so package.json's
	// `"types": "dist/xlsx-preview.d.ts"` resolves. The dev build skips
	// declarations to keep `npm run build` fast.
	const tsPlugin = typescript(
		prod
			? {
					declaration: true,
					declarationDir: 'dist',
					rootDir: 'src',
					outDir: 'dist',
			  }
			: {}
	);

	const config = {
		input: 'src/xlsx-preview.ts',
		output: [umdOutput],
		plugins: [tsPlugin]
	}

	if (prod)
		config.output = [umdOutput,
			{
				...umdOutput,
				file: 'dist/xlsx-preview.min.js',
				plugins: [terser()]
			},
			{
				...output,
				file: 'dist/xlsx-preview.mjs',
				format: 'es',
			},
			{
				...output,
				file: 'dist/xlsx-preview.min.mjs',
				format: 'es',
				plugins: [terser()]
			}];

	return config
};
