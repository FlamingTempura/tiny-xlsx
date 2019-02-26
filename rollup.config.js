import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';
import { terser } from "rollup-plugin-terser";

export default [
	{
		input: 'src/index.js',
		output: {
			file: 'tiny-xlsx.js',
			format: 'iife',
			name: 'TinyXLSX'
		},
		plugins: [
			resolve({
				main: true,
				browser: true
			}),
			commonjs({
				ignore: [ 'fs', 'stream' ]
			}),
			terser()
		]
	}
];
