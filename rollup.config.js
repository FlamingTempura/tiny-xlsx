import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';
import { terser } from "rollup-plugin-terser";
import alias from 'rollup-plugin-alias';
import path from 'path';

export default [
	{
		input: 'src/index.js',
		output: {
			file: 'tiny-xlsx.js',
			format: 'umd',
			name: 'TinyXLSX'
		},
		plugins: [
		alias({
            jszip: path.join(__dirname, './node_modules/jszip/dist/jszip.min.js')
        }),
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
