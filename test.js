//const spawn = require('child_process').spawn;

const TinyXLSX = require('./src/index');
const tap = require('tap');
const XLSX = require('xlsx');
const fs = require('fs');

let data1 = [
	{ title: 'Test1' , data: [
		['foo', 'bar'], ['noo', 'bar'], [9],
		[1, 2, 3],
		[3, 4]] },
	{ title: 'Boo?' , data: [[1, 2], [3, 4]] },
	{ title: 'Another sheet', data: [[1, 2], ['Total', 4]]  }
];

let data2 = [
	{ title: 'blah', data: [[1, 2, 3]] }
];

let loadXLSX = () => {
	let workbook = XLSX.readFile('tmp.xlsx');
	return workbook.SheetNames.map(title => ({
		title,
		data: XLSX.utils.sheet_to_json(workbook.Sheets[title], { header: 1 })
	}));
};

tap.test('it should write a valid xlsx spreadsheet', async t => {
	for (let data of [data1, data2]) {
		await TinyXLSX.generate(data).write('tmp.xlsx');
		let worksheets = loadXLSX('tmp.xlsx');
		t.match(data, worksheets);
		fs.unlinkSync('tmp.xlsx');
	}
});

tap.test('it should stream a valid xlsx spreadsheet', async t => {
	for (let data of [data1, data2]) {
		await new Promise(resolve => {
			TinyXLSX.generate(data).stream()
				.pipe(fs.createWriteStream('tmp.xlsx'))
				.on('finish', () => resolve());
		});
		let worksheets = loadXLSX('tmp.xlsx');
		t.match(data, worksheets);
		fs.unlinkSync('tmp.xlsx');
	}
});

tap.test('it should generate a valid base64 spreadsheet', async t => {
	for (let data of [data1, data2]) {
		let base64 = await TinyXLSX.generate(data).base64();
		fs.writeFileSync('tmp.xlsx', Buffer.from(base64, 'base64'));
		let worksheets = loadXLSX('tmp.xlsx');
		t.match(data, worksheets);
		fs.unlinkSync('tmp.xlsx');
	}
});

// tap.test('it should generate a valid blob spreadsheet', async t => {
// 	for (let data of [data1, data2]) {
// 		await TinyXLSX.generate(data).blob();
// 		// let worksheets = loadXLSX('tmp.xlsx');
// 		// t.match(data, worksheets);
// 		// fs.unlinkSync('tmp.xlsx');
// 	}
// });
