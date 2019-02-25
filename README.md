## tiny-xlsx

Tiny JavaScript XLSX writer.


### Node.js

npm install tiny-xlsx

```js
import TinyXLSX from 'tiny-xlsx';

let summary = [
	['Annual Sales',  40000 ],
	['Costs',         30000 ],
	['Profit',        10000 ]
];

let transactions = [
	['Date',       'Description',       'Amount'],
	['2017-01-01', 'Sales',               200   ],
	['2017-01-01', 'Sales',               -20   ],
	['2017-01-01', 'Sales',              1200   ],
	['2017-01-01', 'Sales',               -10.5 ]
];

let sheets = [
	{ title: 'Summary', data: summary },
	{ title: 'Transactions', data: transactions }
]

TinyXLSX.generate(sheets)
	.write('accounts.xlsx')
	.then(() => console.log('done!'))
```

### Browser


```html
<script script="tiny-xlsx"></script>

<script>

let data = [
	[1, 2]
	['hello', 'world']
];

let sheets = { title: 'Hello World', data };

TinyXLSX.generate(sheets)
	.base64()
	.then(base64 => {
		location.href = "data:application/zip;base64," + base64;
	});

</script>
```

### API

#### TinyXLSX.generate(sheets) ==> workbook

Generates XLSX file for given array of sheets.

#### workbook.write(filename) ==> Promise [Node only]

Writes workbook to given filename.

#### workbook.stream() ==> stream [Node only]

Provides a stream.

#### workbook.base64() ==> Promise

Provides a base64 string.

#### workbook.blob() ==> Promise

Provides a blob.