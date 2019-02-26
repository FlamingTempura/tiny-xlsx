const TinyXLSX = require('..');

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
];

TinyXLSX.generate(sheets)
	.write('accounts.xlsx')
	.then(() => console.log('done!'));