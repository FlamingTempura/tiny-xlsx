import JSZip from 'jszip';

let isNode = typeof module !== 'undefined' && module.exports,
	fs;
if (isNode) {
	fs = require('fs');
}

const RELS_XML = () => `<?xml version="1.0" encoding="UTF-8"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
	</Relationships>`;

const APP_XML = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
		        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
		<Template></Template>
		<TotalTime>0</TotalTime>
		<Application>LifeCourse$Build-1</Application>
	</Properties>`;

const CORE_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<cp:coreProperties
			xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
			xmlns:dc="http://purl.org/dc/elements/1.1/" 
			xmlns:dcterms="http://purl.org/dc/terms/" 
			xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
			xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
		<dcterms:created xsi:type="dcterms:W3CDTF">${workbook.isoDate}</dcterms:created>
		<dc:creator></dc:creator>
		<dc:description></dc:description>
		<dc:language>en-GB</dc:language>
		<cp:lastModifiedBy></cp:lastModifiedBy>
		<dcterms:modified xsi:type="dcterms:W3CDTF">${workbook.isoDate}</dcterms:modified>
		<cp:revision>1</cp:revision>
		<dc:subject></dc:subject>
		<dc:title></dc:title>
	</cp:coreProperties>`;

const WORKBOOK_XML_RELS = workbook => `<?xml version="1.0" encoding="UTF-8"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
		${workbook.sheets.map(sheet => `
		<Relationship Id="${sheet.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheet.id}.xml"/>
		`).join('\n')}
	</Relationships>`;

const SHEET_XML = sheet => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<worksheet
		xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<sheetPr filterMode="false">
			<pageSetUpPr fitToPage="false"/>
		</sheetPr>
		<dimension ref="A1:${sheet.extent}"/>
		<sheetViews>
			<sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" 
					rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" 
					view="normal" topLeftCell="A1" colorId="64" zoomScale="79" zoomScaleNormal="79" 
					zoomScalePageLayoutView="100" workbookViewId="0">
				<selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/>
			</sheetView>
		</sheetViews>
		<sheetFormatPr defaultRowHeight="12.8" zeroHeight="false" outlineLevelRow="0" outlineLevelCol="0"></sheetFormatPr>
		<cols>
			<col collapsed="false" customWidth="false" hidden="false" outlineLevel="0" max="1025" min="1" style="0" width="11.52"/>
		</cols>
		<sheetData>
			${sheet.rows.map(row => `
			<row r="${row.id}" customFormat="false" ht="12.8" hidden="false" customHeight="false" outlineLevel="0" collapsed="false">
				${row.cells.map(cell => `
				<c r="${cell.id}" s="0" t="${cell.type}"><v>${cell.value}</v></c>
				`).join('\n')}
			</row>
			`).join('\n')}
		</sheetData>
	</worksheet>`;

const SHAREDSTRINGS_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${workbook.stringcount}" uniqueCount="${workbook.stringcount}">
		${workbook.strings.map(string => `
		<si><t xml:space="preserve">${string}</t></si>
		`).join('\n')}
	</sst>`;

const STYLES_XML = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts><fonts count="4"><font><sz val="10"/><name val="Arial"/><family val="2"/></font><font><sz val="10"/><name val="Arial"/><family val="0"/></font><font><sz val="10"/><name val="Arial"/><family val="0"/></font><font><sz val="10"/><name val="Arial"/><family val="0"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border diagonalUp="false" diagonalDown="false"><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="20"><xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="true" applyAlignment="true" applyProtection="true"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="43" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="41" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="44" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="42" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf><xf numFmtId="9" fontId="1" fillId="0" borderId="0" applyFont="true" applyBorder="false" applyAlignment="false" applyProtection="false"></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false" applyBorder="false" applyAlignment="false" applyProtection="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf></cellXfs><cellStyles count="6"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Comma" xfId="15" builtinId="3"/><cellStyle name="Comma [0]" xfId="16" builtinId="6"/><cellStyle name="Currency" xfId="17" builtinId="4"/><cellStyle name="Currency [0]" xfId="18" builtinId="7"/><cellStyle name="Percent" xfId="19" builtinId="5"/></cellStyles></styleSheet>`;

const WORKBOOK_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<fileVersion appName="LifeCourse"/>
		<workbookPr backupFile="false" showObjects="all" date1904="false"/>
		<workbookProtection/>
		<bookViews>
			<workbookView showHorizontalScroll="true" showVerticalScroll="true" showSheetTabs="true"
				xWindow="0" yWindow="0" windowWidth="1024" windowHeight="768" tabRatio="500" firstSheet="0"
				activeTab="0"/>
		</bookViews>
		<sheets>
			${workbook.sheets.map(sheet => `
			<sheet name="${sheet.title}" sheetId="${sheet.id}" state="visible" r:id="${sheet.rId}"/>
			`).join('\n')}
		</sheets>
	</workbook>`;

const CONTENT_TYPES_XML = workbook => `<?xml version="1.0" encoding="UTF-8"?>
	<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
		<Default Extension="xml" ContentType="application/xml"/>
		<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
		<Default Extension="png" ContentType="image/png"/>
		<Default Extension="jpeg" ContentType="image/jpeg"/>
		<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
		<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
		${workbook.sheets.map(sheet => `
		<Override PartName="/xl/worksheets/sheet${sheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
		`).join('\n')}
		<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
		<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
		<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
		<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
		<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
	</Types>`;

const toAlpha = (number) => {
	var baseChar = ("A").charCodeAt(0),
	letters  = "";

	do {
		number -= 1;
		letters = String.fromCharCode(baseChar + (number % 26)) + letters;
		number = (number / 26) >> 0; // quick `floor`
	} while(number > 0);

	return letters;
};

export const generate = (sheets) => {
	let workbook = {
		isoDate: new Date().toISOString(),
		sheets: [],
		strings: []
	};

	let i = 0;
	for (let { title, data } of sheets) {
		let id = ++i,
			rId = `rId${i + 2}`,
			colCount = 1,
			rowCount = data.length,
			rows = [];
		for (let y = 1; y <= rowCount; y++) {
			let row = data[y - 1],
				cells = [],
				cellCount = row.length;
			if (cellCount > colCount) { colCount = cellCount; }
			for (let x = 1; x <= cellCount; x++) {
				let cell = row[x - 1],
					type = typeof cell === 'string' ? 's' : 'n',
					value = cell;
				if (type === 's') {
					value = workbook.strings.indexOf(cell);
					if (value === -1) {
						workbook.strings.push(cell);
						value = workbook.strings.length - 1;
					}
				}
				cells.push({ id: toAlpha(x) + y, type, value });
			}
			rows.push({ id: y, cells });
		}

		let extent = toAlpha(colCount) + rowCount;
		workbook.sheets.push({ id, rId, title, rows, extent });
	}
	workbook.stringcount = workbook.strings.length;

	var zip = new JSZip();

	zip.file('[Content_Types].xml', CONTENT_TYPES_XML(workbook));
	zip.file('_rels/.rels', RELS_XML(workbook)); // applications look here first
	zip.file('docProps/app.xml', APP_XML(workbook)); // metadata about which application generated the file
	zip.file('docProps/core.xml', CORE_XML(workbook));
	zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_XML_RELS(workbook));
	zip.file('xl/sharedStrings.xml', SHAREDSTRINGS_XML(workbook));
	zip.file('xl/styles.xml', STYLES_XML(workbook));
	zip.file('xl/workbook.xml', WORKBOOK_XML(workbook));
	for (let sheet of workbook.sheets) {
		zip.file(`xl/worksheets/sheet${sheet.id}.xml`, SHEET_XML(sheet));
	}

	return {
		async blob () {
			if (isNode) { throw new Error('blob not supported on this platform'); }
			return await zip.generateAsync({ type: 'blob' });
		},
		async base64 () {
			return await zip.generateAsync({ type: 'base64' });
		},
		stream () {
			if (!isNode) { throw new Error('stream not supported on this platform'); }
			return zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true });
		},
		async write (filename) {
			if (!isNode) { throw new Error('write not supported on this platform'); }
			return await new Promise(resolve => {
				zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
					.pipe(fs.createWriteStream(filename))
					.on('finish', () => resolve());
			});
		}
	};
};

export default { generate };