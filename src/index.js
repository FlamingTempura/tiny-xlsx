import JSZip from 'jszip';

let isNode = typeof module !== 'undefined' && module.exports,
	fs;
if (isNode) {
	fs = require('fs');
}

const RELS_XML = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
	</Relationships>`;

const APP_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
		        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
		<Template></Template>
		<TotalTime>0</TotalTime>
		<Application>LifeCourse</Application>
		<DocSecurity>0</DocSecurity>
		<ScaleCrop>false</ScaleCrop>
		<HeadingPairs>
			<vt:vector size="${workbook.sheets.length}" baseType="variant">
				<vt:variant>
					<vt:lpstr>Worksheets</vt:lpstr>
				</vt:variant>
				<vt:variant>
					<vt:i4>${workbook.sheets.length}</vt:i4>
				</vt:variant>
			</vt:vector>
		</HeadingPairs>
		<TitlesOfParts>
			<vt:vector size="${workbook.sheets.length}" baseType="lpstr">
				${workbook.sheets.map(sheet => `
				<vt:lpstr>${sheet.title}</vt:lpstr>
				`)}
			</vt:vector>
		</TitlesOfParts>
		<LinksUpToDate>false</LinksUpToDate>
		<SharedDoc>false</SharedDoc>
		<HyperlinksChanged>false</HyperlinksChanged>
		<AppVersion>1.0000</AppVersion>
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
			<sheetView tabSelected="1" zoomScale="79" zoomScaleNormal="79" workbookViewId="0"/>
		</sheetViews>
		<sheetFormatPr defaultRowHeight="12.8"></sheetFormatPr>
		<cols>
			<col max="1025" min="1" style="0" width="11.52"/>
		</cols>
		<sheetData>
			${sheet.rows.map(row => `
			<row r="${row.id}">
				${row.cells.map(cell => `
				<c r="${cell.id}" t="${cell.type}"><v>${cell.value}</v></c>
				`).join('\n')}
			</row>
			`).join('\n')}
		</sheetData>
		<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
	</worksheet>`;

const SHAREDSTRINGS_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${workbook.stringCount}" uniqueCount="${workbook.strings.length}">
		${workbook.strings.map(string => `
		<si><t>${string}</t></si>
		`).join('\n')}
	</sst>`;

const STYLES_XML = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"><fonts count="1" x14ac:knownFonts="1"><font><sz val="10"/><name val="Arial"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`;

const THEME_XML = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>`;

const WORKBOOK_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<fileVersion appName="LifeCourse"/>
		<workbookPr backupFile="false" showObjects="all" date1904="false"/>
		<workbookProtection/>
		<bookViews>
			<workbookView xWindow="0" yWindow="0" windowWidth="1020" windowHeight="765" tabRatio="500"/>
		</bookViews>
		<sheets>
			${workbook.sheets.map(sheet => `
			<sheet name="${sheet.title}" sheetId="${sheet.id}" r:id="${sheet.rId}"/>
			`).join('\n')}
		</sheets>
	</workbook>`;

const CONTENT_TYPES_XML = workbook => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
		<Default Extension="xml" ContentType="application/xml"/>
		<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
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
		isoDate: new Date().toISOString().slice(0, -2) + 'Z',
		sheets: [],
		strings: [],
		stringCount: 0
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
					workbook.stringCount++;
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

	var zip = new JSZip();

	zip.file('[Content_Types].xml', CONTENT_TYPES_XML(workbook));
	zip.file('_rels/.rels', RELS_XML(workbook)); // applications look here first
	zip.file('docProps/app.xml', APP_XML(workbook)); // metadata about which application generated the file
	zip.file('docProps/core.xml', CORE_XML(workbook));
	zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_XML_RELS(workbook));
	zip.file('xl/sharedStrings.xml', SHAREDSTRINGS_XML(workbook));
	zip.file('xl/styles.xml', STYLES_XML(workbook));
	zip.file('xl/theme/theme1.xml', THEME_XML(workbook));
	zip.file('xl/workbook.xml', WORKBOOK_XML(workbook));
	for (let sheet of workbook.sheets) {
		zip.file(`xl/worksheets/sheet${sheet.id}.xml`, SHEET_XML(sheet));
	}

	const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

	return {
		async blob () {
			if (isNode) { throw new Error('blob not supported on this platform'); }
			return await zip.generateAsync({ type: 'blob', mimeType });
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