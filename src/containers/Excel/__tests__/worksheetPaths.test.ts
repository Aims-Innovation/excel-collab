import JSZip from 'jszip';
import { importXLSX } from '../importXLSX';

// Workbook relationships may be relative to xl/workbook.xml or rooted in the
// ZIP package. openpyxl, for example, writes /xl/worksheets/sheet1.xml.
describe('XLSX worksheet relationship paths (Y1-159)', () => {
  test.each([
    'worksheets/sheet1.xml',
    '/xl/worksheets/sheet1.xml',
    './worksheets/sheet1.xml',
    '../xl/worksheets/sheet1.xml',
  ])('retains both sheets and cell content for %s', async (target) => {
    const zip = new JSZip();
    zip.file('xl/workbook.xml', `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>
      <sheet name="Orders" sheetId="1" r:id="rId1"/>
      <sheet name="Summary" sheetId="2" r:id="rId2"/>
      </sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships>
      <Relationship Id="rId1" Target="${target}"/>
      <Relationship Id="rId2" Target="/xl/worksheets/sheet2.xml"/>
      </Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet><sheetData>
      <row r="1"><c r="A1" t="inlineStr"><is><t>Reference</t></is></c></row>
      <row r="2"><c r="A2" t="inlineStr"><is><t>QA159-A</t></is></c>
      <c r="B2" t="n"><v>3</v></c><c r="C2" t="n"><v>11</v></c>
      <c r="D2"><f>B2*C2</f><v>33</v></c></row>
      </sheetData><mergeCells><mergeCell ref="A4:B4"/></mergeCells></worksheet>`);
    zip.file('xl/worksheets/sheet2.xml', `<worksheet><sheetData>
      <row r="1"><c r="A1" t="inlineStr"><is><t>PRESERVE-20261113</t></is></c></row>
      </sheetData></worksheet>`);

    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const model = await importXLSX(bytes as unknown as File);
    expect(Object.values(model.workbook).map((sheet) => sheet.name)).toEqual(['Orders', 'Summary']);
    expect(model.worksheets['1_0_0'].value).toBe('Reference');
    expect(model.worksheets['1_1_0'].value).toBe('QA159-A');
    expect(String(model.worksheets['1_1_1'].value)).toBe('3');
    expect(model.worksheets['1_1_3'].formula).toBe('=B2*C2');
    expect(String(model.worksheets['1_1_3'].value)).toBe('33');
    expect(model.worksheets['2_0_0'].value).toBe('PRESERVE-20261113');
    expect(model.mergeCells['Orders!A4:B4']).toMatchObject({ row: 3, col: 0, colCount: 2 });
  });
});
