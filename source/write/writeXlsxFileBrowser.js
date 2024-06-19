import FileSaver from 'file-saver';
import generateWorkbookXml from './statics/workbook.xml.js';
import generateWorkbookXmlRels from './statics/workbook.xml.rels.js';
import rels from './statics/rels.js';
import contentTypes from './statics/[Content_Types].xml.js';
import { generateSheets } from './writeXlsxFile.common.js';
import { compress } from 'fflate'; // Import fflate

export default function writeXlsxFile(data, { fileName, ...rest } = {}) {
  return generateXlsxFile(data, rest).then((blob) => {
    if (fileName) {
      return FileSaver.saveAs(blob, fileName);
    }
    return blob;
  });
}

function generateXlsxFile(data, {
  sheet: sheetName,
  sheets: sheetNames,
  schema,
  columns,
  headerStyle,
  fontFamily,
  fontSize,
  orientation,
  stickyRowsCount,
  stickyColumnsCount,
  dateFormat
}) {
  const zip = new compress(); // Initialize fflate compressor

  zip.file('_rels/.rels', rels);
  zip.file('[Content_Types].xml', contentTypes);

  const {
    sheets,
    getSharedStringsXml,
    getStylesXml
  } = generateSheets({
    data,
    sheetName,
    sheetNames,
    schema,
    columns,
    headerStyle,
    fontFamily,
    fontSize,
    orientation,
    stickyRowsCount,
    stickyColumnsCount,
    dateFormat
  });

  const xl = zip.folder('xl');
  xl.file('_rels/workbook.xml.rels', generateWorkbookXmlRels({ sheets }));
  xl.file('workbook.xml', generateWorkbookXml({ sheets, stickyRowsCount, stickyColumnsCount }));
  xl.file('styles.xml', getStylesXml());
  xl.file('sharedStrings.xml', getSharedStringsXml());

  for (const { id, data } of sheets) {
    xl.file(`worksheets/sheet${id}.xml`, data);
  }

  return zip.compress(); // Use fflate's compression method
}
