import { strToU8, zipSync } from "fflate";
import FileSaver from "file-saver";

import generateWorkbookXml from "./statics/workbook.xml.js";
import generateWorkbookXmlRels from "./statics/workbook.xml.rels.js";
import rels from "./statics/rels.js";
import contentTypes from "./statics/[Content_Types].xml.js";

import { generateSheets } from "./writeXlsxFile.common.js";

export default function writeXlsxFile(data, { fileName, ...rest } = {}) {
  return generateXlsxFile(data, rest).then((blob) => {
    if (fileName) {
      return FileSaver.saveAs(blob, fileName);
    }
    return blob;
  });
}

/**
 * Writes an *.xlsx file into a "blob".
 * https://github.com/egeriis/zipcelx/issues/68
 * "The reason if you want to send the excel file or store it natively on Cordova/capacitor app".
 * @return {Blob}
 */
function generateXlsxFile(
  data,
  {
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
    dateFormat,
  }
) {
  let zip = {};

  // Encode and add static files to the zip object
  zip["_rels/.rels"] = strToU8(rels);
  zip["[Content_Types].xml"] = strToU8(contentTypes);

  // Generate sheet-related XML data
  const { sheets, getSharedStringsXml, getStylesXml } = generateSheets({
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
    dateFormat,
  });

  const xl = {};
  xl["_rels/workbook.xml.rels"] = strToU8(generateWorkbookXmlRels({ sheets }));
  xl["workbook.xml"] = strToU8(generateWorkbookXml({ sheets, stickyRowsCount, stickyColumnsCount }));
  xl["styles.xml"] = strToU8(getStylesXml());
  xl["sharedStrings.xml"] = strToU8(getSharedStringsXml());

  for (const { id, data } of sheets) {
    xl[`worksheets/sheet${id}.xml`] = strToU8(data);
  }

  zip["xl"] = xl;

  // Create the zip buffer
  const zipBuffer = zipSync(zip, { level: 9 });

  // Create a blob from the zip buffer
  const blob = new Blob([zipBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  return Promise.resolve(blob);
}

// function strToU8(str) {
//   return new TextEncoder().encode(str);
// }
