const fontSize = 12
const fontFamily = 'Calibri'

export const BOLD_FONT_WEIGHT_STYLE_ID = 1

export default function generateStyles(data, { schema }) {
  const customNumberFormats = []
  const cellStyles = [
    // Regular font.
    {
      xfId: 0
    },
    // Bold font.
    {
      xfId: 1
    }
  ]

  const formatStyles = {}

  function addFormat(format) {
    formatStyles[format] = cellStyles.length
    // There seem to be about 100 "built-in" formats in Excel.
    // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)?redirectedfrom=MSDN
    const formatId = 100 + customNumberFormats.length
    // A custom format.
    customNumberFormats.push({
      id: formatId,
      template: format
    })
    cellStyles.push({
      numberFormat: {
        id: formatId
      }
    })
  }

  if (schema) {
    for (const column of schema) {
      if (column.format) {
        if (!formatStyles[column.format]) {
          addFormat(column.format)
        }
      }
    }
  } else {
    for (const row of data) {
      for (const cell of row) {
        if (cell.format) {
          if (!formatStyles[cell.format]) {
            addFormat(cell.format)
          }
        }
      }
    }
  }

  // // A "built-in" format.
  // if (column.formatId) {
  //   if (formatStyles[String(column.formatId)]) {
  //     continue
  //   }
  //   formatStyles[String(column.formatId)] = cellStyles.length
  //   cellStyles.push({
  //     numberFormat: {
  //       id: column.formatId
  //     }
  //   })
  // }

  let styles = '<?xml version="1.0" ?>'
  styles += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

  // Turns out, as weird as it sounds, the order of XML elements matters to MS Office Excel.
  // https://social.msdn.microsoft.com/Forums/office/en-US/cc47ab65-dab7-4e32-b676-b641aa1e1411/how-to-validate-the-xlsx-that-i-generate?forum=oxmlsdk
  // For example, previously this library was inserting `<cellXfs/>` before `<fonts/>`
  // and that caused MS Office 2007 Excel to throw an error about the file being corrupt:
  // "Excel found unreadable content in '*.xlsx'. Do you want to recover the contents of this workbook?"
  // "Excel was able to open the file by repairing or removing the unreadable content."
  // "Removed Part: /xl/styles.xml part with XML error.  (Styles) Load error. Line 1, column ..."
  // "Repaired Records: Cell information from /xl/worksheets/sheet1.xml part"

  styles += `<numFmts count="${customNumberFormats.length}">`
  for (const numberFormat of customNumberFormats) {
    styles += `<numFmt numFmtId="${numberFormat.id}" formatCode="${numberFormat.template}"/>`
  }
  styles += `</numFmts>`

  styles += '<fonts count="2">'
  styles += `<font><sz val="${fontSize}"/><color theme="1"/><name val="${fontFamily}"/><family val="2"/><scheme val="minor"/></font>`
  styles += `<font><b/><sz val="${fontSize}"/><color theme="1"/><name val="${fontFamily}"/><family val="2"/><scheme val="minor"/></font>`
  styles += '</fonts>'

  styles += '<fills count="1">'
  styles += '<fill><patternFill patternType="none"/></fill>'
  styles += '</fills>'

  styles += '<borders count="1">'
  styles += '<border><left/><right/><top/><bottom/><diagonal/></border>'
  styles += '</borders>'

  styles += '<cellStyleXfs count="2">'
  styles += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
  styles += '<xf numFmtId="0" fontId="1" applyFont="1" fillId="0" borderId="0"/>'
  styles += '</cellStyleXfs>'

  // What is `<cellXfs/>` and `<cellStyleXfs/>`:
  // http://officeopenxml.com/SSstyles.php
  styles += `<cellXfs count="${cellStyles.length}">`
  for (const cellStyle of cellStyles) {
    // It's not clear what `applyNumberFormat="1"` means.
    // Perhaps it's required to be present.
    styles += `<xf numFmtId="${cellStyle.numberFormat ? cellStyle.numberFormat.id : 0}" applyNumberFormat="${cellStyle.numberFormat ? 1 : 0}" fontId="0" fillId="0" borderId="0" xfId="${cellStyle.xfId || 0}"/>`
  }
  styles += `</cellXfs>`

  styles += '</styleSheet>'

  return {
    styles,
    formatStyles
  }
}