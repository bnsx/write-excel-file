// There seem to be about 100 "built-in" formats in Excel.
// https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)?redirectedfrom=MSDN
const FORMAT_ID_STARTS_FROM = 100

export default function initStyles({
  fontFamily,
  fontSize
}) {
  const customFont = fontFamily || fontSize

  if (fontFamily === undefined) {
    fontFamily = 'Calibri'
  }

  if (fontSize === undefined) {
    fontSize = 12
  }

  const formats = []
  const formatsIndex = {}

  const styles = []
  const stylesIndex = {}

  const fonts = []
  const fontsIndex = {}

  const fills = []
  const fillsIndex = {}

  const borders = []
  const bordersIndex = {}

  // Default font.
  fonts.push({
    size: fontSize,
    family: fontFamily
  })
  fontsIndex['-/-'] = 0

  // Default fill.
  fills.push({})
  fillsIndex['-'] = 0

  // Default border.
  borders.push({})
  bordersIndex['-'] = 0

  // "gray125" fill.
  // For some weird reason, MS Office 2007 Excel seems to require that to be present.
  // Otherwise, if absent, it would replace the first `backgroundColor`.
  fills.push({
    gray125: true
  })

  function getStyle({ fontWeight, align, alignVertical, format, wrap, color, backgroundColor }) {
    // Custom borders aren't supported.
    const border = undefined
    // Look for an existing style.
    const fontKey = `${fontWeight || '-'}/${color || '-'}`
    const fillKey = backgroundColor || '-'
    const borderKey = border ? JSON.stringify(border) : '-'
    const key = `${align || '-'}/${alignVertical || '-'}/${format || '-'}/${wrap || '-'}/${fontKey}/${fillKey}/${borderKey}`
    const styleId = stylesIndex[key]
    if (styleId !== undefined) {
      return styleId
    }
    // Create new style.
    // Get format ID.
    let formatId
    if (format) {
      formatId = formatsIndex[format]
      if (formatId === undefined) {
        formatId = formatsIndex[format] = String(FORMAT_ID_STARTS_FROM + formats.length)
        formats.push(format)
      }
    }
    // Get font ID.
    let fontId = customFont ? 0 : undefined
    if (fontWeight || color) {
      fontId = fontsIndex[fontKey]
      if (fontId === undefined) {
        fontId = fontsIndex[fontKey] = String(fonts.length)
        fonts.push({
          size: fontSize,
          family: fontFamily,
          weight: fontWeight,
          color
        })
      }
    }
    // Get fill ID.
    let fillId
    if (backgroundColor) {
      fillId = fillsIndex[fillKey]
      if (fillId === undefined) {
        fillId = fillsIndex[fillKey] = String(fills.length)
        fills.push({
          color: backgroundColor
        })
      }
    }
    // Get border ID.
    let borderId
    if (border) {
      borderId = bordersIndex[borderKey]
      if (borderId === undefined) {
        borderId = bordersIndex[borderKey] = String(borders.length)
        borders.push(border)
      }
    }
    // Add a style.
    styles.push({
      fontId,
      fillId,
      borderId,
      align,
      alignVertical,
      wrap,
      formatId
    })
    return stylesIndex[key] = String(styles.length - 1)
  }

  // Add the default style.
  getStyle({})

  return {
    getStylesXml() {
      return generateXml({ formats, styles, fonts, fills, borders })
    },
    getStyle
  }
}

function generateXml({ formats, styles, fonts, fills, borders }) {
  let xml = '<?xml version="1.0" ?>'
  xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

  // Turns out, as weird as it sounds, the order of XML elements matters to MS Office Excel.
  // https://social.msdn.microsoft.com/Forums/office/en-US/cc47ab65-dab7-4e32-b676-b641aa1e1411/how-to-validate-the-xlsx-that-i-generate?forum=oxmlsdk
  // For example, previously this library was inserting `<cellXfs/>` before `<fonts/>`
  // and that caused MS Office 2007 Excel to throw an error about the file being corrupt:
  // "Excel found unreadable content in '*.xlsx'. Do you want to recover the contents of this workbook?"
  // "Excel was able to open the file by repairing or removing the unreadable content."
  // "Removed Part: /xl/styles.xml part with XML error.  (Styles) Load error. Line 1, column ..."
  // "Repaired Records: Cell information from /xl/worksheets/sheet1.xml part"

  if (formats.length > 0) {
    xml += `<numFmts count="${formats.length}">`
    for (let i = 0; i < formats.length; i++) {
      xml += `<numFmt numFmtId="${FORMAT_ID_STARTS_FROM + i}" formatCode="${formats[i]}"/>`
    }
    xml += `</numFmts>`
  }

  xml += `<fonts count="${fonts.length}">`
  for (const font of fonts) {
    const {
      size,
      family,
      color,
      weight
    } = font
    xml += '<font>'
    xml += `<sz val="${size}"/>`
    xml += `<color ${color ? 'rgb="' + getColor(color) + '"' : 'theme="1"'}/>`
    xml += `<name val="${family}"/>`
    // It's not clear what the `<family/>` tag means or does.
    // It seems to always be `<family val="2"/>` even for different
    // font families (Calibri, Arial, etc).
    xml += '<family val="2"/>'
    // It's not clear what the `<scheme/>` tag means or does.
    xml += '<scheme val="minor"/>'
    if (weight === 'bold') {
      xml += '<b/>'
    }
    xml += '</font>'
  }
  xml += '</fonts>'

  // MS Office 2007 Excel seems to require a `<fills/>` element to exist.
  // without it, MS Office 2007 Excel thinks that the file is broken.
  xml += `<fills count="${fills.length}">`
  for (const fill of fills) {
    const { color, gray125 } = fill
    xml += '<fill>'
    if (color) {
      xml += '<patternFill patternType="solid">'
      xml += `<fgColor rgb="${getColor(color)}"/>`
      // Whatever that could mean.
      xml += '<bgColor indexed="64"/>'
      xml += '</patternFill>'
    } else if (gray125) {
      // "gray125" fill.
      // For some weird reason, MS Office 2007 Excel seems to require that to be present.
      // Otherwise, if absent, it would replace the first `backgroundColor`.
      xml += '<patternFill patternType="gray125"/>'
    } else {
      xml += '<patternFill patternType="none"/>'
    }
    xml += '</fill>'
  }
  xml += '</fills>'

  // MS Office 2007 Excel seems to require a `<borders/>` element to exist:
  // without it, MS Office 2007 Excel thinks that the file is broken.
  xml += `<borders count="${borders.length}">`
  for (const border of borders) {
    xml += '<border>'
    xml += '<left/>'
    xml += '<right/>'
    xml += '<top/>'
    xml += '<bottom/>'
    xml += '<diagonal/>'
    xml += '</border>'
  }
  xml += '</borders>'

  // What are `<cellXfs/>` and `<cellStyleXfs/>`:
  // http://officeopenxml.com/SSstyles.php
  //
  // `<cellStyleXfs/>` are referenced from `<cellXfs/>` as `<xf xfId="..."/>`.
  // `<cellStyleXfs/>` defines abstract "cell styles" that can be "extended"
  // by "cell styles" defined by `<cellXfs/>` that can be applied to individual cells:
  // 1. `<cellStyleXfs><xf .../></cellStyleXfs>`
  // 2. `<cellXfs><xf xfId={cellStyleXfs.xf.index}/></cellXfs>`
  // 3. `<c s={cellXfs.xf.index}/>`
  // Seems like "cell styles" defined by `<cellXfs/>` have to reference
  // some abstract "cell styles" defined by `<cellStyleXfs/>` by the spec.
  // Otherwise, there would be no need to use `<cellStyleXfs/>` at all.
  // The naming is ambiguous and weird. The whole scheme is needlessly redundant.

  // xml += '<cellStyleXfs count="2">'
  // xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
  // // `applyFont="1"` means apply a custom font in this "abstract" "cell style"
  // // rather than using a default font.
  // // Seems like by default `applyFont` is `"0"` meaning that,
  // // unless `"1"` is specified, it would ignore the `fontId` attribute.
  // xml += '<xf numFmtId="0" fontId="1" applyFont="1" fillId="0" borderId="0"/>'
  // xml += '</cellStyleXfs>'

  xml += `<cellXfs count="${styles.length}">`
  for (const cellStyle of styles) {
    const {
      fontId,
      fillId,
      borderId,
      align,
      alignVertical,
      wrap,
      formatId
    } = cellStyle
    // `applyNumberFormat="1"` means "apply the `numFmtId` attribute".
    // Seems like by default `applyNumberFormat` is `"0"` meaning that,
    // unless `"1"` is specified, it would ignore the `numFmtId` attribute.
    xml += '<xf ' +
      [
        formatId !== undefined ? `numFmtId="${formatId}"` : undefined,
        formatId !== undefined ? 'applyNumberFormat="1"' : undefined,
        fontId !== undefined ? `fontId="${fontId}"` : undefined,
        fontId !== undefined ? 'applyFont="1"' : undefined,
        fillId !== undefined ? `fillId="${fillId}"` : undefined,
        fillId !== undefined ? 'applyFill="1"' : undefined,
        borderId !== undefined ? `borderId="${borderId}"` : undefined,
        borderId !== undefined ? 'applyBorder="1"' : undefined,
        align || alignVertical || wrap ? 'applyAlignment="1"' : undefined,
        // 'xfId="0"'
      ].filter(_ => _).join(' ') +
    '>' +
      // Possible horizontal alignment values:
      //  left, center, right, fill, justify, center_across, distributed.
      // Possible vertical alignment values:
      //  top, vcenter, bottom, vjustify, vdistributed.
      // https://xlsxwriter.readthedocs.io/format.html#set_align
      (align || alignVertical || wrap
        ? '<alignment' +
          (align ? ` horizontal="${align}"` : '') +
          (alignVertical ? ` vertical="${alignVertical}"` : '') +
          (wrap ? ` wrapText="1"` : '') +
          '/>'
        : ''
      ) +
    '</xf>'
  }
  xml += `</cellXfs>`

  xml += '</styleSheet>'

  return xml
}

function getColor(color) {
  if (color[0] !== '#') {
    throw new Error(`Color "${color}" must start with a "#"`)
  }
  return `FF${color.slice('#'.length).toUpperCase()}`
}