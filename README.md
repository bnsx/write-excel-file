# `write-excel-file`

Write simple `*.xlsx` files in a browser or Node.js.

[Demo](https://catamphetamine.gitlab.io/write-excel-file/)

Also check [`read-excel-file`](https://gitlab.com/catamphetamine/read-excel-file).

## Install

```js
npm install write-excel-file --save
```

If you're not using a bundler then use a [standalone version from a CDN](#cdn).

## Use

To write an `*.xlsx` file, either provide a `value` and a `type` for each cell:

```js
const data = [
  // Row #1
  [
    // Column #1
    {
      value: 18,
      type: Number
    },
    // Column #2
    {
      value: new Date(),
      type: Date,
      format: 'mm/dd/yyyy'
    },
    // Column #3
    {
      value: 'John Smith',
      type: String
    },
    // Column #4
    {
      value: true,
      type: Boolean
    }
  ],
  // Row #2
  [
    // Column #1
    {
      value: 16,
      type: Number
    },
    // Column #2
    {
      value: new Date(),
      type: Date,
      format: 'mm/dd/yyyy'
    },
    // Column #3
    {
      value: 'Alice Brown',
      type: String
    },
    // Column #4
    {
      value: false,
      type: Boolean
    }
  ]
]
```

Or provide data `objects` and a `schema`:

```js
const objects = [
  // Row #1
  {
    name: 'John Smith',
    age: 18,
    dateOfBirth: new Date(),
    graduated: true
  },
  // Row #2
  {
    name: 'Alice Brown',
    age: 16,
    dateOfBirth: new Date(),
    graduated: false
  }
]
```

```js
const schema = [
  // Column #1
  {
    column: 'Name',
    type: String,
    value: student => student.name,
    // (optional) Column width (in characters).
    width: 20
  },
  // Column #2
  {
    column: 'Age',
    type: Number,
    value: student => student.age
  },
  // Column #3
  {
    column: 'Date of Birth',
    type: Date,
    format: 'mm/dd/yyyy',
    value: student => student.dateOfBirth
  },
  // Column #4
  {
    column: 'Graduated',
    type: Boolean,
    value: student => student.graduated
  }
]
```

If no `type` is specified for a column or a cell then it's assumed to be a `String`.

There're also some additional exported `type`s available:

* `Integer` for integer `Number`s.
* `URL` for URLs.
* `Email` for email addresses.

Each column or cell, aside from having `type` and `value`, can also have:

* `width: number` — Approximate column width (in characters). Example: `20`.

<!--
* `formatId: number` — A [built-in](https://xlsxwriter.readthedocs.io/format.html#format-set-num-format) Excel data format ID (like a date or a currency). Example: `4` for formatting `12345.67` as `12,345.67`.
-->

* `format: string` — A custom cell data format. Can only be used on `Date`, `Number` or `Integer` cells. [Examples](https://xlsxwriter.readthedocs.io/format.html#format-set-num-format):

  * `"0.000"` for printing a floating-point number with 3 decimal places.
  * `"#,##0.00"` for printing currency.
  * `"mm/dd/yy"` for formatting a date (all `Date` cells or columns require a `format`).

* `fontWeight: string` — Can be used to print text in bold. Can only be used on `String` cells. Example: `"bold"`.

### Browser

```js
import writeXlsxFile from 'write-excel-file'

await writeXlsxFile(objects, table, {
  fileName: 'Students.xlsx'
})
```

Uses [`file-saver`](https://www.npmjs.com/package/file-saver) to save the `*.xlsx` file from a web browser.

If `fileName` parameter is not passed, the returned `Promise` resolves to a ["blob"](https://github.com/egeriis/zipcelx/issues/68).

### Node.js

```js
const { writeXlsxFile } = require('write-excel-file/node')

await writeXlsxFile(objects, table, {
  filePath: '/path/to/file.xlsx'
})
```

If `filePath` parameter is not passed, the returned `Promise` resolves to a `Stream`-like object having a `.pipe()` method:

```js
const output = fs.createWriteStream(...)
const stream = await writeXlsxFile(objects)
stream.pipe(output)
```

## TypeScript

Not implemented. I'm not familiar with TypeScript.

## Browser compatibility

Node.js `*.xlxs` parser uses `xpath` and `xmldom` packages for XML parsing. The same packages could be used in a browser because [all modern browsers](https://caniuse.com/#search=domparser) (except IE 11) have native `DOMParser` built-in which could is used instead (meaning smaller footprint and better performance) but since Internet Explorer 11 support is still required the browser version doesn't use the native `DOMParser` and instead uses `xpath` and `xmldom` packages for XML parsing just like the Node.js version.

## Gotchas

### Formulas

Dynamically calculated cells using formulas (`SUM`, etc) are not supported.

## Advanced

By default it reads the first sheet in the document. If you have multiple sheets in your spreadsheet then pass either `sheet: number` (sheet index, starting from `1`) or `sheet: string` (sheet name) as part of the `options` argument (`options.sheet` is `1` by default):

```js
readXlsxFile(file, { sheet: 2 }).then((data) => {
  ...
})
```

```js
readXlsxFile(file, { sheet: 'Sheet1' }).then((data) => {
  ...
})
```

To get the list of sheets one can pass `getSheets: true` option:

```js
readXlsxFile(file, { getSheets: true }).then((sheets) => {
  // sheets === [{ name: 'Sheet1' }, { name: 'Sheet2' }]
})
```

## CDN

One can use any npm CDN service, e.g. [unpkg.com](https://unpkg.com) or [jsdelivr.net](https://jsdelivr.net)

```html
<script src="https://unpkg.com/write-excel-file@1.x/bundle/write-excel-file.min.js"></script>

<script>
  writeXlsxFile(objects, tableDefinition, 'data.xlsx')
</script>
```

## References

For XML parsing [`xmldom`](https://github.com/jindw/xmldom) and [`xpath`](https://github.com/goto100/xpath) are used.

Writing `*.xlsx` files has been copy-pasted from [`zipcelx`](https://medium.com/@Nopziiemoo/create-excel-files-using-javascript-without-all-the-fuss-2c4aa5377813) package.

## GitHub

On March 9th, 2020, GitHub, Inc. silently [banned](https://medium.com/@catamphetamine/how-github-blocked-me-and-all-my-libraries-c32c61f061d3) my account (and all my libraries) without any notice. I opened a support ticked but they didn't answer. Because of that, I had to move all my libraries to [GitLab](https://gitlab.com/catamphetamine).

## License

[MIT](LICENSE)

