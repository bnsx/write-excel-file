# `write-excel-file`

Write simple `*.xlsx` files in a browser or Node.js

[Demo](https://catamphetamine.gitlab.io/write-excel-file/)

Also check [`read-excel-file`](https://www.npmjs.com/package/read-excel-file)

## Install

```js
npm install write-excel-file --save
```

If you're not using a bundler then use a [standalone version from a CDN](#cdn).

## Use

To write an `*.xlsx` file, provide the `data` — an array of rows, each row being an array of cells, each cell having a `type` and a `value`:

```js
const data = [
  // Row #1
  [
    // Column #1
    {
      type: Number,
      value: 18
    },
    // Column #2
    {
      type: Date,
      value: new Date(),
      format: 'mm/dd/yyyy'
    },
    // Column #3
    {
      type: String,
      value: 'John Smith'
    },
    // Column #4
    {
      type: Boolean,
      value: true
    }
  ],
  // Row #2
  [
    // Column #1
    {
      type: Number,
      value: 16
    },
    // Column #2
    {
      type: Date,
      value: new Date(),
      format: 'mm/dd/yyyy'
    },
    // Column #3
    {
      type: String,
      value: 'Alice Brown'
    },
    // Column #4
    {
      type: Boolean,
      value: false
    }
  ]
]
```

Alternatively, provide a list of `objects` and a `schema` to transform those `objects` to `data`:

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
    value: student => student.name
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

If no `type` is specified, it defaults to a `String`.

<!--
There're also some additional exported `type`s available:

* `Integer` for integer `Number`s.
* `URL` for URLs.
* `Email` for email addresses.
-->

Aside from having a `type` and a `value`, each cell (or schema column) can also have:

<!-- * `width: number` — Approximate column width (in characters). Example: `20`. -->

<!--
* `formatId: number` — A [built-in](https://xlsxwriter.readthedocs.io/format.html#format-set-num-format) Excel data format ID (like a date or a currency). Example: `4` for formatting `12345.67` as `12,345.67`.
-->

* `format: string` — A custom cell data format. Can only be used on `Date` or `Number` <!-- or `Integer` --> cells. [Examples](https://xlsxwriter.readthedocs.io/format.html#format-set-num-format):

  * `"0.000"` for printing a floating-point number with 3 decimal places.
  * `"#,##0.00"` for printing currency.
  * `"mm/dd/yy"` for formatting a date. All `Date` cells (or schema columns) require a `format`.

* `fontWeight: string` — Can be used to print text in bold. Available values: `"bold"`.

* `align: string` — Can be used to align cell content horizontally. Available values: `"left"`, `"center"`, `"right"`.

### Column Width

Column width can also be specified (in "characters").

#### Schema

To specify column width when using a `schema`, set a `width` on a schema column:

```js
const schema = [
  // Column #1
  {
    column: 'Name',
    type: String,
    value: student => student.name,
    width: 20 // Column width (in characters).
  },
  ...
]
```

#### Cell Data

When not using a schema, one can provide a `columns` parameter to specify a width of a column:

```js
// Set Column #3 width to "20 characters".
const columns = [
  {},
  {},
  { width: 20 }, // in characters
  {}
]
```

### Browser

```js
import writeXlsxFile from 'write-excel-file'

// When passing `data` for each cell.
await writeXlsxFile(data, {
  columns, // optional
  fileName: 'file.xlsx'
})

// When passing `objects` and `schema`.
await writeXlsxFile(objects, {
  schema,
  fileName: 'file.xlsx'
})

```

Uses [`file-saver`](https://www.npmjs.com/package/file-saver) to save an `*.xlsx` file from a web browser.

If `fileName` parameter is not passed then the returned `Promise` resolves to a ["blob"](https://github.com/egeriis/zipcelx/issues/68) with the contents of the `*.xlsx` file.

### Node.js

```js
const { writeXlsxFile } = require('write-excel-file/node')

// When passing `data` for each cell.
await writeXlsxFile(data, {
  columns, // optional
  filePath: '/path/to/file.xlsx'
})

// When passing `objects` and `schema`.
await writeXlsxFile(objects, {
  schema,
  filePath: '/path/to/file.xlsx'
})
```

If `filePath` parameter is not passed then the returned `Promise` resolves to a `Stream`-like object having a `.pipe()` method:

```js
const output = fs.createWriteStream(...)
const stream = await writeXlsxFile(objects)
stream.pipe(output)
```

## TypeScript

Not implemented. I'm not familiar with TypeScript.

## CDN

One can use any npm CDN service, e.g. [unpkg.com](https://unpkg.com) or [jsdelivr.net](https://jsdelivr.net)

```html
<script src="https://unpkg.com/write-excel-file@1.x/bundle/write-excel-file.min.js"></script>

<script>
  writeXlsxFile(objects, schema, {
    fileName: 'file.xlsx'
  })
</script>
```

## References

Writing `*.xlsx` files was originally copy-pasted from [`zipcelx`](https://medium.com/@Nopziiemoo/create-excel-files-using-javascript-without-all-the-fuss-2c4aa5377813) package, and then rewritten.

## GitHub

On March 9th, 2020, GitHub, Inc. silently [banned](https://medium.com/@catamphetamine/how-github-blocked-me-and-all-my-libraries-c32c61f061d3) my account (erasing all my repos, issues and comments, even in my employer's private repos) without any notice or explanation. Because of that, all source codes had to be promptly moved to GitLab. The [GitHub repo](https://github.com/catamphetamine/write-excel-file) is now only used as a backup (you can star the repo there too), and the primary repo is now the [GitLab one](https://gitlab.com/catamphetamine/write-excel-file). Issues can be reported in any repo.

## License

[MIT](LICENSE)

