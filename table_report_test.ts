import { assertEquals, ExcelJS, pathExistsSync } from './deps.ts'
import { Table } from './mod.ts'
import {
  convertCacadeArrayKey2ObjectKey,
  DataRow,
  GroupColumn,
  recursiveGenDataRowExtProperties,
} from './table_report.ts'
import {
  CellStyle,
  flattenColumnByChildren,
  genColumnExtProperties,
  genSingleSheetWorkbook,
  HeadColumn,
  recursiveGenDataRowCell,
} from './table_report.ts'

if (!pathExistsSync('temp')) Deno.mkdirSync('temp')

Deno.test('genColumnExtProperties', () => {
  let t: Table
  const startRow = 1

  // case 1
  t = { headColumns: [{ key: 'k0' }] }
  genColumnExtProperties(t.headColumns, { startRow })
  assertEquals(t.headColumns.length, 1)
  assertEquals(t.headColumns[0].ext, { row: 1, col: 1, rowspan: 1, colspan: 1, depth: 1 })

  // case 2
  t = { headColumns: [{ children: [{ key: 'k0' }] }] }
  genColumnExtProperties(t.headColumns, { startRow })
  assertEquals(t.headColumns.length, 1)
  let c = t.headColumns[0] as GroupColumn
  assertEquals(c.ext, { row: 1, col: 1, rowspan: 2, colspan: 1, depth: 2 })
  assertEquals(c.children?.length, 1)
  assertEquals(c.children?.[0]?.ext, { row: 2, col: 1, rowspan: 1, colspan: 1, depth: 1 })

  // case 3
  t = { headColumns: [{ children: [{ key: 'k0' }, { key: 'k1' }] }] }
  genColumnExtProperties(t.headColumns, { startRow })
  assertEquals(t.headColumns.length, 1)
  c = t.headColumns[0] as GroupColumn
  assertEquals(c.ext, { row: 1, col: 1, rowspan: 2, colspan: 2, depth: 2 })
  assertEquals(c.children?.length, 2)
  assertEquals(c.children?.[0]?.ext, { row: 2, col: 1, rowspan: 1, colspan: 1, depth: 1 })
  assertEquals(c.children?.[1]?.ext, { row: 2, col: 2, rowspan: 1, colspan: 1, depth: 1 })

  // case 4
  t = { headColumns: [{ key: 'k0' }, { children: [{ key: 'k10' }, { key: 'k11' }] }] }
  genColumnExtProperties(t.headColumns, { startRow })
  assertEquals(t.headColumns.length, 2)
  assertEquals(t.headColumns[0].ext, { row: 1, col: 1, rowspan: 2, colspan: 1, depth: 1 })
  c = t.headColumns[1] as GroupColumn
  assertEquals(c.ext, { row: 1, col: 2, rowspan: 2, colspan: 2, depth: 2 })
  assertEquals(c.children?.length, 2)
  assertEquals(c.children?.[0]?.ext, { row: 2, col: 2, rowspan: 1, colspan: 1, depth: 1 })
  assertEquals(c.children?.[1]?.ext, { row: 2, col: 3, rowspan: 1, colspan: 1, depth: 1 })

  // case 5
  t = {
    headColumns: [{ key: 'k0' }, { children: [{ key: 'k10' }, { children: [{ key: 'k110' }, { key: 'k111' }] }] }],
  }
  genColumnExtProperties(t.headColumns, { startRow })
  assertEquals(t.headColumns.length, 2)
  assertEquals(t.headColumns[0].ext, { row: 1, col: 1, rowspan: 3, colspan: 1, depth: 1 })
  assertEquals(t.headColumns[1].ext, { row: 1, col: 2, rowspan: 3, colspan: 3, depth: 3 })
  c = t.headColumns[1] as GroupColumn
  assertEquals(c.children?.length, 2)
  assertEquals(c.children?.[0]?.ext, { row: 2, col: 2, rowspan: 2, colspan: 1, depth: 1 })
  assertEquals(c.children?.[1]?.ext, { row: 2, col: 3, rowspan: 2, colspan: 2, depth: 2 })
  const cc = c.children?.[1] as GroupColumn
  assertEquals(cc?.children?.length, 2)
  assertEquals(cc?.children?.[0]?.ext, { row: 3, col: 3, rowspan: 1, colspan: 1, depth: 1 })
  assertEquals(cc?.children?.[1]?.ext, { row: 3, col: 4, rowspan: 1, colspan: 1, depth: 1 })
})

Deno.test('flattenColumnByChildren', () => {
  const columns: HeadColumn[] = [
    { key: 'l0', label: 'l0' },
    {
      key: 'l1',
      label: 'l1',
      children: [
        { key: 'l10', label: 'l10' },
        { key: 'l11', label: 'l11', children: [{ key: 'l110', label: 'l110' }] },
      ],
    },
  ]
  const flatten = flattenColumnByChildren(columns)
  assertEquals(flatten.length, 5)
  assertEquals(flatten[0], columns[0])
  assertEquals(flatten[1], columns[1])
  assertEquals(flatten[2], (columns[1] as GroupColumn).children[0])
  assertEquals(flatten[3], (columns[1] as GroupColumn).children[1])
  assertEquals(flatten[4], ((columns[1] as GroupColumn).children[1] as GroupColumn).children[0])
})

const style: CellStyle = {
  alignment: { vertical: 'middle' },
  border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } },
}

Deno.test('recursiveGenDataCell', async () => {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Sheet1')
  const datas: Record<string, unknown>[] = [
    {
      k: 't0',
      a: [
        { k: 'a00', ext: { rowspan: 1 } },
        { k: 'a01', b: [{ k: 'a01b0' }], ext: { rowspan: 1 } },
        { k: 'a02', b: [{ k: 'a02b0' }, { k: 'a02b1' }], ext: { rowspan: 2 } },
      ],
      ext: { rowspan: 4 },
    },
    {
      k: 't1',
      a: [
        { k: 'a10', b: [{ k: 'a10b0' }, { k: 'a10b1' }], ext: { rowspan: 2 } },
        { k: 'a11', ext: { rowspan: 1 } },
      ],
      ext: { rowspan: 3 },
    },
  ]

  const startRow = 3
  // column 1
  let row = startRow, col = 2
  ws.getCell(row - 1, col).value = 'k'
  datas.forEach((data, index) => {
    row = recursiveGenDataRowCell(['k'], data, {
      ws,
      index,
      row: row,
      col: col,
      style,
    })
  })

  // column 2
  row = startRow, col++
  ws.getCell(row - 1, col).value = 'a.k'
  datas.forEach((data, index) => {
    row = recursiveGenDataRowCell(['a', 'k'], data, {
      ws,
      index,
      row: row,
      col: col,
      style,
    })
  })

  // column 3
  row = startRow, col++
  ws.getCell(row - 1, col).value = 'a.b.k'
  datas.forEach((data, index) => {
    row = recursiveGenDataRowCell(['a', 'b', 'k'], data, {
      ws,
      index,
      row: row,
      col: col,
      style,
    })
  })

  await wb.xlsx.writeFile('temp/recursive_gen_data_cell.xlsx')
})

Deno.test('convertCacadeArrayKey2ObjectKey', () => {
  assertEquals(convertCacadeArrayKey2ObjectKey([]), {})
  assertEquals(convertCacadeArrayKey2ObjectKey([['a']]), { a: 0 })
  assertEquals(convertCacadeArrayKey2ObjectKey([['a', 'b']]), { a: { b: 0 } })
  assertEquals(convertCacadeArrayKey2ObjectKey([['a'], ['a', 'b']]), { a: { b: 0 } })
  assertEquals(convertCacadeArrayKey2ObjectKey([['a', 'b'], ['a']]), { a: { b: 0 } })
  assertEquals(convertCacadeArrayKey2ObjectKey([['a', 'b'], ['c'], ['a']]), { a: { b: 0 }, c: 0 })
  assertEquals(convertCacadeArrayKey2ObjectKey([['a', 'b', 'c'], ['d'], ['a', 'b']]), { a: { b: { c: 0 } }, d: 0 })
})

Deno.test('recursiveGenDataRowExtProperties', async (test) => {
  await test.step('case 1', () => {
    assertEquals(recursiveGenDataRowExtProperties([]), 1)
    assertEquals(recursiveGenDataRowExtProperties([], { startRow: 2 }), 2)
  })

  await test.step('case 2', () => {
    const dataRows: DataRow[] = [{}]
    const nextRow = recursiveGenDataRowExtProperties(dataRows)
    assertEquals(nextRow, 2)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 1 })
  })

  await test.step('case 3', () => {
    const dataRows: DataRow[] = [{}]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: 0 } })
    assertEquals(nextRow, 2)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 1 })
  })

  await test.step('case 4', () => {
    const dataRows: DataRow[] = [{ a: [] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 } } })
    assertEquals(nextRow, 2)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 1 })
  })

  await test.step('case 5', () => {
    const dataRows: DataRow[] = [{ a: [{ b: [] }] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 } } })
    assertEquals(nextRow, 2)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 1 })
  })

  await test.step('case 6', () => {
    const dataRows: DataRow[] = [{ a: [{ b: [{}] }] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 } } })
    assertEquals(nextRow, 2)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 1 })
  })

  await test.step('case 7', () => {
    const dataRows: DataRow[] = [{ a: [{ b: [{}, {}] }] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 } } })
    assertEquals(nextRow, 3)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 2 })
    assertEquals(dataRows[0].a[0].ext, { row: 1, rowspan: 2 })
  })

  await test.step('case 8', () => {
    const dataRows: DataRow[] = [{ a: [{}, { b: [{}, {}] }] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 } } })
    assertEquals(nextRow, 4)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 3 })
    assertEquals(dataRows[0].a[0].ext, { row: 1, rowspan: 1 })
    assertEquals(dataRows[0].a[1].ext, { row: 2, rowspan: 2 })
    assertEquals(dataRows[0].a[1].b[0].ext, { row: 2, rowspan: 1 })
    assertEquals(dataRows[0].a[1].b[1].ext, { row: 3, rowspan: 1 })
  })

  await test.step('case 9', () => {
    const dataRows: DataRow[] = [{ a: [{}, { b: [{}, {}] }], c: [{}] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 }, c: 0 } })
    assertEquals(nextRow, 4)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 3 })
    assertEquals(dataRows[0].a[0].ext, { row: 1, rowspan: 1 })
    assertEquals(dataRows[0].a[1].ext, { row: 2, rowspan: 2 })
    assertEquals(dataRows[0].a[1].b[0].ext, { row: 2, rowspan: 1 })
    assertEquals(dataRows[0].a[1].b[1].ext, { row: 3, rowspan: 1 })
    assertEquals(dataRows[0].c[0].ext, { row: 1, rowspan: 1 })
  })

  await test.step('case 10', () => {
    const dataRows: DataRow[] = [{ a: [{}, { b: [{}, {}] }], c: [{}, {}, {}, {}] }]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 }, c: 0 } })
    assertEquals(nextRow, 5)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 4 })
    assertEquals(dataRows[0].a[0].ext, { row: 1, rowspan: 1 })
    assertEquals(dataRows[0].a[1].ext, { row: 2, rowspan: 2 })
    assertEquals(dataRows[0].a[1].b[0].ext, { row: 2, rowspan: 1 })
    assertEquals(dataRows[0].a[1].b[1].ext, { row: 3, rowspan: 1 })
    ;(dataRows[0].c as DataRow[]).forEach((c, i) => assertEquals(c.ext, { row: 1 + i, rowspan: 1 }))
  })

  await test.step('case 11', () => {
    const dataRows: DataRow[] = [{ a: [{}, { b: [{}, {}] }], c: [{}, {}, {}, {}] }, {}]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 }, c: 0 } })
    assertEquals(nextRow, 6)
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 4 })
    assertEquals(dataRows[0].a[0].ext, { row: 1, rowspan: 1 })
    assertEquals(dataRows[0].a[1].ext, { row: 2, rowspan: 2 })
    assertEquals(dataRows[0].a[1].b[0].ext, { row: 2, rowspan: 1 })
    assertEquals(dataRows[0].a[1].b[1].ext, { row: 3, rowspan: 1 })
    ;(dataRows[0].c as DataRow[]).forEach((c, i) => assertEquals(c.ext, { row: 1 + i, rowspan: 1 }))
  })

  await test.step('case 12', () => {
    const dataRows: DataRow[] = [
      { a: [{}, { b: [{}, {}] }], c: [{}, {}] },
      { a: [{ b: [{}, {}] }], c: [{}, {}, {}] },
    ]
    const nextRow = recursiveGenDataRowExtProperties(dataRows, { cascadeKey: { a: { b: 0 }, c: 0 } })
    assertEquals(nextRow, 7)

    // dataRow[0]
    assertEquals(dataRows[0].ext, { row: 1, rowspan: 3 })
    assertEquals(dataRows[0].a[0].ext, { row: 1, rowspan: 1 })
    assertEquals(dataRows[0].a[1].ext, { row: 2, rowspan: 2 })
    assertEquals(dataRows[0].a[1].b[0].ext, { row: 2, rowspan: 1 })
    assertEquals(dataRows[0].a[1].b[1].ext, { row: 3, rowspan: 1 })
    ;(dataRows[0].c as DataRow[]).forEach((c, i) => assertEquals(c.ext, { row: 1 + i, rowspan: 1 }))

    // dataRow[1]
    assertEquals(dataRows[1].ext, { row: 4, rowspan: 3 })
    assertEquals(dataRows[1].a[0].ext, { row: 4, rowspan: 2 })
    assertEquals(dataRows[1].a[0].b[0].ext, { row: 4, rowspan: 1 })
    assertEquals(dataRows[1].a[0].b[1].ext, { row: 5, rowspan: 1 })
    ;(dataRows[1].c as DataRow[]).forEach((c, i) => assertEquals(c.ext, { row: 4 + i, rowspan: 1 }))
  })
})

Deno.test('gen simple table', async () => {
  // define head-column
  const headColumns: HeadColumn[] = [
    {
      key: 'sn',
      label: 'SN',
      width: 5,
      mapper: ({ index }) => index + 1,
      dataCellStyle: { numFmt: '#0', alignment: { horizontal: 'right' } },
    },
    { key: 'name', label: 'Name', width: 15, mapper: ({ row }) => `${row.firstName} ${row.lastName}` },
    { key: 'date', width: 12, dataCellStyle: { numFmt: 'yyyy-MM-dd', alignment: { horizontal: 'center' } } },
    { key: 'int', label: 'Int', dataCellStyle: { numFmt: '#0', alignment: { horizontal: 'right' } } },
    { key: 'decimal', label: 'Decimal', dataCellStyle: { numFmt: '#0.00', alignment: { horizontal: 'right' } } },
    {
      key: 'money',
      label: 'Money',
      width: 22,
      dataCellStyle: { numFmt: '￥#,###,###,##0.00', alignment: { horizontal: 'right' } },
    },
    { key: 'percent', label: 'Percent', dataCellStyle: { numFmt: '0.00%', alignment: { horizontal: 'right' } } },
  ]

  // define data-row
  const dataRows = [
    {
      firstName: 'John',
      lastName: 'Smith',
      date: '2024-01-01',
      int: 1234,
      decimal: 1234.567,
      money: 1234567890.345,
      percent: 0.12345,
    },
    {
      firstName: 'Li',
      lastName: 'Xiao',
    },
    {
      firstName: 'Chen',
      lastName: 'Hui',
      date: '2024-01-03',
      int: 222.333,
      decimal: 888.666,
      percent: 1,
    },
  ]

  // generate a Workbook
  const workbook = await genSingleSheetWorkbook({
    headColumns,
    dataRows,
    sheetName: 'Sheet1',
    // table title
    caption: { value: 'Simple table example', style: { font: { bold: true } } },
    // share style for all head-cell and data-cell
    cellStyle: {
      alignment: { vertical: 'middle' },
      border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } },
    },
    // specific style for all head-cell
    headCellStyle: {
      alignment: { horizontal: 'center' },
      font: { bold: true },
    },
    // specific style for all data-cell
    dataCellStyle: { alignment: { horizontal: 'left' } },
    // workbook properties
    bookProperties: {
      creator: 'NextRJ',
      created: new Date(),
      lastModifiedBy: 'NextRJ',
      modified: new Date(),
    },
    // sheet properties
    sheetProperties: {
      defaultRowHeight: 20, // default 15
      defaultColWidth: 10,
    },
    sheetView: {
      showGridLines: false,
      state: 'frozen',
      xSplit: 1,
      ySplit: 2,
      activeCell: 'B3',
    },
    sheetPageSetup: {
      paperSize: 9, // 8-A3、9-A4
      orientation: 'landscape',
      // units is inches, 0.2"*2.54=0.5cm
      margins: { top: 0.2, left: 0.2, bottom: 0.2, right: 0.2, header: 0.2, footer: 0.2 },
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      printArea: 'A:F',
      printTitlesRow: '2:2',
    },
  })

  // write to file
  const file = 'temp/sample_simple_table.xlsx'
  await workbook.xlsx.writeFile(file)
  console.log(`write to file '${file}'`)
})

Deno.test('gen table with group', async () => {
  // define head-column
  const headColumns: HeadColumn[] = [
    {
      key: 'sn',
      label: 'SN',
      width: 5,
      mapper: ({ index }) => index + 1,
      dataCellStyle: { numFmt: '#0', alignment: { horizontal: 'right' } },
    },
    { key: 'teacher', label: 'Teacher', width: 15, mapper: ({ row }) => `${row.firstName} ${row.lastName}` },
    {
      key: 'workdate',
      label: 'Workdate',
      width: 12,
      dataCellStyle: { numFmt: 'yyyy-MM-dd', alignment: { horizontal: 'center' } },
    },
    {
      label: 'Students',
      children: [
        {
          keys: ['students', 'sn'],
          label: 'SN',
          width: 5,
          mapper: ({ index }) => index + 1,
          dataCellStyle: { numFmt: '#0', alignment: { horizontal: 'right' } },
        },
        { keys: ['students', 'name'], label: 'Name' },
        {
          keys: ['students', 'birthdate'],
          label: 'Birthdate',
          width: 12,
          dataCellStyle: { numFmt: 'yyyy-MM-dd', alignment: { horizontal: 'center' } },
        },
      ],
    },
    {
      label: 'Favorites',
      children: [
        {
          keys: ['favorites', 'sn'],
          label: 'SN',
          width: 5,
          mapper: ({ index }) => index + 1,
          dataCellStyle: { numFmt: '#0', alignment: { horizontal: 'right' } },
        },
        { keys: ['favorites', 'name'], label: 'Name' },
        {
          keys: ['favorites', 'priority'],
          label: 'Priority',
          width: 12,
          dataCellStyle: { numFmt: '#', alignment: { horizontal: 'right' } },
        },
      ],
    },
  ]

  // define data-row
  const dataRows = [
    { firstName: 'John', lastName: 'Smith', workdate: '2000-01-01' },
    {
      firstName: 'Li',
      lastName: 'Xiao',
      students: [
        { name: 'Lili', birthdate: '2020-01-01' },
        { name: 'Suson', birthdate: '2023-12-01' },
      ],
      favorites: [
        { priority: 10, name: 'Draw' },
        { priority: 30, name: 'Math' },
        { priority: 20, name: 'Sport' },
      ],
    },
    {
      firstName: 'Chen',
      lastName: 'Hui',
      workdate: '2007-02-01',
      students: [
        { name: 'Peter', birthdate: '2019-01-01' },
        { name: 'Alan', birthdate: '2024-03-13' },
        { name: 'Rocky', birthdate: '2022-11-23' },
      ],
      favorites: [
        { priority: 1, name: 'Tech' },
        { priority: 2, name: 'Swim' },
      ],
    },
    { firstName: 'Zhang', lastName: 'HuQi', favorites: [{ priority: 1, name: 'Sing' }] },
  ]

  // generate a Workbook
  const workbook = await genSingleSheetWorkbook({
    headColumns,
    dataRows,
    sheetName: 'Sheet1',
    caption: { value: 'ABC primary school', style: { font: { bold: true } } },
    subCaption: 'Class One Grade Two',
    cellStyle: {
      alignment: { vertical: 'middle' },
      border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } },
    },
    headCellStyle: { alignment: { horizontal: 'center' }, font: { bold: true } },
    dataCellStyle: { alignment: { horizontal: 'left' } },
    sheetProperties: {
      defaultRowHeight: 20, // default 15
      defaultColWidth: 10,
    },
    sheetView: {
      showGridLines: false,
      state: 'frozen',
      xSplit: 1,
      ySplit: 4,
      activeCell: 'B5',
    },
  })

  // write to file
  const file = 'temp/sample_table_with_group.xlsx'
  await workbook.xlsx.writeFile(file)
  console.log(`write to file '${file}'`)
})
