import { assertEquals, pathExistsSync } from './deps.ts'
import { genHeadColumnExtParams, genSingleSheetWorkbook, HeadColumn } from './table_report.ts'

if (!pathExistsSync('temp')) Deno.mkdirSync('temp')

Deno.test('genHeadColumnExtParams', () => {
  let cs: HeadColumn[]

  // case 1
  genHeadColumnExtParams(cs = [{}], 1)
  assertEquals(cs.length, 1)
  assertEquals(cs[0].ext, { row: 1, col: 1, rowspan: 1, colspan: 1, depth: 1 })

  // case 2
  genHeadColumnExtParams(cs = [{ children: [{}] }], 1)
  assertEquals(cs.length, 1)
  assertEquals(cs[0].ext, { row: 1, col: 1, rowspan: 2, colspan: 1, depth: 2 })
  assertEquals(cs[0].children?.length, 1)
  assertEquals(cs[0].children?.[0]?.ext, { row: 2, col: 1, rowspan: 1, colspan: 1, depth: 1 })

  // case 3
  genHeadColumnExtParams(cs = [{ children: [{}, {}] }], 1)
  assertEquals(cs.length, 1)
  assertEquals(cs[0].ext, { row: 1, col: 1, rowspan: 2, colspan: 2, depth: 2 })
  assertEquals(cs[0].children?.length, 2)
  assertEquals(cs[0].children?.[0]?.ext, { row: 2, col: 1, rowspan: 1, colspan: 1, depth: 1 })
  assertEquals(cs[0].children?.[1]?.ext, { row: 2, col: 2, rowspan: 1, colspan: 1, depth: 1 })

  // case 4
  genHeadColumnExtParams(cs = [{}, { children: [{}, {}] }], 1)
  assertEquals(cs.length, 2)
  assertEquals(cs[0].ext, { row: 1, col: 1, rowspan: 2, colspan: 1, depth: 1 })
  assertEquals(cs[1].ext, { row: 1, col: 2, rowspan: 2, colspan: 2, depth: 2 })
  assertEquals(cs[1].children?.length, 2)
  assertEquals(cs[1].children?.[0]?.ext, { row: 2, col: 2, rowspan: 1, colspan: 1, depth: 1 })
  assertEquals(cs[1].children?.[1]?.ext, { row: 2, col: 3, rowspan: 1, colspan: 1, depth: 1 })

  // case 5
  genHeadColumnExtParams(cs = [{}, { children: [{}, { children: [{}, {}] }] }], 1)
  assertEquals(cs.length, 2)
  assertEquals(cs[0].ext, { row: 1, col: 1, rowspan: 3, colspan: 1, depth: 1 })
  assertEquals(cs[1].ext, { row: 1, col: 2, rowspan: 3, colspan: 3, depth: 3 })
  assertEquals(cs[1].children?.length, 2)
  assertEquals(cs[1].children?.[0]?.ext, { row: 2, col: 2, rowspan: 2, colspan: 1, depth: 1 })
  assertEquals(cs[1].children?.[1]?.ext, { row: 2, col: 3, rowspan: 2, colspan: 2, depth: 2 })
  assertEquals(cs[1].children?.[1]?.children?.length, 2)
  assertEquals(cs[1].children?.[1]?.children?.[0]?.ext, { row: 3, col: 3, rowspan: 1, colspan: 1, depth: 1 })
  assertEquals(cs[1].children?.[1]?.children?.[1]?.ext, { row: 3, col: 4, rowspan: 1, colspan: 1, depth: 1 })
})

Deno.test('gen simple table', async () => {
  // define head-column
  const headColumns: HeadColumn[] = [
    { id: 'name', label: 'Name', width: 15, value: (row) => `${row.firstName} ${row.lastName}` },
    { id: 'date', width: 12, dataCellStyle: { numFmt: 'yyyy-MM-dd', alignment: { horizontal: 'center' } } },
    { id: 'int', label: 'Int', dataCellStyle: { numFmt: '#', alignment: { horizontal: 'right' } } },
    { id: 'decimal', label: 'Decimal', dataCellStyle: { numFmt: '#0.00', alignment: { horizontal: 'right' } } },
    {
      id: 'money',
      label: 'Money',
      width: 22,
      dataCellStyle: { numFmt: '￥#,###,###,##0.00', alignment: { horizontal: 'right' } },
    },
    { id: 'percent', label: 'Percent', dataCellStyle: { numFmt: '0.00%', alignment: { horizontal: 'right' } } },
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
  console.log(`see ${file}`)
})

Deno.test('gen table with group', async () => {
  // define head-column
  const headColumns: HeadColumn[] = [
    { id: 'teacher', label: 'Teacher', width: 15, value: (row) => `${row.firstName} ${row.lastName}` },
    {
      id: 'workdate',
      label: 'Workdate',
      width: 12,
      dataCellStyle: { numFmt: 'yyyy-MM-dd', alignment: { horizontal: 'center' } },
    },
    {
      label: 'Students',
      children: [
        { pid: 'students', id: 'name', label: 'Name' },
        {
          pid: 'students',
          id: 'birthdate',
          label: 'Birthdate',
          width: 12,
          dataCellStyle: { numFmt: 'yyyy-MM-dd', alignment: { horizontal: 'center' } },
        },
      ],
    },
    {
      label: 'Favorites',
      children: [
        { pid: 'favorites', id: 'name', label: 'Name' },
        {
          pid: 'favorites',
          id: 'priority',
          label: 'Priority',
          width: 12,
          dataCellStyle: { numFmt: '#', alignment: { horizontal: 'right' } },
        },
      ],
    },
  ]

  // define data-row
  // deno-lint-ignore no-explicit-any
  const dataRows: Record<string, any>[] = [
    { firstName: 'John', lastName: 'Smith', workdate: '2000-01-01' },
    {
      firstName: 'Li',
      lastName: 'Xiao',
      students: [
        { name: 'Lili', birthdate: '2020-01-01' },
        { name: 'Suson', birthdate: '2023-12-01' },
      ],
      favorites: [
        { priority: 1, name: 'Draw' },
        { priority: 2, name: 'Math' },
        { priority: 3, name: 'Sport' },
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
  console.log(`see ${file}`)
})
