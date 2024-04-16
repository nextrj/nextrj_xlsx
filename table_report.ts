import { ExcelJS, recursiveAssign } from './deps.ts'

// #region for type definition

export type CellStyle = Partial<ExcelJS.Style>
export type DataColumn = DirectDataColumn | CascadeDataColumn
export type HeadColumn = GroupColumn | DataColumn

export type ColumnShareProperties = {
  label?: string
  width?: number
  cellStyle?: CellStyle
  headCellStyle?: CellStyle
  dataCellStyle?: CellStyle
  /** For inner usage */
  ext?: ColumnExtProperties
}

export type DirectDataColumn = ColumnShareProperties & {
  key: string
  mapper?: ValueMapper
}

export type CascadeDataColumn = Omit<DirectDataColumn, 'key'> & {
  keys: string[]
}

export type GroupColumn = ColumnShareProperties & {
  label?: string
  children: HeadColumn[]
}

/** The inner properties of the column */
export type ColumnExtProperties = {
  row: number
  col: number
  rowspan: number
  colspan: number
  depth: number
  /** the head cell style combined all ancestor and self */
  headCellStyle?: CellStyle
  /** the data cell style combined all ancestor and self */
  dataCellStyle?: CellStyle
}

/** The inner properties of the column */
export type DataRowExtProperties = {
  row: number
  rowspan: number
}

/** Data row of the table */
// deno-lint-ignore no-explicit-any
export type DataRow = Record<string, any> & {
  /** For inner usage */
  ext?: DataRowExtProperties
}

export type WorkbookProperties = {
  subject?: string
  title?: string
  company?: string
  creator?: string
  created?: Date
  lastModifiedBy?: string
  modified?: Date
  description?: string
}

export type WorksheetOptions = {
  name?: string
  properties?: Partial<ExcelJS.WorksheetProperties>
  pageSetup?: Partial<ExcelJS.PageSetup>
  views?: Partial<ExcelJS.WorksheetView>[]
  table?: Table
}

export type Table = {
  headColumns: HeadColumn[]
  dataRows?: DataRow[]
  caption?: string | { value: string; style?: CellStyle }
  subCaption?: string | { value: string; style?: CellStyle }
  cellStyle?: CellStyle
  headCellStyle?: CellStyle
  dataCellStyle?: CellStyle
}

/** Map the origin value to anthor value */
// deno-lint-ignore no-explicit-any
export type ValueMapper = ({ value, index, row }: { value: any; index: number; row: DataRow }) => any

/** Define something like `{a: {b: {c: 0}}, d: 0}`
 * - Note: `0` means the leaf
 */
export type CascadeKey = { [key: string]: 0 | CascadeKey }

// #endregion

// #region for gen workbook

export async function genSingleSheetWorkbook({
  headColumns,
  dataRows,
  caption,
  subCaption,
  cellStyle,
  headCellStyle,
  dataCellStyle,
  sheetName = 'Sheet1',
  sheetProperties,
  sheetPageSetup,
  sheetView,
  bookProperties,
}: {
  headColumns: HeadColumn[]
  dataRows: DataRow[]
  caption?: string | { value: string; style?: CellStyle }
  subCaption?: string | { value: string; style?: CellStyle }
  cellStyle?: CellStyle
  headCellStyle?: CellStyle
  dataCellStyle?: CellStyle
  sheetName?: string
  sheetProperties?: Partial<ExcelJS.WorksheetProperties>
  sheetPageSetup?: Partial<ExcelJS.PageSetup>
  sheetView?: Partial<ExcelJS.WorksheetView>
  bookProperties?: WorkbookProperties
}): Promise<ExcelJS.Workbook> {
  return await genWorkbook([{
    name: sheetName,
    properties: sheetProperties,
    pageSetup: sheetPageSetup,
    views: sheetView ? [sheetView] : [],
    table: {
      headColumns,
      dataRows,
      caption,
      subCaption,
      cellStyle,
      headCellStyle,
      dataCellStyle,
    },
  }], {
    bookProperties: bookProperties,
  })
}

export async function genWorkbook(sheets: WorksheetOptions[], { bookProperties }: {
  bookProperties?: WorkbookProperties
} = {}): Promise<ExcelJS.Workbook> {
  // verify
  if (!sheets?.length) throw new Error('sheets could not be null or empty.')

  const wb = new ExcelJS.Workbook()

  // set book properties
  if (bookProperties) {
    if (Object.hasOwn(bookProperties, 'subject')) wb.subject = bookProperties.subject as string
    if (Object.hasOwn(bookProperties, 'title')) wb.title = bookProperties.title as string
    if (Object.hasOwn(bookProperties, 'company')) wb.company = bookProperties.company as string
    if (Object.hasOwn(bookProperties, 'creator')) wb.creator = bookProperties.creator as string
    else wb.creator = 'NextRJ'
    if (Object.hasOwn(bookProperties, 'created')) wb.created = bookProperties.created as Date
    if (Object.hasOwn(bookProperties, 'lastModifiedBy')) wb.lastModifiedBy = bookProperties.lastModifiedBy as string
    if (Object.hasOwn(bookProperties, 'modified')) wb.modified = bookProperties.modified as Date
    if (Object.hasOwn(bookProperties, 'description')) wb.description = bookProperties.description as string
    else wb.description = 'Power by https://deno.land/x/nextrj_xlsx'
  } else {
    wb.creator = 'NextRJ'
    wb.description = 'Power by https://deno.land/x/nextrj_xlsx'
  }

  // gen sheet
  sheets.forEach((sheet, index) => {
    // initial sheet
    const ws = wb.addWorksheet(sheet.name || `NONAME${index + 1}`, {
      properties: sheet.properties,
      pageSetup: sheet.pageSetup,
      views: sheet.views,
    })

    // gen sheet table
    const table = sheet.table
    if (table) {
      let startRow = 1

      // gen caption
      if (table.caption) {
        const cell = ws.getCell(startRow++, 1)
        if (typeof table.caption === 'string') cell.value = table.caption
        else {
          cell.value = table.caption.value
          if (table.caption.style) cell.style = table.caption.style
        }
      }

      // gen subCaption
      if (table.subCaption) {
        const cell = ws.getCell(startRow++, 1)
        if (typeof table.subCaption === 'string') cell.value = table.subCaption
        else {
          cell.value = table.subCaption.value
          if (table.subCaption.style) cell.style = table.subCaption.style
        }
      }

      // gen all head-column cells
      if (table.headColumns?.length) {
        const headCellStyle: CellStyle = table.cellStyle || table.headCellStyle
          ? recursiveAssign({}, table.cellStyle || {}, table.headCellStyle || {})
          : undefined
        const dataCellStyle: CellStyle = table.cellStyle || table.dataCellStyle
          ? recursiveAssign({}, table.cellStyle || {}, table.dataCellStyle || {})
          : undefined
        const range = genAllColumnHeadCells(table.headColumns, ws, { startRow, headCellStyle, dataCellStyle })

        // gen all data-row cells
        if (table.dataRows?.length) {
          genAllDataRowCells(table.dataRows, ws, {
            headColumns: table.headColumns,
            startRow: range.bottom + 1,
          })
        } else {
          // no data-row
          const cell = ws.getCell(startRow, 1)
          cell.value = 'NO_DATA'
          if (dataCellStyle) cell.style = dataCellStyle
          if (range.right > range.left) {
            ws.mergeCells({
              top: range.top + 1,
              left: range.left,
              bottom: range.top + 1,
              right: range.right,
            })
          }
        }
      } else {
        // no head-column
        ws.getCell('A1').value = 'NO_HEAD_COLUMN'
      }
    } else {
      // no table
      ws.getCell('A1').value = 'NO_TABLE'
    }
  })

  // return
  return await Promise.resolve(wb)
}

// #endregion

// #region for HeadColumn

function genAllColumnHeadCells(
  headColumns: HeadColumn[],
  ws: ExcelJS.Worksheet,
  { startRow = 1, headCellStyle, dataCellStyle }: {
    startRow?: number
    headCellStyle?: CellStyle
    dataCellStyle?: CellStyle
  } = {},
): ExcelJS.Location {
  // gen column ext-properties
  genColumnExtProperties(headColumns, { startRow, headCellStyle, dataCellStyle })

  // gen head-cells
  if (!headColumns?.length) throw new Error('headColumns could not be null or empty.')
  recursiveGenColumnHeadCell(headColumns, ws)

  return {
    top: startRow,
    bottom: startRow + headColumns[0]!.ext!.rowspan - 1,
    left: 1,
    right: headColumns[headColumns.length - 1]!.ext!.col + headColumns[headColumns.length - 1]!.ext!.colspan - 1,
  }
}

/** Auto gen column ext-properties */
export function genColumnExtProperties(headColumns: HeadColumn[], { startRow = 1, headCellStyle, dataCellStyle }: {
  startRow?: number
  headCellStyle?: CellStyle
  dataCellStyle?: CellStyle
} = {}): void {
  // recursive gen ext properties exclude rowspan
  let nextCol = 1
  headColumns.forEach((column) => {
    // combine head cell style
    const columnHeadCellStyle = headCellStyle || column.cellStyle || column.headCellStyle
      ? recursiveAssign({}, headCellStyle || {}, column.cellStyle || {}, column.headCellStyle || {})
      : undefined

    // combine data cell style
    const columnDataCellStyle = dataCellStyle || column.cellStyle || column.dataCellStyle
      ? recursiveAssign({}, dataCellStyle || {}, column.cellStyle || {}, column.dataCellStyle || {})
      : undefined

    const ext = recursiveGenColumnExtProperties(column, {
      row: startRow,
      col: nextCol,
      headCellStyle: columnHeadCellStyle,
      dataCellStyle: columnDataCellStyle,
    })
    nextCol += ext.colspan
  })

  // calc maxDepth
  const depth = Math.max(...headColumns.map((column) => column.ext!.depth))

  // recursive gen rowspan
  headColumns.forEach((column) => recursiveGenColumnRowspan(column, depth))
}

/** Recursive gen ext-properties exclude rowspan */
function recursiveGenColumnExtProperties(
  column: HeadColumn,
  { row, col, headCellStyle, dataCellStyle }: {
    row: number
    col: number
    headCellStyle?: CellStyle
    dataCellStyle?: CellStyle
  },
): ColumnExtProperties {
  const ext = column.ext ??= { row, col, rowspan: 1, colspan: 1, depth: 1 }
  if (headCellStyle) ext.headCellStyle = headCellStyle
  if (dataCellStyle) ext.dataCellStyle = dataCellStyle

  if (Object.hasOwn(column, 'children')) { // GroupColumn
    const groupColumn = column as GroupColumn
    let nextCol = col

    // recursive deal all child column
    const ext1s = groupColumn.children?.map((column) => {
      // combine head cell style
      const columnHeadCellStyle = ext.headCellStyle || column.cellStyle || column.headCellStyle
        ? recursiveAssign({}, ext.headCellStyle || {}, column.cellStyle || {}, column.headCellStyle || {})
        : undefined

      // combine data cell style
      const columnDataCellStyle = ext.dataCellStyle || column.cellStyle || column.dataCellStyle
        ? recursiveAssign({}, ext.dataCellStyle || {}, column.cellStyle || {}, column.dataCellStyle || {})
        : undefined

      const ext1 = recursiveGenColumnExtProperties(column, {
        row: row + 1,
        col: nextCol,
        headCellStyle: columnHeadCellStyle,
        dataCellStyle: columnDataCellStyle,
      })
      nextCol += ext1.colspan
      return ext1
    })

    // colspan = sum(child-colspan)
    ext.colspan = nextCol - col

    // depth = 1 + maxChildDepth
    ext.depth = 1 + Math.max(...ext1s.map((t) => t.depth))
  }

  return ext
}

/** Recursive gen rowspan */
function recursiveGenColumnRowspan(column: HeadColumn, rowspan: number): void {
  column.ext!.rowspan = rowspan
  if (Object.hasOwn(column, 'children')) { // GroupColumn
    const groupColumn = column as GroupColumn
    groupColumn.children?.forEach((column) => recursiveGenColumnRowspan(column, rowspan - 1))
  }
}

/** recursive gen head cells */
function recursiveGenColumnHeadCell(
  columns: HeadColumn[],
  ws: ExcelJS.Worksheet,
): void {
  columns.forEach((column) => {
    const ext = column.ext!
    // set cell value
    const row = ext.row
    const col = ext.col
    const cell = ws.getCell(row, col)
    cell.value = column.label || (Object.hasOwn(column, 'key') ? (column as DirectDataColumn).key : 'NONAME')

    // set column width
    if (column.width) ws.getColumn(col).width = column.width

    // set cell style
    if (ext.headCellStyle) cell.style = ext.headCellStyle

    // merge cell if necessary
    const colspan = ext.colspan
    const rowspan = Object.hasOwn(column, 'children') ? 1 : ext.rowspan
    if (colspan > 1 || rowspan > 1) {
      const l = {
        top: row,
        left: col,
        bottom: row + rowspan - 1,
        right: col + colspan - 1,
      }
      ws.mergeCells(l)
    }

    // recursive gen child-column-cell
    if (Object.hasOwn(column, 'children')) {
      recursiveGenColumnHeadCell((column as GroupColumn).children, ws)
    }
  })
}

// #endregion

// #region for DataRow

/**
 * Generate rows line by line.
 * Note: Need to call after invoke genColumnHeadCell.
 */
function genAllDataRowCells(dataRows: DataRow[], ws: ExcelJS.Worksheet, { headColumns, startRow }: {
  headColumns: HeadColumn[]
  startRow: number
}): void {
  // no data not render
  if (!dataRows?.length) return

  // get all DataColumn
  const dataColumns = flattenColumnByChildren(headColumns)
    .filter((c) => !Object.hasOwn(c, 'children')) as DataColumn[]

  // get all max-depth cascade keys
  const arrayKeys: string[][] = dataColumns.filter((c) => Object.hasOwn(c, 'keys'))
    .map((c) => {
      const keys = [...(c as CascadeDataColumn).keys]
      // delete last key because it is not a cascade key
      keys.pop()
      return keys
    })
  const cascadeKey = convertCacadeArrayKey2ObjectKey(arrayKeys)

  // gen data-row ext-properties
  recursiveGenDataRowExtProperties(dataRows, { cascadeKey, startRow })

  // gen data-row-cells line by line
  dataRows?.forEach((dataRow, index) => genDataRowCells(dataRow, dataColumns, { ws, index }))
}

/** Flatten column by children.
 * - `[{a:1, children: [{b:1, children: [{c:1}]}]}}]` => `[{{a:1, ...}, {b:1, ...}, {c:1}}]`
 */
export function flattenColumnByChildren(columns: HeadColumn[]): HeadColumn[] {
  return columns.reduce((pre, cur) => {
    // deno-lint-ignore no-explicit-any
    if ((cur as any)?.children?.length) pre.push(...flattenColumnByChildren((cur as GroupColumn).children))
    return pre
  }, [...columns])
}

/** Coonvert cacade-array-key to cacade-object-key.
 * - For remove duplcate parent key
 * - `[['a']]` to `{ a: 0 }`
 * - `[['a'], ['a', 'b']]` to `{ a: { b: 0 } }`
 * - `[['a', 'b'], ['a']]` to `{ a: { b: 0 } }`
 * - `[['a', 'b'], ['c'], ['a']]` to `{ a: { b: 0 }, c: 0 }`
 * - `[['a', 'b', 'c'], ['d'], ['a', 'b']]` to `{ a: { b: { c: 0 } }, d: 0 }`
 */
export function convertCacadeArrayKey2ObjectKey(arrayKeys: string[][]): CascadeKey {
  const cd: CascadeKey = {}
  arrayKeys.forEach((keys) => recursiveConvertCacadeArrayKey2ObjectKey(keys, cd))
  return cd
}

function recursiveConvertCacadeArrayKey2ObjectKey(keys: string[], cd: CascadeKey) {
  if (keys.length > 1) {
    const nestedKeys = [...keys]
    const key = nestedKeys.shift()!
    const nextedCd = (cd[key] ? cd[key] : cd[key] = {}) as CascadeKey
    recursiveConvertCacadeArrayKey2ObjectKey(nestedKeys, nextedCd)
  } else {
    if (!Object.hasOwn(cd, keys[0])) cd[keys[0]] = 0
  }
}

/** Recursive gen `DataRow.ext`.
 * @returns the next row number
 */
export function recursiveGenDataRowExtProperties(
  dataRows: DataRow[],
  { cascadeKey = {}, startRow = 1, depth = 0, debug = false }: {
    cascadeKey?: CascadeKey
    startRow?: number
    depth?: number
    debug?: boolean
  } = {},
): number {
  const sp = ' '.repeat(depth * 2)
  // cache cascadeKey's key-value pairs
  const cascadeKVs = Object.entries(cascadeKey)
  if (debug) {
    console.log(
      `${sp}depth=${depth}, startRow=${startRow}, dataRowLen=${dataRows.length}, cascadeKey=${
        JSON.stringify(cascadeKey)
      }`,
    )
  }

  // deal each data-row
  let nextRow = startRow
  dataRows.forEach((dataRow, index) => {
    if (debug) console.log(`${sp}index=${index}, dataRow=${JSON.stringify(dataRow)}`)
    // initial `dataRow.ext`
    const ext = dataRow.ext ??= { row: nextRow, rowspan: 1 }

    if (!cascadeKVs.length) {
      // no cascade-key treat as take up one row
      nextRow++
    } else {
      let nestedNextRow = nextRow
      // gen each cascade-key-value's `dataRow.ext`
      const nestedNextRows = cascadeKVs.map(([key, nestedCascadeKey]) => {
        const nestedDataRows = dataRow[key]
        if (debug) {
          console.log(
            `${sp}key=${key}, row=${nestedNextRow}, cascadeKey=${JSON.stringify(nestedCascadeKey)}, dataRows=${
              JSON.stringify(nestedDataRows)
            }`,
          )
        }
        if (nestedDataRows) {
          // verify it must to be array value
          if (!Array.isArray(nestedDataRows)) {
            throw new Error(`Value type illegal: the value of dataRow["${key}"] is not a array value.`)
          }

          if (nestedDataRows.length) { // has nested data
            if (nestedCascadeKey) {
              // recursive gen nested-cascade-key-value's `dataRow.ext`
              nestedNextRow = recursiveGenDataRowExtProperties(
                nestedDataRows,
                { cascadeKey: nestedCascadeKey, startRow: nestedNextRow, depth: depth + 1, debug },
              )
              return nestedNextRow
            } else {
              // no nested-cascade-key treat as normal rows
              nestedDataRows.forEach((dataRow, index) => dataRow.ext ??= { row: nextRow + index, rowspan: 1 })
              return nextRow + nestedDataRows.length
            }
          } else { // empty nested data treat as empty row and take up one row
            return nestedNextRow + 1
          }
        } else {
          // no nested data treat as empty row and take up one row
          return nestedNextRow + 1
        }
      })

      // calc max nextRow of same level
      if (debug) console.log(`${sp}nextRows=${JSON.stringify(nestedNextRows)}`)
      nextRow = Math.max(...nestedNextRows)
      ext.rowspan = nextRow - ext.row
    }
  })
  return nextRow
}

function _recursiveGenDataRowExtProperties(
  dataRow: DataRow,
  cascadeKey: CascadeKey,
  startRow: number,
): DataRowExtProperties {
  // initial `dataRow.ext`
  const ext = dataRow.ext ??= { row: startRow, rowspan: 1 }

  // gen each root-key-value's `dataRow.ext`
  const exts = Object.entries(cascadeKey).map(([key, nestedCascadeKey]) => {
    const nestedDataRows = dataRow[key]
    if (nestedDataRows) {
      // verify it must to be array value
      if (!Array.isArray(nestedDataRows)) {
        throw new Error(`Value type illegal: the value of dataRow["${key}"] is not a array value.`)
      }

      // recursive gen next-level-key-value's `dataRow.ext`
      if (nestedCascadeKey) {
        let nextRow = startRow
        nestedDataRows.forEach((dataRow) => {
          const nestedExt = _recursiveGenDataRowExtProperties(dataRow, nestedCascadeKey, nextRow)
          nextRow += nestedExt.rowspan
          return nestedExt
        })
        return { row: startRow, rowspan: nextRow - startRow, depth: 1 }
      } else {
        return { row: startRow, rowspan: 1, depth: 1 }
      }
    } else {
      return { row: startRow, rowspan: 1, depth: 1 }
    }
  })

  // calc max rowspan of same level
  ext.rowspan = Math.max(...exts.map((ext) => ext.rowspan))

  return ext
}

function genDataRowCells(
  dataRow: DataRow,
  dataColumns: DataColumn[],
  { ws, index }: { ws: ExcelJS.Worksheet; index: number },
): void {
  dataColumns.forEach((c) => {
    recursiveGenDataRowCell(
      Object.hasOwn(c, 'keys') ? (c as CascadeDataColumn).keys : [(c as DirectDataColumn).key],
      dataRow,
      {
        ws,
        index,
        row: dataRow.ext!.row,
        col: c.ext!.col,
        style: c.ext!.dataCellStyle,
        mapper: c.mapper,
      },
    )
  })
}

/** Gen column data-cells for `keys=['k']` or `keys=['a', 'b', ..., 'k']`.
 * @returns The next row number
 */
export function recursiveGenDataRowCell(
  keys: string[],
  // deno-lint-ignore no-explicit-any
  data: Record<string, any>,
  { ws, index, row, col, style, mapper }: {
    ws: ExcelJS.Worksheet
    index: number
    row: number
    col: number
    style?: CellStyle
    mapper?: ValueMapper
  },
): number {
  if (keys.length === 1) {
    const key = keys[0]
    const value = mapper ? mapper({ value: data[key], index: index, row: data }) : data[key]
    genDataCell({ ws, row, col, value, style })

    const rowspan = data.ext?.rowspan || 1
    if (rowspan > 1) {
      // merge cell
      ws.mergeCells({
        top: row,
        left: col,
        bottom: row + rowspan - 1,
        right: col,
      })
      // return next row
      return row + rowspan
    } else {
      // return next row
      return row + 1
    }
  } else {
    const keysClone = [...keys]
    const key = keysClone.shift()!
    const value = data[key] //|| [{}] // default single empty object array for write empty cell
    let nextRow = row
    if (value) {
      ;(value as Record<string, unknown>[]).forEach((data, index) => {
        nextRow = recursiveGenDataRowCell(keysClone, data, { ws, index, row: nextRow, col, style, mapper })
      })
    }
    return nextRow
  }
}

function genDataCell(
  { ws, row, col, value, style }: {
    ws: ExcelJS.Worksheet
    row: number
    col: number
    // deno-lint-ignore no-explicit-any
    value?: any
    style?: CellStyle
  },
): void {
  const cell = ws.getCell(row, col)
  cell.value = value
  if (style) cell.style = style
}

// #endregion
