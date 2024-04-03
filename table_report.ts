import { ExcelJS, PageSetup, recursiveAssign, Style, Worksheet, WorksheetProperties, WorksheetView } from './deps.ts'

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
  properties?: Partial<WorksheetProperties>
  pageSetup?: Partial<PageSetup>
  views?: Partial<WorksheetView>[]
  table?: Table
}
export type Table = {
  headColumns: HeadColumn[]
  // deno-lint-ignore no-explicit-any
  dataRows: Record<string, any>[]
  caption?: string | { value: string; style?: Partial<Style> }
  subCaption?: string | { value: string; style?: Partial<Style> }
  cellStyle?: Partial<Style>
  headCellStyle?: Partial<Style>
  dataCellStyle?: Partial<Style>
}
export type HeadColumn = {
  id?: string
  label?: string
  width?: number
  value?: ValueMapper
  cellStyle?: Partial<Style>
  headCellStyle?: Partial<Style>
  dataCellStyle?: Partial<Style>
  pid?: string
  children?: HeadColumn[]
  /** The inner ext params */
  ext?: HeadColumnExtParams
}
/** Map the origin value to anthor value */
// deno-lint-ignore no-explicit-any
export type ValueMapper = (row: Record<string, any>) => any
type HeadColumnExtParams = {
  row: number
  col: number
  rowspan: number
  colspan: number
  depth: number
}

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
  // deno-lint-ignore no-explicit-any
  dataRows: Record<string, any>[]
  caption?: string | { value: string; style?: Partial<Style> }
  subCaption?: string | { value: string; style?: Partial<Style> }
  cellStyle?: Partial<Style>
  headCellStyle?: Partial<Style>
  dataCellStyle?: Partial<Style>
  sheetName?: string
  sheetProperties?: Partial<WorksheetProperties>
  sheetPageSetup?: Partial<PageSetup>
  sheetView?: Partial<WorksheetView>
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

      // gen column ext-params
      genHeadColumnExtParams(table.headColumns, startRow)

      // gen head-cells
      if (!table.headColumns.length) throw new Error('headColumns could not be null or empty.')
      const tableHeadCellStyle: Partial<Style> = table.cellStyle || table.headCellStyle
        ? recursiveAssign({}, table.cellStyle || {}, table.headCellStyle || {})
        : undefined
      recursiveGenHeadCell(table.headColumns, ws, tableHeadCellStyle)

      // gen data-cells
      genDataRow(table, ws)
    } else {
      // no data
      ws.getCell('A1').value = 'NODATA'
    }
  })

  // return
  return await Promise.resolve(wb)
}

/** Auto gen column ext-params */
export function genHeadColumnExtParams(columns: HeadColumn[], startRow: number): HeadColumn[] {
  // recursive gen ext params exclude rowspan
  let nextCol = 1
  columns.forEach((column) => {
    const ext = recursiveGenHeadColumnExtParams(column, startRow, nextCol)
    nextCol += ext.colspan
  })

  // calc maxDepth
  const depth = Math.max(...columns.map((column) => column.ext!.depth))

  // recursive gen rowspan
  columns.forEach((column) => recursiveGenRowspan(column, depth))

  return columns
}

/** Recursive gen ext params exclude rowspan */
function recursiveGenHeadColumnExtParams(column: HeadColumn, row: number, col: number): HeadColumnExtParams {
  const ext = column.ext ??= { row, col, rowspan: 1, colspan: 1, depth: 1 }

  if (column.children?.length) {
    let nextCol = col

    // recursive deal all child column
    const ext1s = column.children.map((column) => {
      const ext1 = recursiveGenHeadColumnExtParams(column, row + 1, nextCol)
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
function recursiveGenRowspan(column: HeadColumn, rowspan: number): void {
  column.ext!.rowspan = rowspan
  if (column.children?.length) {
    column.children.forEach((column) => recursiveGenRowspan(column, rowspan - 1))
  }
}

/** recursive gen head cells */
function recursiveGenHeadCell(columns: HeadColumn[], ws: Worksheet, parentHeadCellStyle?: Partial<Style>): void {
  columns.forEach((column) => {
    // set cell value
    const row = column.ext!.row
    const col = column.ext!.col
    const cell = ws.getCell(row, col)
    cell.value = column.label || column.id || undefined

    // set column width
    if (column.width) ws.getColumn(column.ext!.col).width = column.width

    // set cell style
    const columnHeadCellStyle = parentHeadCellStyle || column.cellStyle || column.headCellStyle
      ? recursiveAssign({}, parentHeadCellStyle || {}, column.cellStyle || {}, column.headCellStyle || {})
      : undefined
    if (columnHeadCellStyle) cell.style = columnHeadCellStyle

    // merge cell if necessary
    const colspan = column.ext!.colspan
    const rowspan = column.children?.length ? 1 : column.ext!.rowspan
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
    if (column.children?.length) recursiveGenHeadCell(column.children, ws, columnHeadCellStyle)
  })
}

function genDataRow(table: Table, ws: Worksheet) {
  const tableDataCellStyle = table.cellStyle || table.dataCellStyle
    ? recursiveAssign({}, table.cellStyle || {}, table.dataCellStyle)
    : undefined
  table.headColumns?.forEach((column) => {
    const ext = column.ext!
    let nextRow = ext.row + ext.rowspan
    const nextCol = ext.col
    const columnDataCellStyle = tableDataCellStyle || column.cellStyle || column.dataCellStyle
      ? recursiveAssign({}, tableDataCellStyle || {}, column.cellStyle || {}, column.dataCellStyle || {})
      : undefined
    table.dataRows?.forEach((dataRow) => {
      // the max length of all array value
      const rowspan = Math.max(
        1,
        ...Object.values(dataRow).filter((v) => Array.isArray(v)).map((v) => v.length as number),
      )
      if (!column.children) {
        // set cell value
        const cell = ws.getCell(nextRow, nextCol)
        if (column.value) cell.value = column.value(dataRow)
        else if (column.id) cell.value = dataRow[column.id]

        // set cell style
        if (columnDataCellStyle) cell.style = columnDataCellStyle

        // merge cell if necessary
        if (rowspan > 1) {
          const l = {
            top: nextRow,
            bottom: nextRow + rowspan - 1,
            left: nextCol,
            right: nextCol,
          }
          ws.mergeCells(l)
        }
      } else { // deal children column cell
        column.children?.forEach((childColumn) => {
          const childColumnDataCellStyle = columnDataCellStyle || childColumn.cellStyle || childColumn.dataCellStyle
            ? recursiveAssign(
              {},
              columnDataCellStyle || {},
              childColumn.cellStyle || {},
              childColumn.dataCellStyle || {},
            )
            : undefined
          for (let i = 0; i < rowspan; i++) {
            // set cell value
            const pid = childColumn.pid || column.id
            const childDataRow = pid ? dataRow[pid]?.[i] : dataRow
            const cell = ws.getCell(nextRow + i, childColumn.ext?.col!)
            if (childColumn.value) cell.value = childColumn.value(childDataRow)
            else if (childColumn.id) cell.value = childDataRow?.[childColumn.id!]

            // set cell style
            if (childColumnDataCellStyle) cell.style = childColumnDataCellStyle
          }
        })
      }
      nextRow += rowspan
    })
  })
}
