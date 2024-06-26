# nextrj_xlsx changelog

## 0.5.0 2024-04-17

- Refactor `ValueMapper` signature and rename `HeadColumn.value` to `.mapper`
  ```json
  { key: 'name', label: 'Name', width: 15, mapper: ({ row }) =>`${row.firstName} ${row.lastName}`}
  ```
- Fixed missing-rows style

## 0.4.0 2024-04-15

- Totally refactor to support multiple nested data-row
  - Change type `ValueMapper` sinature
  - Rename `HeadColumn.Mapper` to `HeadColumn.value`
  - Add more type definition

- Uncommit .DS_Store

## 0.3.0 2024-04-04

- Change type `ValueMapper` sinature

## 0.2.0 2024-04-03

- Change to use ExcelJS namespace directly
- Rename `HeadColumn.value` to `HeadColumn.valueMapper`

## 0.1.0 2024-04-03

- Initial table report Implements
