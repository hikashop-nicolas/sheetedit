# Real-world fixtures

- `msft-simple-query.xlsx` - the SIMPLE_QUERY_WORKBOOK_TEMPLATE embedded in
  [microsoft/connected-workbooks](https://github.com/microsoft/connected-workbooks)
  (MIT license, Microsoft Corporation), decoded from src/workbookTemplate.ts. A real
  Excel-toolchain workbook whose DataMashup customXml item is UTF-16 LE encoded - the
  encoding detail our first synthetic fixtures missed.
