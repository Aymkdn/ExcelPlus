# Changelog

**Change Log v2.5 (April 14, 2017)**

  - Add compatibility with the last js-xlsx releases (see https://github.com/Aymkdn/ExcelPlus/issues/8)
  - Add option `propertiesOnly` for `read`
  - Add option `compression` for `saveAs`
  - Add function `open()` (see https://github.com/Aymkdn/ExcelPlus/issues/13)
  - Fix Angular compatibility issue (see https://github.com/Aymkdn/ExcelPlus/issues/14)
  - Change `parseDate` option for `read` (see documentation)
  - Change demo pages to always get the latest release of js-xlsx
  - Change license to Apache v2

**Change Log v2.4.1 (June 21, 2016)**

  - Better handle of cells with a Date

**Change Log v2.4 (June 14, 2016)**

  - Add `writeRow()` and `writeNextRow()`
  - Show `write()` into the documentation

**Change Log v2.3 (January 15, 2016)**

  - Fix `writeCell()` when the content was an empty string
  - Fix documentation
  - IE10/11 compatibility
  - Stop using readAsBinaryString to use readAsArrayBuffer
  - Fix button size when using Flash
  - Fix: with IE and Flash object, the Excel was returning column A as null
  - Add same options for both readAll() and read()
  - SheetJS/JS-XLSX 0.8.0 doesn't handle correctly the cell format, so I've changed the file (see https://github.com/SheetJS/js-xlsx/pull/349)
  - SheetJS/JS-XLSX 0.8.0 and IE8 didn't return the first column of a file, so I've fixed it (see https://github.com/SheetJS/js-xlsx/issues/350)
  - Fix for the read()/readAll() when `parseDate` is true, it should now recognize a date cell
  - Add options `parameters` as a boolean for read()/readAll()

**Change Log v2.2 (December 23, 2014)**

  - Rewrite everything to use the XLSX-JS library
  - Official release
