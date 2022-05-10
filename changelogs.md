* 0.4.4:
    * Add support for reset merging cells and copy the subsequent cells to avoid sharing.
    * Add support for merging cells and sharing all its style.
    * Refactor getCell and addCell to allow adding empty cell and different style.
    * Refactor local test.
* 0.4.3:
    * Add support for fill and fetch hyperlink cell.
* 0.4.2:
    * Add ranges= and autoFilter proc to sheet.
    * Add `createdAt` to set excel creation date properties.
* 0.4.1:
    * Add `shareStyle` API for easy referring the same style.
    * Add `copyStyle` API for easy copying the same style.
* 0.4.0:
    * Add cell styles API.
    * Add row display API.
    * Remove deprecated `addRow`.
    * Fix mispoint when adding new shared strings
    * Change internal shared strings implementation.
    * Fix created properties when calling newExcel.
* 0.3.6:
    * Fix unreaded embed files when reading excel.
* 0.3.5:
    * Fix unreaded .rels when reading excel.
* 0.3.4:
    * Assign default values first when fetching cell data.
    * Remove unnecessary checking index 0 when converting col string to number in toNum.
    * Add `clear` to remove all existing cells.
    * Fix default value for `getCell` when fetching `DateTime` or `Time`
* 0.3.3:
    * Fix toCol and toNum when the cell number or col string is more than 701 or "ZZ".
* 0.3.2:
    * Fix row assignment when no rows are available in new sheet.
* 0.3.1:
    * Add support for filling and fetching formula cell.
    * Support fetching filled cells row.
* 0.3.0:
    * Export helpers `toCol` and `toNum`.
    * Implement row cells fill whether sparse or filled.
    * Modify to make empty.xlsx template smaller.
    * Make `unixSep` template as private as it should.
    * Enhance checking/adding shared string instead of adding it continuously.
    * Make string entry as inline string when it's small string, less than 64 characters.
    * Enforce column to be upper case when accessing cell in row.
* 0.2.2:
    * Add `getCellIt` to access the string value directly as `it`.
    * Change temporary file name and hashed it when calling `$`.
    * Refactor `getCell`.
    * Add unit test.
* 0.2.1:
    * Refactor `[]=` and fix addSheet failed to get last id and returning nil.
* 0.2.0
  * Deprecate `addRow proc(Sheet): Row` and `addRow proc(Sheet, Positive): Row` in favor of `row proc(Sheet, Positive): Row`
  * Add example working with Excel sheets in readme.
  * Fix `addSheet` path, fix `row` duplicated row entry when existing rows are 0.
  * Add Github action for generating docs page.
* 0.1.0
  * Initial release
