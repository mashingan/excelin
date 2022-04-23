* 0.2.1:
    * Refactor `[]=` and fix addSheet failed to get last id and returning nil.
* 0.2.0
  * Deprecate `addRow proc(Sheet): Row` and `addRow proc(Sheet, Positive): Row` in favor of `row proc(Sheet, Positive): Row`
  * Add example working with Excel sheets in readme.
  * Fix `addSheet` path, fix `row` duplicated row entry when existing rows are 0.
  * Add Github action for generating docs page.
* 0.1.0
  * Initial release
