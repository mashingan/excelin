include internal_cells

proc row*(s: Sheet, rowNum: Positive, fill = cfSparse): Row =
  ## Add row by selecting which row number to work with.
  ## This will return new row if there's no existing row
  ## or will return an existing one.
  let sdata = s.body.retrieveChildOrNew "sheetData"
  let rowsExists = sdata.len
  if rowNum > rowsExists:
    for i in rowsExists+1 ..< rowNum:
      sdata.add <>row(r= $i, hidden="false", collapsed="false", cellfill= $fill)
  else:
    return Row(sheet:s, body: sdata[rowNum-1])
  result = Row(
    sheet: s,
    body: <>row(r= $rowNum, hidden="false", collapsed="false", cellfill= $fill),
  )
  sdata.add result.body
  s.modifiedAt

proc `hide=`*(row: Row, yes: bool) =
  ## Hide the current row
  row.body.attrs["hidden"] = $(if yes: 1 else: 0)

proc hidden*(row: Row): bool =
  ## Check whether row is hidden
  "1" == row.body.attr "hidden"

proc `height=`*(row: Row, height: Natural) =
  ## Set the row height which sets its attribute to custom height.
  ## If the height 0, will reset its custom height.
  if height == 0:
    for key in ["ht", "customHeight"]:
      row.body.attrs.del key
  else:
    row.body.attrs["customHeight"] = "1"
    row.body.attrs["ht"] = $height

proc height*(row: Row): Natural =
  ## Check the row height if it has custom height and its value set.
  ## If not will by default return 0.
  try: parseInt(row.body.attr "ht") except: 0

proc `outlineLevel=`*(row: Row, level: Natural) =
  ## Set the outline level for the row. Level 0 means resetting the level.
  if level == 0: row.body.attrs.del "outlineLevel"
  else: row.body.attrs["outlineLevel"] = $level

proc outlineLevel*(row: Row): Natural =
  ## Check current row outline level. 0 when it's not outlined.
  try: parseInt(row.body.attr "outlineLevel") except: 0

proc `collapsed=`*(row: Row, yes: bool) =
  ## Collapse the current row, usually used together with outline level.
  if yes: row.body.attrs["collapsed"] = $1
  else: row.body.attrs.del "collapsed"

proc clear*(row: Row) = row.body.clear
  ## Clear all cells in the row.

proc retrieveCell(row: Row, col: string): XmlNode =
  if $cfSparse == row.body.attr "cellfill":
    let colrow = fmt"{col}{row.rowNum}"
    let fetchpos = row.body.fetchCell colrow
    if fetchpos < 0:
      row[col] = ""
      row.body[row.body.len-1]
    else: row.body[fetchpos]
  else:
    row.body[col.toNum]