include internal_cells

proc row*(s: Sheet, rowNum: Positive, fill = cfSparse): Row =
  ## Add row by selecting which row number to work with.
  ## This will return new row if there's no existing row
  ## or will return an existing one.
  let sdata = s.body.retrieveChildOrNew "sheetData"
  template resultRow(rnum: Positive): Row =
    Row(sheet: s,
        body: <>row(r= $rnum, hidden= $false, collapsed= $false, cellfill= $fill))
  var pos = -1
  var idx = -1
  for r in sdata:
    inc idx
    let rnum = try: parseInt(r.attr("r")) except: -1
    if rnum == rowNum:
      return Row(sheet: s, body: r)
    elif rnum == -1 or (rnum > rowNum and r.len > 0):
      result = resultRow(rowNum)
      pos = idx
      break
  if pos == -1:
    result = resultRow(rowNum)
    sdata.add result.body
  else:
    sdata.insert result.body, pos
  s.modifiedAt

proc `hide=`*(row: Row, yes: bool) =
  ## Hide the current row
  row.body.attrs["hidden"] = $(if yes: 1 else: 0)

proc hidden*(row: Row): bool =
  ## Check whether row is hidden
  try: parseBool(row.body.attr "hidden") except: false

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

template retrieveCol(node: XmlNode, colnum: int, test, target, whenNotFound: untyped) =
  let colstr {.inject.} = $colnum
  var found = false
  for n {.inject.} in node:
    if `test`:
      `target` = n
      found = true
      break
  if not found:
    `target` = `whenNotFound`
    if `target` != nil:
      node.add `target`

proc pageBreak*(row: Row, maxCol, minCol = 0, manual = true) =
  ## Add horizontal page break after the current row working on.
  ## Set the horizontal page break length up to intended maxCol.
  let rbreak = row.sheet.body.retrieveChildOrNew "rowBreaks"
  let rnum = row.rowNum-1
  var brkn: XmlNode
  rbreak.retrieveCol(rnum, n.attr("id") == colstr, brkn, <>brk(id=colstr))
  if minCol > 0: brkn.attrs["min"] = $minCol
  if maxCol > 0: brkn.attrs["max"] = $maxCol
  brkn.attrs["man"] = $manual
  let newcount = $rbreak.len
  rbreak.attrs["count"] = newcount
  if manual:
    rbreak.attrs["manualBreakCount"] = newcount
  else:
    rbreak.attrs["manualBreakCount"] = $(rbreak.len-1)

proc lastRow*(sheet: Sheet, getEmpty = false, getHidden = false): Row =
  ## Fetch the last row available with option to fetch whether it's empty/hidden
  ## or not.
  let sdata = sheet.body.retrieveChildOrNew "sheetData"
  var idx = sdata.len - 1
  if idx < 0: return
  for i in countdown(idx, 0):
    let body = sdata[i]
    if not getEmpty and body.len == 0:
      continue
    elif not getHidden and (try: parseBool(body.attr "hidden") except: false):
      continue
    else:
      return Row(body: body, sheet: sheet)
