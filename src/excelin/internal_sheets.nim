include internal_styles

const
  xmlnsx14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

proc getSheet*(e: Excel, name: string): Sheet =
  ## Fetch the sheet from the Excel file for further work.
  ## Will return nil for unavailable sheet name.
  ## Check with `sheetNames proc<#sheetNames,Excel>`_ to find out which sheets' name available.
  var rid = ""
  for s in e.workbook.body.findAll "sheet":
    if name == s.attr "name":
      rid = s.attr "r:id"
      break
  if rid == "": return
  var targetpath = ""
  for rel in e.workbook.rels[1]:
    if rid == rel.attr "Id":
      targetpath = rel.attr "Target"
      break
  if targetpath == "": return
  let thepath = (e.workbook.path.parentDir / targetpath).unixSep
  if thepath notin e.sheets:
    return
  e.sheets[thepath]

# when adding new sheet, need to update workbook to add
# ✓ to sheets,
# ✓ its new id,
# ✓ its package relationships
# ✓ add entry to content type
# ✗ add complete worksheet nodes
# ✓ add sheet relations file pre-emptively
proc addSheet*(e: Excel, name = ""): Sheet =
  ## Add new sheet to excel with supplied name and return it to enable further working.
  ## The name can use the existing sheet name. Sheet name by default will in be
  ## `"Sheet{num}"` format with `num` is number of available sheets increased each time
  ## adding sheet. The new empty Excel file starting with Sheet1 will continue to Sheet2,
  ## Sheet3 ... each time this function is called.
  ## For example snip code below:
  ##
  ## .. code-block:: Nim
  ##
  ##   let (excel, sheet1) = newExcel()
  ##   doAssert sheet1.name == "Sheet1"
  ##   excel.deleteSheet "Sheet1"
  ##   let newsheet = addSheet excel
  ##   doAssert newsheet.name == "Sheet2" # instead of to be Sheet1
  ##
  ## This is because the counter for sheet will not be reduced despite of deleting
  ## the sheet in order to reduce maintaining relation-id cross reference.
  const
    sheetTypeNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    contentTypeNs = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  inc e.sheetCount
  var name = name
  if name == "":
    name = fmt"Sheet{e.sheetCount}"
  let wbsheets = e.workbook.body.retrieveChildOrNew "sheets"
  let rel = e.workbook.rels
  var availableId: int
  discard scanf(rel[1].findAll("Relationship")[^1].attr("Id"), "rId$i+", availableId)
  inc availableId
  let
    rid = fmt"rId{availableId}"
    sheetfname = fmt"sheet{e.sheetCount}"
    targetpath = block:
      var thepath: string
      for fpath in e.sheets.keys:
        thepath = fpath
        break
      (thepath.relativePath(e.workbook.path).parentDir.tailDir / sheetfname).unixSep.addFileExt "xml"
    sheetrelname = fmt"{sheetfname}.xml.rels"
    sheetrelpath = block:
      let (path, _) = splitPath targetpath
      (path / "_rels" / sheetrelname).unixSep

  let worksheet = newXmlTree(
    "worksheet", [<>sheetData()],
    {"xmlns:x14": xmlnsx14, "xmlns:r": xmlnsr, "xmlns:xdr": xmlnsxdr,
      "xmlns": mainns, "xmlns:mc": xmlnsmc}.toXmlAttributes)
  let sheetworkbook = newXmlTree("sheet", [],
    {"name": name, "sheetId": $availableId, "r:id": rid, "state": "visible"}.toXmlAttributes)
  let sheetrel = <>Relationships(xmlns=relPackageSheet)

  let fpath = (e.workbook.path.parentDir / targetpath).unixSep
  result = Sheet(
    body: worksheet,
    parent: e,
    privName: name,
    rid: rid,
    filename: sheetfname & ".xml",
  )
  e.sheets[fpath] = result
  wbsheets.add sheetworkbook
  e.workbook.sheetsInfo.add result.body
  rel[1].add <>Relationship(Target=targetpath, Type=sheetTypeNs, Id=rid)
  e.content.add <>Override(PartName="/" & fpath, ContentType=contentTypeNs)
  e.content.add <>Override(PartName="/" & sheetrelpath, ContentType= packagetypefmt % ["relationships"])
  e.otherfiles[sheetrelname] = (sheetrelpath, sheetrel)

# deleting sheet needs to delete several related info viz:
# ✓ deleting from the sheet table
# ✓ deleting from content
# ✓ deleting from package relationships
# ✓ deleting from workbook body
# ✗ deleting from workbook sheets info
proc deleteSheet*(e: Excel, name = "") =
  ## Delete sheet based on its name. Will ignore when it cannot find the name.
  ## Delete the first (or older) sheet when there's a same name.
  ## Check `sheetNames proc<#sheetNames,Excel>`_ to get available names.

  # delete from workbook
  var rid = ""
  let ss = e.workbook.body.child "sheets"
  if ss == nil: return
  var nodepos = -1
  for node in ss:
    inc nodepos
    if name == node.attr "name":
      rid = node.attr "r:id"
      break
  if rid == "": return
  ss.delete nodepos

  # delete from package relation and sheets table
  var targetpath = ""
  nodepos = -1
  for node in e.workbook.rels[1]:
    inc nodepos
    if rid == node.attr "Id":
      let (dir, _)  = splitPath e.workbook.path
      targetpath = (dir / node.attr "Target").unixSep.addFileExt "xml"
      break
  if targetpath == "": return
  e.workbook.rels[1].delete nodepos
  e.sheets.del targetpath

  # delete from content
  nodepos = -1
  var found = false
  for node in e.content:
    inc nodepos
    if ("/" & targetpath) == node.attr "PartName":
      found = true
      break
  if not found: return
  e.content.delete nodepos

  #nodepos = -1
  #for node in e.workbook.sheetsInfo:
    #inc nodepos

proc name*(s: Sheet): string = s.privName
  ## Get the name of current sheet

proc `name=`*(s: Sheet, newname: string) =
  ## Update sheet's name.
  s.privName = newname
  for node in s.parent.workbook.body.findAll "sheet":
    if s.rid == node.attr "r:id":
      var currattr = node.attrs
      currattr["name"] = newname
      node.attrs = currattr

proc resetMerge*(sheet: Sheet, `range`: Range) =
  ## Remove any merge cell with defined range.
  ## Ignored if there's no such such range supplied.
  let
    refrange = $`range`
    mcells = sheet.body.child "mergeCells"
  if mcells == nil: return
  var
    pos = -1
    found = false
  for mcell in mcells:
    inc pos
    if refrange == mcell.attr "ref":
      found = true
      break
  if not found: return
  mcells.delete pos
  styleRange(sheet, `range`, copyStyle)

proc `ranges=`*(sheet: Sheet, `range`: Range) =
  ## Set the ranges of data/table within sheet.

  var dim = $`range`
  if dim == "": dim = "A1"
  let dimn = sheet.body.retrieveChildOrNew "dimension"
  dimn.attrs["ref"] = dim

proc `autoFilter=`*(sheet: Sheet, `range`: Range) =
  ## Add auto filter to selected range. Setting this range
  ## will override the previous range setting to sheet.
  ## Providing with range ("", "") will delete the auto filter
  ## in the sheet.
  if `range`[0] == "" and `range`[1] == "":
    var
      autoFilterPos = -1
      autoFilterFound = false
    for n in sheet.body:
      inc autoFilterPos
      if n.tag == "autoFilter":
        autoFilterFound = true
        break
    if autoFilterFound:
      sheet.body.delete autoFilterPos
    return
  sheet.ranges = `range`
  let dim = $`range`
  (sheet.body.retrieveChildOrNew "autoFilter").attrs["ref"] = dim

  (sheet.body.retrieveChildOrNew "sheetPr").attrs["filterMode"] = $true

proc autoFilter*(sheet: Sheet): Range =
  ## Retrieve the set range for auto filter. Mainly used to check
  ## whether the range for set is already set to add filtering to
  ## its column number range (0-based).
  let autoFilter = sheet.body.child "autoFilter"
  if autoFilter == nil: return
  discard scanf(autoFilter.attr "ref", "$w:$w", result[0], result[1])

proc filterCol*(sheet: Sheet, colId: Natural, filter: Filter) =
  ## Set filter to the sheet range. Ignored if sheet hasn't
  ## set its auto filter range. Set the col with default Filter()
  ## to reset it.
  let autoFilter = sheet.body.child "autoFilter"
  if autoFilter == nil: return
  let fcolumns = <>filterColumn(colId= $colId)
  case filter.kind
  of ftFilter:
    let filters = <>filters()
    for val in filter.valuesStr:
      filters.add <>filter(val=val)
    fcolumns.add filters
  of ftCustom:
    let cusf = newXmlTree("customFilters", [],
      { $filter.logic: $true }.toXmlAttributes)
    for (op, val) in filter.customs:
      cusf.add <>costumFilter(operator= $op, val=val)
    fcolumns.add cusf

  autoFilter.add fcolumns

proc `mergeCells=`*(sheet: Sheet, `range`: Range) =
  ## Merge cells will remove any existing values within
  ## range cells to be merged. Will only retain the topleft
  ## cell value when merging the range.
  let
    (topleftcol, topleftrow) = `range`[0].colrow
    (botrightcol, botrightrow) = `range`[1].colrow
    mcells = sheet.body.retrieveChildOrNew "mergeCells"
    horizontalRange = toSeq[topleftcol.toNum .. botrightcol.toNum]

  mcells.add <>mergeCell(ref= $`range`)

  let
    r = sheet.row topleftrow
    topleftcell = r.fetchValNode(topleftcol, $cfSparse == r.body.attr "cellfill")
  var styleattr = if topleftcell == nil: "0" else: topleftcell.attr "s"
  if styleattr == "": styleattr = "0"

  template addEmptyCell(r: Row, col, s: string): untyped =
    r.addCell col, cellType = "", text = "", valelem = "",
      emptyCell = true, style = s

  for cn in horizontalRange[1..^1]:
    r.addEmptyCell cn.toCol, styleattr
  for rnum in topleftrow+1 .. botrightrow:
    for cn in horizontalRange:
      let r = sheet.row rnum
      r.addEmptyCell cn.toCol, styleattr

proc retrieveCol(node: XmlNode, colnum: int): XmlNode =
  let colstr = $colnum
  for n in node:
    if n.attr("min") == colstr and n.attr("max") == colstr:
      return n
  result = <>col(min=colstr, max=colstr)
  node.add result

template modifyCol(sheet: Sheet, col, attr, val: string) =
  let coln = sheet.body.retrieveChildOrNew("cols").retrieveCol col.toNum+1
  coln.attrs[attr] = val

proc colHide*(sheet: Sheet, col: string, hide: bool) =
  ## Hide entire column in the sheet.
  sheet.modifyCol(col, "hidden", $hide)

proc colOutlineLevel*(sheet: Sheet, col: string, level: Natural) =
  ## Set outline level for the entire column in the sheet.
  sheet.modifyCol(col, "outlineLevel", $level)

proc colWidth*(sheet: Sheet, col: string, width: float) =
  ## Set the entire column width. Set with 0 width to reset it.
  let cols = sheet.body.retrieveChildOrNew "cols"
  let coln = cols.retrieveCol col.toNum+1
  if width <= 0:
    coln.attrs.del "width"
    coln.attrs["customWidth"] = $false
    return
  coln.attrs["customWidth"] = $true
  coln.attrs["width"] = $width