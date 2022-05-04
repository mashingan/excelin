# Excelin
# Library to read and create Excel file purely in Nim
# MIT License Copyright (c) 2022 Rahmatullah

## Excelin
## *******
##
## A library to work with spreadsheet file (strictly .xlsx) without dependency
## outside of Nim compiler (and its requirement) and its development environment.
##

from std/xmltree import XmlNode, findAll, `$`, child, items, attr, `<>`,
     newXmlTree, add, newText, toXmlAttributes, delete, len, xmlHeader,
     attrs, `attrs=`, innerText, `[]`, insert, clear
from std/xmlparser import parseXml
from std/strutils import endsWith, contains, parseInt, `%`, replace,
  parseFloat, parseUint, toUpperAscii, join
from std/sequtils import toSeq, mapIt, repeat
from std/tables import TableRef, newTable, `[]`, `[]=`, contains, pairs,
     keys, del, values, initTable, len
from std/strformat import fmt
from std/times import DateTime, Time, now, format, toTime, toUnixFloat,
  parse, fromUnix, local
from std/os import `/`, addFileExt, parentDir, splitPath,
  getTempDir, removeFile, extractFilename, relativePath, tailDir
from std/strtabs import `[]=`
from std/sugar import dump, `->`
from std/strscans import scanf
from std/sha1 import secureHash, `$`
from std/math import `^`

from zippy/ziparchives import openZipArchive, extractFile, ZipArchive,
  ArchiveEntry, writeZipArchive

export xmltree.items
#export xmltree.`$`

const
  datefmt = "yyyy-MM-dd'T'HH:mm:ss'.'fffzz"
  xmlnsx14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlnsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlnsxdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlnsmc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
  spreadtypefmt = "application/vnd.openxmlformats-officedocument.spreadsheetml.$1+xml"
  mainns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  relSharedStrScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  emptyxlsx = currentSourcePath.parentDir() / "empty.xlsx"
  excelinVersion* = "0.3.6"

type
  Excel* = ref object
    ## The object the represent as Excel file, mostly used for reading the Excel
    ## and writing it to Zip file and to memory buffer (at later version).
    content: XmlNode
    rels: XmlNode
    workbook: Workbook
    sheets: TableRef[FilePath, Sheet]
    sharedStrings: FileRep
    otherfiles: TableRef[string, FileRep]
    embedfiles: TableRef[string, EmbedFile]
    sheetCount: int

  InternalBody = object of RootObj
    body: XmlNode

  Workbook* = ref object of InternalBody
    ## The object that used for managing package information of Excel file.
    ## Most users won't need to use this.
    path: string
    sheetsInfo: seq[XmlNode]
    rels: FileRep
    parent: Excel

  Sheet* = ref object of InternalBody
    ## The main object that will be used most of the time for many users.
    ## This object will represent a sheet in Excel file such as adding row
    ## getting row, and/or adding cell directly.
    sharedStrings: XmlNode
    parent: Excel
    rid: string
    privName: string

  Row* = ref object of InternalBody
    ## The object that will be used for working with values within cells of a row.
    ## Users can get the value within cell and set its value.
    sheet: Sheet
  FilePath = string
  FileRep = (FilePath, XmlNode)
  EmbedFile = (FilePath, string)

  ExcelError* = object of CatchableError
    ## Error when the Excel file read is invalid, specifically Excel file
    ## that doesn't have workbook.

  CellFill* = enum
    cfSparse = "sparse"
    cfFilled = "filled"

  Formula* = object
    ## Object exclusively working with formula in a cell.
    ## The equation is simply formula representative and
    ## the valueStr is the value in its string format,
    ## which already calculated beforehand.
    equation*: string
    valueStr*: string

template unixSep(str: string): untyped = str.replace('\\', '/')
  ## helper to change the Windows path separator to Unix path separator

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


proc getSheetData(s: XmlNode): XmlNode =
  result = s.child "sheetData"
  if result == nil:
    result = <>sheetData()

proc modifiedAt*[T: DateTime | Time](e: Excel, t: T = now()) =
  ## Update Excel modification time.
  const key = "core.xml"
  if key notin e.otherfiles: return
  let (_, prop) = e.otherfiles[key]
  let modifn = prop.child "dcterms:modified"
  if modifn == nil: return
  modifn.delete 0
  modifn.add newText(t.format datefmt)

proc modifiedAt*[Node: Workbook|Sheet, T: DateTime | Time](w: Node, t: T = now()) =
  ## Update workbook or worksheet modification time.
  w.parent.modifiedAt t

proc addRow*(s: Sheet): Row {.deprecated: "use `row(Sheet, Positive): Row` instead".} =
  ## Add row directly from sheet after any existing row.
  ## This will return a new pristine row to work further.
  ## Row numbering is 1-based.
  let sdata = s.body.getSheetData
  let rowExists = sdata.len
  result = Row(
    sheet: s,
    body: <>row(r= $(rowExists+1), hidden="false", collapsed="false"),
  )
  sdata.add result.body
  s.modifiedAt

proc row*(s: Sheet, rowNum: Positive, fill = cfSparse): Row =
  ## Add row by selecting which row number to work with.
  ## This will return new row if there's no existing row
  ## or will return an existing one.
  let sdata = s.body.getSheetData
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

proc addRow*(s: Sheet, rowNum: Positive): Row
  {.deprecated: "use `row(Sheet, Positive): Row` instead".}
  = s.row rowNum

proc rowNum*(r: Row): Positive =
  ## Getting the current row number of Row object users working it.
  result = try: parseInt(r.body.attr "r") except: 1

proc fetchCell(body: XmlNode, colrow: string): int =
  var count = -1
  for n in body:
    inc count
    if colrow == n.attr "r": return count
  -1

proc toNum*(col: string): int =
  ## Convert our column string to its numeric representation.
  ## Make sure the supplied column is already in upper case.
  ## 0-based e.g.: "A".toNum == 0, "C".toNum == 2
  ## Complement of `toCol <#toCol,Natural,string>`_.
  runnableExamples:
    let colnum = [("A", 0), ("AA", 26), ("AB", 27), ("ZZ", 701), ("AAA", 702),
      ("AAB", 703)]
    for cn in colnum:
      doAssert cn[0].toNum == cn[1]
  for i in countdown(col.len-1, 0):
    let cnum = col[col.len-i-1].ord - 'A'.ord + 1
    result += cnum * (26 ^ i)
  dec result

let atoz = toSeq('A'..'Z')

proc toCol*(n: Natural): string =
  ## Convert our numeric column to string representation.
  ## The numeric should be 0-based, e.g.: 0.toCol == "A", 25.toCol == "Z"
  ## Complement of `toNum <#toNum,string,int>`_.
  runnableExamples:
    let colnum = [("A", 0), ("AA", 26), ("AB", 27), ("ZZ", 701), ("AAA", 702),
      ("AAB", 703)]
    for cn in colnum:
      doAssert cn[1].toCol == cn[0]
  if n < atoz.len:
    return $atoz[n]
  var count = n
  while count >= atoz.len:
    let c = count div atoz.len
    if c >= atoz.len:
      result &= (c-1).toCol
    else:
      result &= atoz[c-1]
    count = count mod atoz.len
  result &= atoz[count]

proc addCell(row: Row, col, cellType, text: string, valelem = "v", altnode: seq[XmlNode] = @[]) =
  let rn = row.body.attr "r"
  let sparse = $cfSparse == row.body.attr "cellfill"
  let col = col.toUpperAscii
  let cellpos = fmt"{col}{rn}"
  let innerval = if altnode.len > 0: altnode else: @[newText text]
  let cnode = if cellType != "" and valelem != "":
                <>c(r=cellpos, s="0", t=cellType, newXmlTree(valelem, innerval))
              elif valelem != "": <>c(r=cellpos, s="0", newXmlTree(valelem, innerval))
              else: newXmlTree("c", innerval, {"r": cellpos}.toXmlAttributes)
  if not sparse:
    let cellsTotal = row.body.len
    let colnum = toNum col
    if colnum < cellsTotal:
      row.body.delete colnum
      row.body.insert cnode, colnum
    else:
      for i in cellsTotal ..< colnum:
        let colchar = toCol i
        let cellp = fmt"{colchar}{rn}"
        row.body.add <>c(r=cellp)
      row.body.add cnode
    return
  let nodepos = row.body.fetchCell(cellpos)
  if nodepos < 0:
    row.body.add cnode
  else:
    row.body.delete nodepos
    row.body.insert cnode, nodepos

proc addSharedString(r: Row, col, s: string) =
  var
    pos = 0
    found = false
  let sharedStr = r.sheet.sharedStrings
  for ss in sharedStr:
    inc pos
    if s == ss.innerText:
      found = true
      break

  var
    count = try: parseInt(sharedStr.attr "count") except: 0
    uniq = try: parseInt(sharedStr.attr "uniqueCount") except: 0

  inc count
  if not found:
    inc uniq
    sharedStr.add <>si(newXmlTree("t",
      [newText s], {"xml:space": "preserve"}.toXmlAttributes))
  else:
    dec pos # shared string 0-based

  r.addCell col, "s", $pos
  sharedStr.attrs = {"count": $count,
    "uniqueCount": $uniq, "xmlns": mainns}.toXmlAttributes
  r.sheet.modifiedAt

proc `[]=`*(row: Row, col: string, s: string) =
  ## Add cell with overload for value string. Supplied column
  ## is following the Excel convention starting from A -  Z, AA - AZ ...
  if s.len < 64:
    row.addCell col, "inlineStr", s, "is", @[<>t(newText s)]
    row.sheet.modifiedAt
    return
  row.addSharedString(col, s)

proc `[]=`*(row: Row, col: string, n: SomeNumber) =
  ## Add cell with overload for any number.
  row.addCell col, "n", $n
  row.sheet.modifiedAt

proc `[]=`*(row: Row, col: string, b: bool) =
  ## Add cell with overload for truthy value.
  row.addCell col, "b", $b
  row.sheet.modifiedAt

proc `[]=`*(row: Row, col: string, d: DateTime | Time) =
  ## Add cell with overload for DateTime or Time. The saved value
  ## will be in string format of `yyyy-MM-dd'T'HH:mm:ss'.'fffzz` e.g.
  ## `2200-10-01T11:22:33.456-03`.
  row.addCell col, "d", d.format(datefmt)
  row.sheet.modifiedAt

proc `[]=`*(row: Row, col: string, f: Formula) =
  row.addCell col, "", "", "",
    @[<>f(newText f.equation), <>v(newText f.valueStr)]
  row.sheet.modifiedAt

proc getCell*[R](row: Row, col: string, conv: string -> R = nil): R =
  ## Get cell value from row with optional function to convert it.
  ## When conversion function is supplied, it will be used instead of
  ## default conversion function viz:
  ## * SomeSignedInt (int/int8/int32/int64): strutils.parseInt
  ## * SomeUnsignedInt (uint/uint8/uint32/uint64): strutils.parseUint
  ## * SomeFloat (float/float32/float64): strutils.parseFloat
  ## * DateTime | Time: times.format with layout `yyyy-MM-dd'T'HH:mm:ss'.'fffzz`
  ##
  ## For simple usage example:
  ##
  ## .. code-block:: Nim
  ##
  ##   let strval = row.getCell[:string]("A") # we're fetching string value in colum A
  ##   let intval = row.getCell[:int]("B")
  ##   let uintval = row.getCell[:uint]("C")
  ##   let dtval = row.getCell[:DateTime]("D")
  ##   let timeval = row.getCell[:Time]("E")
  ##
  ##   # below example we'll get the DateTime that has been formatted like 2200/12/01
  ##   # so we supply the optional custom converter function
  ##   let dtCustom = row.getCell[:DateTime]("F", (s: string) -> DateTime => (
  ##      s.parse("yyyy/MM/dd")))
  ##
  ## Any other type that other than mentioned above should provide the closure proc
  ## for the conversion otherwise it will return the default value, for example
  ## any ref object will return nil or for object will get the object with its field
  ## filled with default values.
  when R is SomeSignedInt: result = R.low
  elif R is SomeUnsignedInt: result = R.high
  elif R is SomeFloat: result = NaN
  elif R is DateTime: result = fromUnix(0).local
  elif R is Time: result = fromUnix 0
  else: discard
  let rnum = row.body.attr "r"
  let isSparse = $cfSparse == row.body.attr "cellfill"
  let col = col.toUpperAscii
  var isInnerStr = false
  let v = block:
    var x: XmlNode
    let colnum = col.toNum
    if not isSparse and colnum < row.body.len:
      x = row.body[colnum]
    else:
      for node in row.body:
        if fmt"{col}{rnum}" == node.attr "r":
          isInnerStr = "inlineStr" == node.attr "t"
          x = node
          break
    x
  if v == nil: return
  let t = v.innerText
  template fetchShared(t: string): untyped =
    let refpos = try: parseInt(t) except: -1
    if refpos < 0: return
    let tnode = row.sheet.sharedStrings[refpos]
    if tnode == nil: return
    tnode.innerText
  template retconv =
    if conv != nil:
      var tt = t
      if "s" == v.attr "t":
        tt = fetchShared t
      return conv tt
  retconv()
  when R is string:
    result = if isInnerStr: t else: fetchShared t
  elif R is SomeSignedInt:
    try: result = parseInt(t) except: discard
  elif R is SomeFloat:
    try: result = parseFloat(t) except: discard
  elif R is SomeUnsignedInt:
    try: result = parseUint(t) except: discard
  elif R is DateTime:
    try: result = parse(t, datefmt) except: discard
  elif R is Time:
    try: result = parse(t, datefmt).toTime except: discard
  elif R is Formula:
    result = Formula(equation: v.child("f").innerText,
      valueStr: v.child("v").innerText)
  else:
    discard

template getCellIt*[R](r: Row, col: string, body: untyped): untyped =
  ## Shorthand for `getCell <#getCell,Row,string,typeof(nil)>`_ with
  ## injected `it` in body.
  ## For example:
  ##
  ## .. code-block:: Nim
  ##
  ##   from std/times import parse, year, month, DateTime, Month, monthday
  ##
  ##   # the value in cell is "2200/12/01"
  ##   let dt = row.getCell[:DateTime]("F", (s: string) -> DateTime => (
  ##      s.parse("yyyy/MM/dd")))
  ##   doAssert dt.year == 2200
  ##   doAssert dt.month == mDec
  ##   doAssert dt.monthday = 1
  ##
  r.getCell[:R](col, proc(it {.inject.}: string): R = `body`)

proc `[]`*(r: Row, col: string, ret: typedesc): ret =
  ## Getting cell value from supplied return typedesc. This is overload
  ## of basic supported values that will return default value e.g.:
  ## * string default to ""
  ## * SomeSignedInt default to int.low
  ## * SomeUnsignedInt default to uint.high
  ## * SomeFloat default to NaN
  ## * DateTime and Time default to empty object time
  ##
  ## Other than above mentioned types, see `getCell proc<#getCell,Row,string,typeof(nil)>`_
  ## for supplying the converting closure for getting the value.
  getCell[ret](r, col)

proc clear*(row: Row) = row.body.clear
  ## Clear all cells in the row.

# when adding new sheet, need to update workbook to add
# ✓ to sheets,
# ✓ its new id,
# ✓ its package relationships
# ✓ add entry to content type
# ✗ add complete worksheet nodes
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
  var wbsheets = e.workbook.body.child "sheets"
  if wbsheets == nil:
    wbsheets = <>sheets()
    e.workbook.body.add wbsheets
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

  let worksheet = newXmlTree(
    "worksheet", [<>sheetData()],
    {"xmlns:x14": xmlnsx14, "xmlns:r": xmlnsr, "xmlns:xdr": xmlnsxdr,
      "xmlns": mainns, "xmlns:mc": xmlnsmc}.toXmlAttributes)
  let sheetworkbook = newXmlTree("sheet", [],
    {"name": name, "sheetId": $availableId, "r:id": rid, "state": "visible"}.toXmlAttributes)

  let fpath = (e.workbook.path.parentDir / targetpath).unixSep
  result = Sheet(
    body: worksheet,
    sharedStrings: e.sharedStrings[1],
    parent: e,
    privName: name,
    rid: rid)
  e.sheets[fpath] = result
  wbsheets.add sheetworkbook
  e.workbook.sheetsInfo.add result.body
  rel[1].add <>Relationship(Target=targetpath, Type=sheetTypeNs, Id=rid)
  e.content.add <>Override(PartName="/" & fpath, ContentType=contentTypeNs)

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


proc retrieveSheetsInfo(n: XmlNode): seq[XmlNode] =
  let sheets = n.child "sheets"
  if sheets == nil: return
  result = toSeq sheets.items

# to add shared strings we need to
# ✓ define the xml
# ✓ add to content
# ✓ add to package rels
# ✓ update its path
proc addSharedStrings(e: Excel) =
  e.sharedStrings[0] = (e.workbook.path.parentDir / "sharedStrings.xml").unixSep
  e.sharedStrings[1] = <>sst(xmlns=mainns, count="0", uniqueCount="0")
  e.content.add <>Override(PartName="/" & e.sharedStrings[0],
    ContentType=spreadtypefmt % ["sharedStrings"])
  let relslen = e.workbook.rels[1].len
  e.workbook.rels[1].add <>Relationship(Target="sharedStrings.xml",
    Id=fmt"rId{relslen+1}", Type=relSharedStrScheme)

proc assignSheetInfo(e: Excel) =
  var mapRidName = initTable[string, string]()
  var mapFilenameRid = initTable[string, string]()
  for s in e.workbook.body.findAll "sheet":
    mapRidName[s.attr "r:id"] = s.attr "name"
  for s in e.workbook.rels[1]:
    mapFilenameRid[s.attr("Target").extractFilename] = s.attr "Id"
  for path, sheet in e.sheets:
    sheet.rid = mapFilenameRid[path.extractFilename]
    sheet.privName = mapRidName[sheet.rid]

proc readExcel*(path: string): Excel =
  ## Read Excel file from supplied path. Will raise OSError
  ## in case path is not exists, IOError when system errors
  ## during reading the file, ExcelError when the Excel file
  ## is not valid (Excel file that has no workbook).
  let reader = openZipArchive path
  result = Excel()
  template extract(path: string): untyped =
    parseXml reader.extractFile(path)
  template fileRep(path: string): untyped =
    (path, extract path)
  result.content = extract "[Content_Types].xml"
  result.rels = extract "_rels/.rels"
  var found = false
  result.sheets = newTable[string, Sheet]()
  result.otherfiles = newTable[string, FileRep]()
  result.embedfiles = newTable[string, EmbedFile]()
  for node in result.content.findAll("Override"):
    let wbpath = node.attr "PartName"
    if wbpath == "": continue
    let contentType = node.attr "ContentType"
    let path = wbpath.tailDir # because the wbpath has leading '/' due to top package position
    if wbpath.endsWith "workbook.xml":
      let body = extract path
      result.workbook = Workbook(
        path: path,
        body: body,
        sheetsInfo: body.retrieveSheetsInfo,
        parent: result,
      )
      found = true
    elif "worksheet" in contentType:
      inc result.sheetCount
      let sheet = extract path
      result.sheets[path] = Sheet(body: sheet, parent: result)
    elif wbpath.endsWith "workbook.xml.rels":
      result.workbook.rels = fileRep path
    elif wbpath.endsWith "sharedStrings.xml":
      result.sharedStrings = fileRep path
    elif wbpath.endsWith(".xml") or wbpath.endsWith(".rels"): # any others xml/rels files
      let (_, f) = splitPath wbpath
      result.otherfiles[f] = path.fileRep
    else:
      let (_, f) = splitPath wbpath
      result.embedfiles[f] = (path, reader.extractFile path)
  if not found:
    raise newException(ExcelError, "No workbook found, invalid excel file")
  if result.sharedStrings[1] == nil:
    result.addSharedStrings
  result.assignSheetInfo
  for _, s in result.sheets:
    s.sharedStrings = result.sharedStrings[1]

  if "app.xml" in result.otherfiles:
    var (_, appnode) = result.otherfiles["app.xml"]
    clear appnode
    appnode.add <>Application(newText "Excelin")
    appnode.add <>AppVersion(newText excelinVersion)

proc `prop=`*(e: Excel, prop: varargs[(string, string)]) =
  ## Add information property to Excel file. Will add the properties
  ## to the existing.
  const key = "app.xml"
  if key notin e.otherfiles: return
  let (_, propnode) = e.otherfiles[key]
  for p in prop:
    propnode.add newXmlTree(p[0], [newText p[1]])

proc newExcel*(appName = "Excelin"): (Excel, Sheet) =
  ## Return a new Excel and Sheet at the same time to work for both.
  ## The Sheet returned is by default has name "Sheet1" but user can
  ## use `name= proc<#name=,Sheet,string>`_ to change its name.
  let excel = readExcel emptyxlsx
  (excel, excel.getSheet "Sheet1")

proc writeFile*(e: Excel, targetpath: string) =
  ## Write Excel to file in target path. Raise OSError when it can't
  ## write to the intended path.
  let archive = ZipArchive()
  let lastmod = now().toTime
  template addContents(path, content: string) =
    archive.contents[path] = ArchiveEntry(
      contents: xmlHeader & content,
      lastModified: lastmod,
    )
  "[Content_Types].xml".addContents $e.content
  "_rels/.rels".addContents $e.rels
  e.workbook.path.addContents $e.workbook.body
  e.workbook.rels[0].addContents $e.workbook.rels[1]
  e.sharedStrings[0].addContents $e.sharedStrings[1]
  for p, s in e.sheets:
    p.addContents $s.body
  for rep in e.otherfiles.values:
    rep[0].addContents $rep[1]
  for embed in e.embedfiles.values:
    archive.contents[embed[0]] = ArchiveEntry(
      contents: embed[1],
      lastModified: lastmod,
    )

  archive.writeZipArchive targetpath

proc `$`*(e: Excel): string =
  ## Get Excel file as string. Currently implemented by writing to
  ## temporary dir first because there's no API to get the data
  ## directly.
  let fname = fmt"excelin-{now().toTime.toUnixFloat}"
  let path = (getTempDir() / $fname.secureHash).addFileExt ".xlsx"
  e.writeFile path
  result = readFile path
  try:
    removeFile path
  except:
    discard

proc sheetNames*(e: Excel): seq[string] =
  ## Return all availables sheet names within an Excel file.
  e.workbook.body.findAll("sheet").mapIt( it.attr "name" )

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


when isMainModule:
  let (empty, sheet) = newExcel()

  let colnum = [("A", 0), ("AA", 26), ("AB", 27), ("ZZ", 701)]
  for cn in colnum:
    doAssert cn[0].toNum == cn[1]
    doAssert cn[1].toCol == cn[0]

  empty.writeFile "generate-base-empty.xlsx"
  if sheet != nil:
    let row = sheet.row(1, cfFilled)
    row["A"] = "hehe"
    row["B"] = -1
    row["C"] = 2
    row["D"] = 42.0
    sheet.name = "hehe"
    let newsheet = empty.addSheet("test add new sheet")
    dump newsheet.name
    dump row.getCell[:string]("A")
    dump row["B", int]
    dump row.getCell[:uint]("C")
    dump row["D", float]
    empty.writeFile "generate-modified.xlsx"
    row["D"] = now() # modify to date time
    empty.writeFile "generate-modified-cell.xlsx"
    #empty.deleteSheet "hehe"
    #empty.writeFile "generate-deleted-sheet.xlsx"
    let row5 = sheet.row 5
    row5["A"] = "yeehaa"
    let row6 = sheet.row 6
    row6["B"] = 5
    row6["A"] = -1
    let row10 = sheet.row(10, cfFilled)
    row10["C"] = 11
    row10["D"] = -11
    empty.prop = {"key1": "val1", "prop-custom": "custom-setting"}
    dump row5["A", string]
    row5["A"] = row5["A", string] & " heehaa"
    let tobeShared = "brown fox jumps over the lazy dog".repeat(5).join(";")
    row5["B"] = tobeShared
    row5["c"] = tobeShared
    let row11 = sheet.row 11
    var sum = 0
    for i in 0 .. 9:
      row11[i.toCol] = i
      sum += i
    row11[10.toCol] = Formula(equation: "SUM(A11:J11)", valueStr: $sum)
    dump row11[10.toCol, Formula]
    dump sheet.body
    #dump empty.sharedStrings
    #dump empty.otherfiles["app.xml"]
    empty.writeFile "generated-add-rows.xlsx"
  else:
    echo "sheet is nil"

  #echo $empty.workbook.body
  dump empty.sheetNames
