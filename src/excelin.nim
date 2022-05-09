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
     attrs, `attrs=`, innerText, `[]`, insert, clear, XmlNodeKind, kind,
     tag
from std/xmlparser import parseXml
from std/strutils import endsWith, contains, parseInt, `%`, replace,
  parseFloat, parseUint, toUpperAscii, join, startsWith, Letters, Digits
from std/sequtils import toSeq, mapIt, repeat
from std/tables import TableRef, newTable, `[]`, `[]=`, contains, pairs,
     keys, del, values, initTable, len
from std/strformat import fmt
from std/times import DateTime, Time, now, format, toTime, toUnixFloat,
  parse, fromUnix, local
from std/os import `/`, addFileExt, parentDir, splitPath,
  getTempDir, removeFile, extractFilename, relativePath, tailDir
from std/strtabs import `[]=`, pairs, newStringTable, del
from std/sugar import dump, `->`
from std/strscans import scanf
from std/sha1 import secureHash, `$`
from std/math import `^`
from std/colors import `$`, colWhite, colRed, colGreen, colBlue


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
  relStylesScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  relPackageSheet = "http://schemas.openxmlformats.org/package/2006/relationships"
  relHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  packagetypefmt = "application/vnd.openxmlformats-package.$1+xml"
  emptyxlsx = currentSourcePath.parentDir() / "empty.xlsx"
  excelinVersion* = "0.4.2"

type
  Excel* = ref object
    ## The object the represent as Excel file, mostly used for reading the Excel
    ## and writing it to Zip file and to memory buffer (at later version).
    content: XmlNode
    rels: XmlNode
    workbook: Workbook
    sheets: TableRef[FilePath, Sheet]
    sharedStrings: SharedStrings
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
    parent: Excel
    rid: string
    privName: string
    filename: string

  Row* = ref object of InternalBody
    ## The object that will be used for working with values within cells of a row.
    ## Users can get the value within cell and set its value.
    sheet: Sheet
  FilePath = string
  FileRep = (FilePath, XmlNode)
  EmbedFile = (FilePath, string)

  SharedStrings = ref object of InternalBody
    path: string
    strtables: TableRef[string, int]
    count: Natural
    unique: Natural

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

  Font* = object
    ## Cell font styling. Provide name if it's intended to style the cell.
    ## If no name is supplied, it will ignored. Field `family` and `charset`
    ## are optionals but in order to be optional, provide it with negative value
    ## because there's value for family and charset 0. Since by default int is 0,
    ## it could be yield different style if the family and charset are not intended
    ## to be filled but not assigned with negative value.
    name*: string
    family*: int
    charset*: int
    size*: Positive
    bold*: bool
    italic*: bool
    strike*: bool
    outline*: bool
    shadow*: bool
    condense*: bool
    extend*: bool
    color*: string
    underline*: Underline
    verticalAlign*: VerticalAlign

  Underline* = enum
    uNone = "none"
    uSingle = "single"
    uDouble = "double"
    uSingleAccounting = "singleAccounting"
    uDoubleAccounting = "doubleAccounting"

  VerticalAlign* = enum
    vaBaseline = "baseline"
    vaSuperscript = "superscript"
    vaSubscript = "subscript"

  Border* = object
    ## The object that will define the border we want to apply to cell.
    ## Use `border <#border,BorderProp,BorderProp,BorderProp,BorderProp,BorderProp,BorderProp,bool,bool>`_
    ## to initialize working border instead because the indication whether border can be edited is private.
    edit: bool
    start*: BorderProp # left
    `end`*: BorderProp # right
    top*: BorderProp
    bottom*: BorderProp
    vertical*: BorderProp
    horizontal*: BorderProp
    diagonalUp*: bool
    diagonalDown*: bool

  BorderProp* = object
    ## The object that will define the style and color we want to apply to border
    ## Use `borderProp<#borderProp,BorderStyle,string>`_
    ## to initialize working border prop instead because the indication whether
    ## border properties filled is private.
    edit: bool ## indicate whether border properties is filled
    style*: BorderStyle
    color*: string #in RGB

  BorderStyle* = enum
    bsNone = "none"
    bsThin = "thin"
    bsMedium = "medium"
    bsDashed = "dashed"
    bsDotted = "dotted"
    bsThick = "thick"
    bsDouble = "double"
    bsHair = "hair"
    bsMediumDashed = "mediumDashed"
    bsDashDot = "dashDot"
    bsMediumDashDot = "mediumDashDot"
    bsDashDotDot = "dashDotDot"
    bsMediumDashDotDot = "mediumDashDotDot"
    bsSlantDashDot = "slantDashDot"

  Fill* = object
    ## Fill cell style. Use `fillStyle <#fillStyle,PatternFill,GradientFill>`_
    ## to initialize this object to indicate cell will be edited with this Fill.
    edit: bool
    pattern*: PatternFill
    gradient*: GradientFill

  PatternFill* = object
    ## Pattern to fill the cell. Use `patternFill<#patternFill,string,string,PatternType>`_
    ## to initialize.
    edit: bool
    fgColor*: string
    bgColor: string
    patternType*: PatternType

  PatternType* = enum
    ptNone = "none"
    ptSolid = "solid"
    ptMediumGray = "mediumGray"
    ptDarkGray = "darkGray"
    ptLightGray = "lightGray"
    ptDarkHorizontal = "darkHorizontal"
    ptDarkVertical = "darkVertical"
    ptDarkDown = "darkDown"
    ptDarkUp = "darkUp"
    ptDarkGrid = "darkGrid"
    ptDarkTrellis = "darkTrellis"
    ptLightHorizontal = "lightHorizontal"
    ptLightVertical = "lightVertical"
    ptLightDown = "lightDown"
    ptLightUp = "lightUp"
    ptLightGrid = "lightGrid"
    ptLightTrellis = "lightTrellis"
    ptGray125 = "gray125"
    ptGray0625 = "gray0625"

  GradientFill* = object
    ## Gradient to fill the cell. Use
    ## `gradientFill<#gradientFill,GradientStop,GradientType,float,float,float,float,float>`_
    ## to initialize.
    edit: bool
    stop*: GradientStop
    `type`*: GradientType
    degree*: float
    left*: float
    right*: float
    top*: float
    bottom*: float

  GradientStop* = object
    ## Indicate where the gradient will stop with its color at stopping position.
    color*: string
    position*: float

  GradientType* = enum
    gtLinear = "linear"
    gtPath = "path"

  Range* = (string, string)
    ## Range of table which consist of top left cell and bottom right cell.

  FilterType* = enum
    ftFilter
    ftCustom

  Filter* = object
    ## Filtering that supplied to column id in sheet range. Ignored if the sheet
    ## hasn't set its auto filter range.
    case kind*: FilterType
    of ftFilter:
      valuesStr*: seq[string]
    of ftCustom:
      logic*: CustomFilterLogic
      customs*: seq[(FilterOperator, string)]

  FilterOperator* = enum
    foEq = "equal"
    foLt = "lessThan"
    foLte = "lessThanOrEqual"
    foNeq = "notEqual"
    foGte = "greaterThanOrEqual"
    foGt = "greaterThan"

  CustomFilterLogic* = enum
    cflAnd = "and"
    cflOr = "or"
    cflXor = "xor"

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

template retrieveChildOrNew(node: XmlNode, name: string): XmlNode =
  var r = node.child name
  if r == nil:
    r = newXmlTree(name, [], newStringTable())
    node.add r
  r

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

proc rowNum*(r: Row): Positive =
  ## Getting the current row number of Row object users working it.
  result = try: parseInt(r.body.attr "r") except: 1

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
      cnode.attrs["s"] = row.body[colnum].attr "s"
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
    cnode.attrs["s"] = row.body[nodepos].attr "s"
    row.body.delete nodepos
    row.body.insert cnode, nodepos

proc addSharedString(r: Row, col, s: string) =
  let sstr = r.sheet.parent.sharedStrings
  var pos = sstr.strtables.len
  if  s notin sstr.strtables:
    inc sstr.unique
    sstr.body.add <>si(newXmlTree("t", [newText s], {"xml:space": "preserve"}.toXmlAttributes))
    sstr.strtables[s] = pos
  else:
    pos = sstr.strtables[s]

  inc sstr.count
  r.addCell col, "s", $pos
  sstr.body.attrs = {"count": $sstr.count, "uniqueCount": $sstr.unique, "xmlns": mainns}.toXmlAttributes
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
    let tnode = row.sheet.parent.sharedStrings.body[refpos]
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
  let sstr = SharedStrings(strtables: newTable[string, int]())
  sstr.path = (e.workbook.path.parentDir / "sharedStrings.xml").unixSep
  sstr.body = <>sst(xmlns=mainns, count="0", uniqueCount="0")
  e.content.add <>Override(PartName="/" & sstr.path,
    ContentType=spreadtypefmt % ["sharedStrings"])
  let relslen = e.workbook.rels[1].len
  e.workbook.rels[1].add <>Relationship(Target="sharedStrings.xml",
    Id=fmt"rId{relslen+1}", Type=relSharedStrScheme)
  e.sharedStrings = sstr

proc assignSheetInfo(e: Excel) =
  var mapRidName = initTable[string, string]()
  var mapFilenameRid = initTable[string, string]()
  for s in e.workbook.body.findAll "sheet":
    mapRidName[s.attr "r:id"] = s.attr "name"
  for s in e.workbook.rels[1]:
    mapFilenameRid[s.attr("Target").extractFilename] = s.attr "Id"
  for path, sheet in e.sheets:
    sheet.filename = path.extractFilename
    sheet.rid = mapFilenameRid[sheet.filename]
    sheet.privName = mapRidName[sheet.rid]

proc readSharedStrings(path: string, body: XmlNode): SharedStrings =
  result = SharedStrings(
    path: path,
    body: body,
    count: try: parseInt(body.attr "count") except: 0,
    unique: try: parseInt(body.attr "uniqueCount") except: 0,
  )

  var count = -1
  result.strtables = newTable[string, int](body.len)
  for node in body:
    let tnode = node.child "t"
    if tnode == nil: continue
    inc count
    result.strtables[tnode.innerText] = count

proc addEmptyStyles(e: Excel) =
  const path = "xl/styles.xml"
  e.content.add <>Override(PartName="/" & path,
    ContentType=spreadtypefmt % ["styles"])
  let relslen = e.workbook.rels[1].len
  e.workbook.rels[1].add <>Relationship(Target="styles.xml",
    Id=fmt"rId{relslen+1}", Type=relStylesScheme)
  let styles = <>stylesSheet(xmlns=mainns,
    <>numFmts(count="1", <>numFmt(formatCode="General", numFmtId="164")),
    <>fonts(count= $0),
    <>fills(count="1", <>fill(<>patternFill(patternType="none"))),
    <>borders(count= $1, <>border(diagonalUp="false", diagonalDown="false",
      <>begin(), newXmlTree("end", []), <>top(), <>bottom())),
    <>cellStyleXfs(count= $0),
    <>cellXfs(count= $0),
    <>cellStyles(count= $0),
    <>colors(<>indexedColors()))
  e.otherfiles["styles.xml"] = (path, styles)

template fetchStyles(row: Row): XmlNode =
  let (a, r) = row.sheet.parent.otherfiles["styles.xml"]
  discard a
  r

template retrieveColor(color: string): untyped =
  let r = if color.startsWith("#"): color[1..^1] else: color
  "FF" & r

proc toXmlNode(f: Font): XmlNode =
  result = <>font(<>name(val=f.name), <>sz(val= $f.size))
  template addElem(test, field: untyped): untyped =
    if `test`:
      result.add <>`field`(val= $f.`field`)

  addElem f.family >= 0, family
  addElem f.charset >= 0, charset
  addElem f.strike, strike
  addElem f.outline, outline
  addElem f.shadow, shadow
  addElem f.condense, condense
  addElem f.extend, extend
  if f.bold: result.add <>b(val= $f.bold)
  if f.italic: result.add <>i(val= $f.italic)
  if f.color != "": result.add <>color(rgb = retrieveColor(f.color))
  result.add <>u(val= $f.underline)
  result.add <>vertAlign(val= $f.verticalAlign)

proc addFont(styles: XmlNode, font: Font): (int, bool) =
  if font.name == "": return
  let fontnode = font.toXmlNode
  let applyFont = true

  let fonts = styles.retrieveChildOrNew "fonts"
  let fontCount = try: parseInt(fonts.attr "count") except: 0
  fonts.add fontnode
  fonts.attrs = {"count": $(fontCount+1)}.toXmlAttributes
  fonts.add fontnode
  let fontId = fontCount
  (fontId, applyFont)

proc addBorder(styles: XmlNode, border: Border): (int, bool) =
  if not border.edit: return

  let applyBorder = true
  let bnodes = styles.retrieveChildOrNew "borders"
  let bcount = try: parseInt(bnodes.attr "count") except: 0
  let borderId = bcount

  let bnode = <>border(diagonalUp= $border.diagonalUp,
    diagonalDown= $border.diagonalDown)

  template addBorderProp(fname: string, field: untyped) =
    let fld = border.`field`
    let elem = newXmlTree(fname, [], newStringTable())
    if fld.edit:
      elem.attrs["style"] = $fld.style
      if fld.color != "":
        elem.add <>color(rgb = retrieveColor(fld.color))
    bnode.add elem

  addBorderProp("start", start)
  addBorderProp("end", `end`)
  addBorderProp("top", top)
  addBorderProp("bottom", bottom)
  addBorderProp("vertical", vertical)
  addBorderProp("horizontal", horizontal)

  bnodes.attrs["count"] = $(bcount+1)
  bnodes.add bnode

  (borderId, applyBorder)

proc addPattern(fillnode: XmlNode, patt: PatternFill) =
  if not patt.edit: return
  let patternNode = <>patternFill(patternType= $patt.patternType)

  if patt.fgColor != "":
    patternNode.add <>fgColor(rgb = retrieveColor(patt.fgColor))
  if patt.bgColor != "":
    patternNode.add <>bgColor(rgb = retrieveColor(patt.bgColor))

  fillnode.add patternNode

proc addGradient(fillnode: XmlNode, grad: GradientFill) =
  if not grad.edit: return
  let gradientNode = newXmlTree("gradientFill", [
    <>stop(position= $grad.stop.position,
      <>color(rgb= retrieveColor(grad.stop.color)))
  ], {
    "type": $grad.`type`,
    "degree": $grad.degree,
    "left": $grad.left,
    "right": $grad.right,
    "top": $grad.top,
    "bottom": $grad.bottom,
  }.toXmlAttributes)

  fillnode.add gradientNode

proc addFill(styles: XmlNode, fill: Fill): (int, bool) =
  if not fill.edit: return

  result[1] = true
  let fills = styles.retrieveChildOrNew "fills"
  let count = try: parseInt(fills.attr "count") except: 0

  let fillnode = <>fill()
  fillnode.addPattern fill.pattern
  fillnode.addGradient fill.gradient

  fills.attrs["count"] = $(count+1)
  fills.add fillnode
  result[0] = count

proc border*(start, `end`, top, bottom, vertical, horizontal = BorderProp();
  diagonalUp, diagonalDown = false): Border =
  ## Border initializer. Use this instead of object constructor
  ## to indicate style is ready to apply this border.
  runnableExamples:
    import std/with
    import excelin

    var b = border(diagonalUp = true)
    with b:
      start = borderProp(style = bsMedium) # border style
      diagonalDown = true

    doAssert b.diagonalUp
    doAssert b.diagonalDown
    doAssert b.start.style == bsMedium

  Border(
    edit: true,
    start: start,
    `end`: `end`,
    top: top,
    bottom: bottom,
    vertical: vertical,
    horizontal: horizontal,
    diagonalUp: diagonalUp,
    diagonalDown: diagonalDown)

proc borderProp*(style = bsNone, color = ""): BorderProp =
  BorderProp(edit: true, style: style, color: color)

proc fillStyle*(pattern = PatternFill(), gradient = GradientFill()): Fill =
  Fill(edit: true, pattern: pattern, gradient: gradient)

proc patternFill*(fgColor = $colWhite; patternType = ptNone): PatternFill =
  PatternFill(edit: true, fgColor: fgColor, bgColor: "",
    patternType: patternType)

proc gradientFill*(stop = GradientStop(), `type` = gtLinear,
  degree, left, right, top, bottom = 0.0): GradientFill =
  GradientFill(
    edit: true,
    stop: stop,
    `type`: `type`,
    degree: degree,
    left: left,
    right: right,
    top: top,
    bottom: bottom,
  )


# To add style need to update:
# ✗ numFmts
# ✓ fonts
# ✗ borders
# ✗ fills
# ✗ cellStyleXfs
# ✓ cellXfs (the main reference for style in cell)
# ✗ cellStyles (named styles)
# ✗ colors (if any)
proc style*(row: Row, col: string,
  font = Font(size: 1),
  border = Border(),
  fill = Fill(),
  alignment: openarray[(string, string)] = []) =
  ## Add style to cell in row by selectively providing the font, border, fill
  ## and alignment styles.
  let sparse = $cfSparse == row.body.attr "cellfill"
  let rnum = row.rowNum
  var pos = -1
  var c =
    if not sparse:
      pos = col.toNum
      row.body[pos]
    else:
      var x: XmlNode
      for node in row.body:
        inc pos
        if fmt"{col}{rnum}" == node.attr "r" :
          x = node
          break
      x
  if c == nil:
    pos = row.body.len
    row[col] = ""
    c = row.body[pos]

  let styles = row.fetchStyles
  let (fontId, applyFont) = styles.addFont font
  let styleId = try: parseInt(c.attr "s") except: 0
  let applyAlignment = alignment.len > 0
  let (borderId, applyBorder) = styles.addBorder border
  let (fillId, applyFill) = styles.addFill fill

  let csxfs = styles.retrieveChildOrNew "cellStyleXfs"
  let cxfs = styles.child "cellXfs"
  if cxfs == nil: return
  let xfid = cxfs.len
  let xf =
    if styleId == 0:
      <>xf(applyProtection="false", applyAlignment= $applyAlignment, applyFont= $applyFont,
        numFmtId="0", borderId= $borderId, fillId= $fillId, fontId= $fontId, applyBorder= $applyBorder)
    else:
      cxfs[styleId]
  let alignNode =
    if styleId == 0:
      <>alignment(shrinkToFit="false", indent="0", vertical="bottom",
        horizontal="general", textRotation="0", wrapText="false")
    else:
      xf.child "alignment"
  let protecc = if styleId == 0: <>protection(hidden="false", locked="true")
                else: xf.child "protection"

  for (k, v) in alignment:
    alignNode.attrs[k] = v

  if styleId == 0:
    let cxfscount = try: parseInt(cxfs.attr "count") except: 0
    cxfs.attrs["count"] = $(cxfscount+1)
    csxfs.attrs["count"] = $(csxfs.len + 1)
    xf.add alignNode
    xf.add protecc
    cxfs.add xf
    csxfs.add xf
    c.attrs["s"] = $xfid
  else:
    if font.name != "":
      xf.attrs["fontId"] = $fontId
      xf.attrs["applyFont"] = $applyFont
    xf.attrs["applyAlignment"] = $true
    if border.edit:
      xf.attrs["applyBorder"] = $true
      xf.attrs["borderId"] = $borderId
    if applyFill: xf.attrs["fillId"] = $fillId

proc colrow(cr: string): (string, int) =
  var rowstr: string
  for i, c in cr:
    if c in Letters:
      result[0] &= c
    elif c in Digits:
      rowstr = cr[i .. ^1]
      break
  result[1] = try: parseInt(rowstr) except: 0

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

proc shareStyle*(row: Row, col: string, targets: varargs[string]) =
  ## Share style from source row and col string to any arbitrary cells
  ## in format {Col}{Num} e.g. A1, B2, C3 etc. Changing the shared
  ## style will affect entire cells that has shared style.
  let cnode = row.retrieveCell col
  if cnode == nil: return
  let sid = cnode.attr "s"
  if sid == "" or sid == "0": return

  for cr in targets:
    let (tgcol, tgrow) = cr.colrow
    #let ctgt = row.sheet.row(tgrow).retrieveCell tgcol
    let tgtrownode = row(row.sheet, tgrow)
    let ctgt = tgtrownode.retrieveCell tgcol
    if ctgt == nil: continue
    ctgt.attrs["s"] = sid

proc shareStyle*(sheet: Sheet, source: string, targets: varargs[string]) =
  ## Share style from source {col}{row} to targets {col}{row},
  ## i.e. `sheet.shareStyle("A1", "B2", "C3")`
  ## which shared the style in cell A1 to B2 and C3.
  let (sourceCol, sourceRow) = source.colrow
  let row = sheet.row sourceRow
  row.shareStyle sourceCol, targets

proc copyTo(src, dest: XmlNode) =
  if src == nil or dest == nil or
    dest.kind != xnElement or src.kind != xnElement or
    src.tag != dest.tag:
    return

  for k, v in src.attrs:
    dest.attrs[k] = v

  for child in src:
    let newchild = newXmlTree(child.tag, [], newStringTable())
    child.copyTo newchild
    dest.add newchild


proc copyStyle*(row: Row, col: string, targets: varargs[string]) =
  ## Copy style from row and col source to targets. The difference
  ## with `shareStyle proc<#shareStyle<Row,string,varargs[]>`_ is
  ## copy will make a new same style. So changing targets cell
  ## style later won't affect the source and vice versa.
  let cnode = row.retrieveCell col
  if cnode == nil: return
  let sid = cnode.attr "s"
  if sid == "" or sid == "0": return
  let styles = row.fetchStyles
  if styles == nil: return
  let cxfs = styles.child "cellXfs"
  if cxfs == nil or cxfs.len < 1: return
  let csxfs = styles.retrieveChildOrNew "cellStyleXfs"
  let stylepos = try: parseInt(sid) except: -1
  if stylepos < 0 or stylepos >= cxfs.len: return
  var stylescount = cxfs.len
  let refxf = cxfs[stylepos]

  var count = 0
  for cr in targets:
    let (tgcol, tgrow) = cr.colrow
    let ctgt = row.sheet.row(tgrow).retrieveCell tgcol
    if ctgt == nil: continue

    let newxf = newXmlTree("xf", [], newStringTable())
    refxf.copyTo newxf

    ctgt.attrs["s"] = $stylescount
    inc stylescount
    inc count

    cxfs.add newxf
    csxfs.add newxf

  cxfs.attrs["count"] = $stylescount
  csxfs.attrs["count"] = $(csxfs.len+count)

proc copyStyle*(sheet: Sheet, source: string, targets: varargs[string]) =
  ## Copy style from source {col}{row} to targets {col}{row},
  ## i.e. `sheet.shareStyle("A1", "B2", "C3")`
  ## which copied style from cell A1 to B2 and C3.
  let (sourceCol, sourceRow) = source.colrow
  let row = sheet.row sourceRow
  row.copyStyle sourceCol, targets

template `$`(r: Range): string =
  var dim = r[0]
  if r[1] != "":
    dim &= ":" & r[1]
  dim

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

proc assignSheetRel(excel: Excel) =
  for k in excel.sheets.keys:
    let (path, fname) = splitPath k
    let relname = fmt"{fname}.rels"
    if  relname in excel.otherfiles: continue
    let rels = <>Relationships(xmlns=relPackageSheet)
    let relspathname = (path / "_rels" / relname).unixSep
    excel.rels.add <>Override(PartName="/" & relspathname,
      ContentType= packagetypefmt % ["relationships"])
    excel.otherfiles[relname] = (relspathname, rels)


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
      result.sharedStrings = path.readSharedStrings(extract path)
    elif wbpath.endsWith(".xml") or wbpath.endsWith(".rels"): # any others xml/rels files
      let (_, f) = splitPath wbpath
      result.otherfiles[f] = path.fileRep
    else:
      let (_, f) = splitPath wbpath
      result.embedfiles[f] = (path, reader.extractFile path)
  if not found:
    raise newException(ExcelError, "No workbook found, invalid excel file")
  if result.sharedStrings == nil:
    result.addSharedStrings
  result.assignSheetInfo
  result.assignSheetRel

  if "app.xml" in result.otherfiles:
    var (_, appnode) = result.otherfiles["app.xml"]
    clear appnode
    appnode.add <>Application(newText "Excelin")
    appnode.add <>AppVersion(newText excelinVersion)

  if "styles.xml" notin result.otherfiles:
    result.addEmptyStyles

proc `prop=`*(e: Excel, prop: varargs[(string, string)]) =
  ## Add information property to Excel file. Will add the properties
  ## to the existing.
  const key = "app.xml"
  if key notin e.otherfiles: return
  let (_, propnode) = e.otherfiles[key]
  for p in prop:
    propnode.add newXmlTree(p[0], [newText p[1]])

proc createdAt*(excel: Excel, at: DateTime|Time = now()) =
  ## Set the created at properties to our excel.
  ## Useful when we're creating an excel from template so
  ## we set the creation date to current date which different
  ## with template created date.
  const core = "core.xml"
  if core in excel.otherfiles:
    let (_, cxml) = excel.otherfiles[core]
    if cxml != nil:
      var created = cxml.retrieveChildOrNew "dcterms:created"
      clear created
      created.add newText(at.format datefmt)

proc newExcel*(appName = "Excelin"): (Excel, Sheet) =
  ## Return a new Excel and Sheet at the same time to work for both.
  ## The Sheet returned is by default has name "Sheet1" but user can
  ## use `name= proc<#name=,Sheet,string>`_ to change its name.
  let excel = readExcel emptyxlsx
  excel.createdAt
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
  e.sharedStrings.path.addContents $e.sharedStrings.body
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
  import std/with
  let (empty, sheet) = newExcel()

  let colnum = [("A", 0), ("AA", 26), ("AB", 27), ("ZZ", 701)]
  for cn in colnum:
    doAssert cn[0].toNum == cn[1]
    doAssert cn[1].toCol == cn[0]

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
    row["D"] = now() # modify to date time
    #empty.deleteSheet "hehe"
    #empty.writeFile "generate-deleted-sheet.xlsx"
    let row5 = sheet.row 5
    row5["A"] = "yeehaa"
    row5.style("A",
      Font(name: "DejaVu Sans Mono", size: 11,
        family: -1, charset: -1,
        color: $colBlue,
      ),
      border(
        top = borderProp(style = bsMedium, color = $colRed),
        bottom = borderProp(style = bsMediumDashDot, color = $colGreen),
      ),
      fillStyle(
        pattern = patternFill(patternType = ptDarkGray, fgColor = $colRed)
      ),
      alignment = {"horizontal": "center", "vertical": "center",
        "wrapText": $true, "textRotation": $45})
    row5.height = 200
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
    row5.style("B", alignment = {"wrapText": "true"})
    let row11 = sheet.row 11
    var sum = 0
    for i in 0 .. 9:
      row11[i.toCol] = i
      sum += i
    row11[10.toCol] = Formula(equation: "SUM(A11:J11)", valueStr: $sum)
    dump row11[10.toCol, Formula]

    let row12 = sheet.row 12
    row12.style("C", Font(name: "Cambria", size: 11, family: -1, charset: -1),
      alignment = {"horizontal": "center", "vertical": "center", "wrapText": "true",
      "textRotation": "90"})
    row5.style "A", alignment = {"textRotation": $90} # edit existing style
    template hideLevel(rnum, olevel: int): untyped =
      let r = sheet.row rnum
      with r:
        #hide = true
        outlineLevel = olevel
      r
    let row13 = 13.hideLevel 3
    let row14 = 14.hideLevel 2
    let row15 = 15.hideLevel 1
    let row16 = 16.hideLevel 1
    row16.collapsed = false
    row16.hide = false

    for i in 0 ..< 5:
      let colstyle = i.toCol
      row16[colstyle] = colstyle

    row16.style "A", font = Font(name: "DejaVu Sans Mono", size: 11,
      family: -1, charset: -1)

    row16.shareStyle("A", "B16", "C16", "D16", "E16")
    sheet.shareStyle("A16", "B13", "C14", "D15")
    row13["B"] = "bebebe"
    row14["C"] = "cecece"
    row15["D"] = "dedede"

    row16.copyStyle("A", "C13", "D14", "E15")
    row13["C"] = "copied style from A16"
    row14["D"] = "copied style from A16"
    row15["E"] = "copied style from A16"

    #dump sheet.body
    #dump empty.sharedStrings.body
    #dump empty.otherfiles["app.xml"]
    #dump empty.otherfiles["styles.xml"]
    empty.writeFile "generated-add-rows.xlsx"
  else:
    echo "sheet is nil"

  {.hint[XDeclaredButNotUsed]:off.}

  proc autofiltertest =
    proc populateRow(row: Row, col, cat: string, data: array[3, float]) =
      let startcol = col.toNum + 1
      row[col] = cat
      var sum = 0.0
      for i, d in data:
        row[(startcol+i).toCol] = d
        sum += d
      let rnum = row.rowNum
      let eqrange = fmt"SUM({col}{rnum}:{(startcol+data.len-1).toCol}{rnum})" 
      dump eqrange
      row[(startcol+data.len).toCol] = Formula(equation: eqrange, valueStr: $sum)
    let (excel, sheet) = newExcel()
    let row5 = sheet.row 5
    let startcol = "D".toNum
    for i, s in ["Category", "Num1", "Num2", "Num3", "Total"]:
      row5[(startcol+i).toCol] = s
    let row6 = sheet.row 6
    row6.populateRow("D", "A", [0.18460660235998017, 0.93463071023892952, 0.58647760893211043])
    sheet.row(7).populateRow("D", "A", [0.50425224796279555, 0.25118866081991786, 0.26918159410869791])
    sheet.row(8).populateRow("D", "A", [0.6006019062877066, 0.18319235857964333, 0.12254334000604317])
    sheet.row(9).populateRow("D", "A", [0.78015011938458589, 0.78159963723670689, 6.7448346870105036E-2])
    sheet.row(10).populateRow("D", "B", [0.63608141933645479, 0.35635845012920608, 0.67122053637107193])
    sheet.row(11).populateRow("D", "B", [0.33327331908137214, 0.2256497329592122, 0.5793989116090501])

    sheet.ranges = ("D5", "H11")
    sheet.autoFilter = ("D5", "H11")
    sheet.filterCol 0, Filter(kind: ftFilter, valuesStr: @["A"])
    sheet.filterCol 1, Filter(kind: ftCustom, logic: cflAnd,
      customs: @[(foGt, $0), (foLt, $0.7)])
    #dump sheet.autoFilter
    #dump sheet.body
    excel.writeFile "generated-autofilter.xlsx"

  autofiltertest()
