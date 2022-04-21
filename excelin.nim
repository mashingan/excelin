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
     newXmlTree, add, newText, toXmlAttributes, delete, XmlAttributes,
     len, xmlHeader, attrs, `attrs=`, innerText, `[]`
from std/xmlparser import parseXml, XmlError
from std/strutils import endsWith, join, contains, parseInt, `%`, replace,
  parseFloat, parseUint
from std/sequtils import toSeq, map, mapIt
from std/tables import TableRef, newTable, `[]`, `[]=`, contains, pairs,
     keys, del, values, initTable, len
from std/strformat import fmt
from std/times import DateTime, Time, now, format, toTime, toUnix, parse
from std/os import splitFile, `/`, addFileExt, parentDir, splitPath,
  getTempDir, removeFile, extractFilename, relativePath, tailDir
from std/uri import parseUri, `/`, `$`
from std/strtabs import StringTableRef, `[]=`
from std/sugar import dump, `->`

from zippy/ziparchives import openZipArchive, extractFile, ZipArchive,
  ArchiveEntry, writeZipArchive

#{.hint[XDeclaredButNotUsed]: off.}


const
  datefmt = "yyyy-MM-dd'T'hh:mm:ss'.'fffzz"
  xmlnsx14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlnsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlnsxdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlnsmc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
  spreadtypefmt = "application/vnd.openxmlformats-officedocument.spreadsheetml.$1+xml"
  mainns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  relSharedStrScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  emptyxlsx = currentSourcePath.parentDir() / "empty.xlsx"
  excelinVersion* = "0.1.0"

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
    ## The main object that will be used most of of the time for many users.
    ## This object will represent a sheet in Excel file such as adding row
    ## getting row, and/or adding cell directly.
    sharedStrings: XmlNode
    parent: Excel
    rid: string
    privName: string

  Row* = ref object of InternalBody
    ## The object that will be used for working with values within cells of a row.
    ## Users can get the value within cell and set the its value.
    sheet: Sheet
  FilePath = string
  FileRep = (FilePath, XmlNode)

  ExcelError* = object of CatchableError

template unixSep*(str: string): untyped = str.replace('\\', '/')
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

proc addRow*(s: Sheet): Row =
  ## Add row directly from sheet after any existing row.
  ## This will return a new pristine row to work further.
  ## Row numbering is 1-based.
  let rows = toSeq s.body.findAll("row")
  let newnum = block:
    if rows.len > 0:
      try: parseInt(rows[^1].attr "r")+1 except: 1
    else: 1
  result = Row(sheet: s)
  result.body = <>row(r= $newnum, hidden="false", collapsed="false")
  let sdata = s.body.getSheetData
  sdata.add result.body
  s.modifiedAt

proc addRow*(s: Sheet, rowNum: Positive): Row =
  ## Add row by selecting which row number to work with.
  ## This will return new row if there's no existing row
  ## or will return an existing one.
  var latestRowNum = 0
  for r in s.body.findAll "row":
    let rnum = try: parseInt(r.attr "r") except: 0
    if rnum == rowNum:
      return Row(body: r, sheet: s)
    latestRowNum = max(latestRowNum, rnum)
  result = Row(sheet: s)
  result.body = <>row(r= $latestRowNum, hidden="false", collapsed="false")
  let sdata = s.body.getSheetData
  sdata.add result.body
  s.modifiedAt

proc rowNum*(r: Row): Positive =
  ## Getting the current row number of Row object users working it.
  result = try: parseInt(r.body.attr "r") except: 1

proc addCell*(row: Row, col: string, s: string) =
  ## Add cell with overload for value string. Supplied column
  ## is following the Excel convention starting from A -  Z, AA - AZ ...
  let lastStr = row.sheet.sharedStrings.len
  row.sheet.sharedStrings.add <>si(newXmlTree("t",
    [newText s],
    {"xml:space": "preserve"}.toXmlAttributes))
  row.body.add <>c(r=fmt"""{col}{row.body.attr "r"}""",
    t="s", s="0", <>v(newText $lastStr))
  row.sheet.modifiedAt
  let newcount = lastStr + 1
  row.sheet.sharedStrings.attrs = {"count": $newcount,
    "uniqueCount": $newcount, "xmlns": mainns}.toXmlAttributes
proc addCell*(row: Row, col: string, n: SomeNumber) =
  ## Add cell with overload for any number.
  let cellpos = fmt"""{col}{row.body.attr "r"}"""
  row.body.add <>c(r=cellpos, t="n", <>v(newText $n))
  row.sheet.modifiedAt
proc addCell*(row: Row, col: string, b: bool) =
  ## Add cell with overload for truthy value.
  row.body.add <>c(r=fmt"""{col}{row.body.attr "r"}""", t="b", <>v(newText $b))
  row.sheet.modifiedAt
proc addCell*(row: Row, col: string, d: DateTime | Time) =
  ## Add cell with overload for DateTime or Time. The saved value
  ## will be in string format of `yyyy-MM-dd'T'hh:mm:ss'.'fffzz` e.g.
  ## `2200-10-01T11:22:33.456-03`.
  row.body.add <>c(r=fmt"""{col}{row.body.attr "r"}""", t="d",
              <>v(newText d.format(datefmt)))
  row.sheet.modifiedAt

proc getCell*[R](row: Row, col: string, conv: string -> R = nil): R =
  ## Get cell value from row with optional function to convert it.
  ## When conversion function is supplied, it will be used instead of
  ## default conversion function viz:
  ## * SomeSignedInt (int/int8/int32/int64): strutils.parseInt
  ## * SomeUnsignedInt (uint/uint8/uint32/uint64): strutils.parseUint
  ## * SomeFloat (float/float32/float64): strutils.parseFloat
  ## * DateTime | Time: times.format with layout `yyyy-MM-dd'T'hh:mm:ss'.'fffzz`
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
  ##   let dtCustom = row.getCell[:DateTime]("F", (s: string) => s.parse("yyyy/MM/dd"))
  ##
  ## Any other type that other than mentioned above should provide the closure proc
  ## for the conversion otherwise it will return the default value, for example
  ## any ref object will return nil or for object will get the object with its field
  ## filled with default values.
  let rnum = row.body.attr "r"
  let v = block:
    var x: XmlNode
    for node in row.body:
      if fmt"{col}{rnum}" == node.attr "r":
        x = node
        break
    x
  if v == nil: return
  when R is SomeSignedInt: result = int.low
  elif R is SomeUnsignedInt: result = uint.high
  elif R is SomeFloat: result = NaN
  else: discard
  let t = v.innerText
  template retconv =
    if conv != nil: return conv t
  when R is string:
    retconv()
    let refpos = try: parseInt(t) except: -1
    if refpos == -1: return # failed to find the shared string pos
    let tnode = row.sheet.sharedStrings[refpos]
    if tnode == nil: return
    result = tnode.innerText
  elif R is SomeSignedInt:
    retconv()
    try: result = parseInt(t)
    except: discard
  elif R is SomeFloat:
    retconv()
    try: result = parseFloat(t)
    except: discard
  elif R is SomeUnsignedInt:
    retconv()
    try: result = parseUint t
    except: discard
  elif R is DateTime:
    retconv()
    result = try: parse(t, datefmt) except: discard
  elif R is Time:
    retconv()
    result = try: parse(t, datefmt).toTime except: discard
  else:
    discard

# when adding new sheet, need to update workbook to add
# ✓ to sheets,
# ✓ its new id,
# ✓ its package relationships
# ✓ add entry to content type
# ✗ add complete worksheet nodes
proc addSheet*(e: Excel, name = ""): Sheet =
  # Add new sheet to excel with supplied name and return it to enable further working.
  # The name can use the existing sheet name.
  const
    sheetTypeNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    contentTypeNs = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  inc e.sheetCount
  var name = name
  if name == "":
    name = fmt"Sheet{e.sheetCount}"
  let wbsheets = e.workbook.body.child "sheets"
  if wbsheets == nil:
    dec e.sheetCount # failed to add new sheet, return nil
    return
  let
    rel = e.workbook.rels
    availableId = rel[1].findAll("Relationship").len + 1
    rid = fmt"rId{availableId}"
    sheetfname = fmt"sheet{e.sheetCount}"
    fpath = block:
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
    {"name": name, "sheetId": $e.sheetCount, "r:id": rid, "state": "visible"}.toXmlAttributes)

  result = Sheet(
    body: worksheet,
    sharedStrings: e.sharedStrings[1],
    privName: name,
    rid: rid)
  e.sheets[fpath] = result
  wbsheets.add sheetworkbook
  e.workbook.sheetsInfo.add result.body
  rel[1].add <>Relationship(Target=fpath, Type=sheetTypeNs, Id=rid)
  e.content.add <>Override(PartName="/" & fpath, ContentType=contentTypeNs)

# deleteing sheet needs to delete several related info viz:
# ✓ deleting from the sheet table
# ✓ deleting from content
# ✓ deleting from package relationships
# ✓ deleting from workbook body
# ✗ deleting from workbook sheets info
proc deleteSheet*(e: Excel, name = "") =
  ## Delete sheet based on its name. Will ignore when it's cannot find the name.
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
  dec e.sheetCount

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
    elif wbpath.endsWith(".xml"): # any others xml files
      let (_, f) = splitPath wbpath
      result.otherfiles[f] = path.fileRep
  if not found:
    raise newException(ExcelError, "No workbook found, invalid excel file")
  if result.sharedStrings[1] == nil:
    result.addSharedStrings
  result.assignSheetInfo
  for _, s in result.sheets:
    s.sharedStrings = result.sharedStrings[1]

proc addAppProp*(e: Excel, prop: varargs[(string, string)]) =
  ## Add information property to Excel file. Will overwrite
  ## any existing information property.
  const key = "app.xml"
  if key notin e.otherfiles: return
  let (_, propnode) = e.otherfiles[key]
  for p in prop:
    propnode.add newXmlTree(p[0], [newText p[1]])

proc newExcel*(appName = "Excelin"): (Excel, Sheet) =
  ## Return a new Excel and Sheet at the same time to work for both.
  ## The Sheet returned is by default has name "Sheet1" but user can
  ## use `name= proc<#name=,Sheet>`_ to change its name.
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

  archive.writeZipArchive targetpath

proc `$`*(e: Excel): string =
  ## Get Excel file as string. Currently implemented by writing to
  ## temporary dir first because there's no API to get the data
  ## directly.
  let path = getTempDir() / fmt"excelin-{now().toTime.toUnix}.xml"
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
      var currattr = StringTableRef(node.attrs)
      currattr["name"] = newname
      node.attrs = currattr


when isMainModule:
  let tworows = readExcel "2-rows.xlsx"
  let (empty, sheet) = newExcel()
  empty.writeFile "generate-empty.xlsx"
  dump tworows.sheetNames
  if sheet != nil:
    let row = sheet.addRow
    row.addCell "A", "hehe"
    row.addCell "B", -1
    row.addCell "C", 2
    row.addCell "D", 42.0
    #dump sheet.name
    sheet.name = "hehe"
    #dump empty.sheetNames
    let newsheet = empty.addSheet("test add new sheet")
    dump newsheet.name
    #dump empty.sheetNames
    #dump newsheet.body
    #dump empty.workbook.rels
    #dump empty.workbook.body
    empty.writeFile "generate-empty.xlsx"
    dump row.getCell[:string]("A")
    dump row.getCell[:int]("B")
    dump row.getCell[:uint]("C")
    dump row.getCell[:float]("D")
    empty.deleteSheet "hehe"
    empty.writeFile "generate-empty-sheet.xlsx"
  else:
    echo "sheet is nil"

  #echo $empty.workbook.body
  dump empty.sheetNames
