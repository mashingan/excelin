# internal include dependency is as below:
# internal_sheets <- internal_styles <- internal_rows <-
#   internal_cells <- internal_utilities <- internal_types
# all files with prefix internal_ are considered as package
# wide implementation hence all internal privates are shared.
include excelin/internal_sheets

from std/xmlparser import parseXml
from std/sha1 import secureHash, `$`

from zippy/ziparchives import openZipArchive, extractFile, ZipArchive,
  ArchiveEntry, writeZipArchive, ZippyError

const
  spreadtypefmt = "application/vnd.openxmlformats-officedocument.spreadsheetml.$1+xml"
  relSharedStrScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  relStylesScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  emptyxlsx = currentSourcePath.parentDir() / "empty.xlsx"

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
  var
    workbookfound = false
    workbookrelsExists = false
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
      workbookfound = true
    elif "worksheet" in contentType:
      inc result.sheetCount
      let sheet = extract path
      result.sheets[path] = Sheet(body: sheet, parent: result)
    elif wbpath.endsWith "workbook.xml.rels":
      workbookrelsExists = true
      result.workbook.rels = fileRep path
    elif wbpath.endsWith "sharedStrings.xml":
      result.sharedStrings = path.readSharedStrings(extract path)
    elif wbpath.endsWith(".xml") or wbpath.endsWith(".rels"): # any others xml/rels files
      let (_, f) = splitPath wbpath
      result.otherfiles[f] = path.fileRep
    else:
      let (_, f) = splitPath wbpath
      result.embedfiles[f] = (path, reader.extractFile path)
  if not workbookfound:
    raise newException(ExcelError, "No workbook found, invalid excel file")
  if not workbookrelsExists:
    const relspath = "xl/_rels/workbook.xml.rels"
    try:
      result.workbook.rels = fileRep relspath
    except ZippyError:
      raise newException(ExcelError, "Invalid excel file, no workbook relations exists")
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
