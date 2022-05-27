include internal_utilities

from std/strformat import fmt
from std/strscans import scanf
from std/sugar import dump, `->`

const
  mainns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  relHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

proc retrieveChildOrNew(node: XmlNode, name: string): XmlNode =
  var r = node.child name
  if r == nil:
    r = newXmlTree(name, [], newStringTable())
    node.add r
  r

proc fetchValNode(row: Row, col: string, isSparse: bool): XmlNode =
  var x: XmlNode
  let colnum = col.toNum
  let rnum = row.body.attr "r"
  if not isSparse and colnum < row.body.len:
    x = row.body[colnum]
  else:
    for node in row.body:
      if fmt"{col}{rnum}" == node.attr "r":
        x = node
        break
  x

proc fetchCell(body: XmlNode, colrow: string): int =
  var count = -1
  for n in body:
    inc count
    if colrow == n.attr "r": return count
  -1

proc addCell(row: Row, col, cellType, text: string, valelem = "v",
  altnode: seq[XmlNode] = @[], emptyCell = false, style = "0") =
  let rn = row.body.attr "r"
  let fillmode = try: parseEnum[CellFill](row.body.attr "cellfill") except: cfSparse
  let sparse = fillmode == cfSparse
  let col = col.toUpperAscii
  let cellpos = fmt"{col}{rn}"
  let innerval = if altnode.len > 0: altnode else: @[newText text]
  let cnode = if cellType != "" and valelem != "":
                <>c(r=cellpos, s=style, t=cellType, newXmlTree(valelem, innerval))
              elif valelem != "": <>c(r=cellpos, s=style, newXmlTree(valelem, innerval))
              elif emptyCell: <>c(r=cellpos, s=style)
              else: newXmlTree("c", innerval, {"r": cellpos, "s": style}.toXmlAttributes)
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

proc `[]=`*(row: Row, col: string, h: Hyperlink) =
  let sheetrelname = fmt"{row.sheet.filename}.rels"
  if sheetrelname notin row.sheet.parent.otherfiles: return
  let (_, sheetrel) = row.sheet.parent.otherfiles[sheetrelname]
  row[col] = h.text
  let hlinks = row.sheet.body.retrieveChildOrNew "hyperlinks"
  let colrnum = fmt"{col}{row.rowNum}"
  let ridn = sheetrel.len + 1 # rId is 1-based
  let rid = fmt"rId{ridn}"

  let hlink = newXmlTree("hyperlink", [], {
    "ref": colrnum,
    "r:id": rid}.toXmlAttributes)
  if h.tooltip != "": hlink.attrs["tooltip"] = h.tooltip
  hlinks.add hlink

  sheetrel.add <>Relationship(Type=relHyperlink, Target=h.target,
    Id=rid, TargetMode="external")

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
  let fillmode = try: parseEnum[CellFill](row.body.attr "cellfill") except: cfSparse
  let isSparse = fillmode == cfSparse
  let col = col.toUpperAscii
  let v = row.fetchValNode(col, isSparse)
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
    result = if "inlineStr" == v.attr "t" : t else: fetchShared t
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
  elif R is Hyperlink:
    result.text = if "inlineStr" == v.attr "t": t else: fetchShared t
    let hlinks = row.sheet.body.retrieveChildOrNew "hyperlinks"
    var rid = ""
    for hlink in hlinks:
      if fmt"{col}{row.rowNum}" == hlink.attr "ref":
        result.tooltip = hlink.attr "tooltip"
        rid = hlink.attr "r:id"
        break
    let sheetrelname = fmt"{row.sheet.filename}.rels"
    if rid == "" or sheetrelname notin row.sheet.parent.otherfiles:
      return
    var ridn: int
    if not scanf(rid, "rId$i", ridn): return
    let (_, rels) = row.sheet.parent.otherfiles[sheetrelname]
    let rel = rels[ridn-1]
    result.target = rel.attr "Target"
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
