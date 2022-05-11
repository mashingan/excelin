include internal_types

template unixSep(str: string): untyped = str.replace('\\', '/')
  ## helper to change the Windows path separator to Unix path separator

proc retrieveChildOrNew(node: XmlNode, name: string): XmlNode =
  var r = node.child name
  if r == nil:
    r = newXmlTree(name, [], newStringTable())
    node.add r
  r

proc colrow(cr: string): (string, int) =
  var rowstr: string
  for i, c in cr:
    if c in Letters:
      result[0] &= c
    elif c in Digits:
      rowstr = cr[i .. ^1]
      break
  result[1] = try: parseInt(rowstr) except: 0

template styleRange(sheet: Sheet, `range`: Range, op: untyped) =
  let
    (tlcol, tlrow) = `range`[0].colrow
    (btcol, btrow) = `range`[1].colrow
    r = sheet.row tlrow
  var targets: seq[string]
  for cn in tlcol.toNum+1 .. btcol.toNum:
    let col = cn.toCol
    targets.add col & $tlrow
  for rnum in tlrow+1 .. btrow:
    for cn in tlcol.toNum .. btcol.toNum:
      targets.add cn.toCol & $rnum
  r.`op`(tlcol, targets)

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

const atoz = toseq('A'..'Z')

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

template `$`*(r: Range): string =
  var dim = r[0]
  if r[1] != "":
    dim &= ":" & r[1]
  dim

proc rowNum*(r: Row): Positive =
  ## Getting the current row number of Row object users working it.
  result = try: parseInt(r.body.attr "r") except: 1
