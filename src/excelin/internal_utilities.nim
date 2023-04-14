include internal_types

from std/times import DateTime, Time, now, format, toTime, toUnixFloat,
  parse, fromUnix, local
from std/sequtils import toSeq, mapIt, repeat
from std/strutils import endsWith, contains, parseInt, `%`, replace,
  parseFloat, parseUint, toUpperAscii, join, startsWith, Letters, Digits,
  parseBool, parseEnum
from std/math import `^`
from std/strtabs import `[]=`, pairs, newStringTable, del

const
  datefmt = "yyyy-MM-dd'T'HH:mm:ss'.'fffzz"

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
  result = try: parseInt(r.body.attr "r") except ValueError: 1

proc colrow(cr: string): (string, int) =
  var rowstr: string
  for i, c in cr:
    if c in Letters:
      result[0] &= c
    elif c in Digits:
      rowstr = cr[i .. ^1]
      break
  result[1] = try: parseInt(rowstr) except ValueError: 0
