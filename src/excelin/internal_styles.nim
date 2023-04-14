include internal_rows

import std/macros
from std/colors import `$`, colWhite

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
  if f.underline != uNone:
    result.add <>u(val= $f.underline)
  if f.verticalAlign != vaBaseline:
    result.add <>vertAlign(val= $f.verticalAlign)

proc retrieveCell(row: Row, col: string): XmlNode =
  let fillmode = try: parseEnum[CellFill](row.body.attr "cellfill") except ValueError: cfSparse
  if fillmode == cfSparse:
    let colrow = fmt"{col}{row.rowNum}"
    ## TODO: fix misfetch the cell
    let (fetchpos, _) = row.body.fetchCell(colrow, col.toNum)
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
  let stylepos = try: parseInt(sid) except ValueError: -1
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

proc addFont(styles: XmlNode, font: Font): (int, bool) =
  if font.name == "": return
  let fontnode = font.toXmlNode
  let applyFont = true

  let fonts = styles.retrieveChildOrNew "fonts"
  let fontCount = try: parseInt(fonts.attr "count") except ValueError: 0
  fonts.attrs = {"count": $(fontCount+1)}.toXmlAttributes
  fonts.add fontnode
  let fontId = fontCount
  (fontId, applyFont)

proc addBorder(styles: XmlNode, border: Border): (int, bool) =
  if not border.edit: return

  let applyBorder = true
  let bnodes = styles.retrieveChildOrNew "borders"
  let bcount = try: parseInt(bnodes.attr "count") except ValueError: 0
  let borderId = bcount

  let bnode = <>border(diagonalUp= $border.diagonalUp,
    diagonalDown= $border.diagonalDown)

  macro addBorderProp(field: untyped): untyped =
    let elemtag = $field
    result = quote do:
      let fld = border.`field`
      let elem = newXmlTree(`elemtag`, [], newStringTable())
      if fld.edit:
        elem.attrs["style"] = $fld.style
        if fld.color != "":
          elem.add <>color(rgb = retrieveColor(fld.color))
      bnode.add elem
  addBorderProp start
  addBorderProp `end`
  addBorderProp top
  addBorderProp bottom
  addBorderProp vertical
  addBorderProp horizontal

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
  let count = try: parseInt(fills.attr "count") except ValueError: 0

  let fillnode = <>fill()
  fillnode.addPattern fill.pattern
  fillnode.addGradient fill.gradient

  fills.attrs["count"] = $(count+1)
  fills.add fillnode
  result[0] = count

# Had to add for API style consistency.
proc fontStyle*(name: string, size = 10,
  family, charset = -1,
  bold, italic, strike, outline, shadow, condense, extend = false,
  color = "", underline = uNone, verticalAlign = vaBaseline): Font =
  Font(
    name: name,
    size: size,
    family: family,
    color: color,
    charset: charset,
    bold: bold,
    italic: italic,
    strike: strike,
    outline: outline,
    shadow: shadow,
    condense: condense,
    extend: extend,
    underline: underline,
    verticalAlign: verticalAlign,
  )

proc borderStyle*(start, `end`, top, bottom, vertical, horizontal = BorderProp();
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

proc border*(start, `end`, top, bottom, vertical, horizontal = BorderProp();
  diagonalUp, diagonalDown = false): Border 
  {. deprecated: "use borderStyle" .} =
  borderStyle(start, `end`, top, bottom, vertical, horizontal,
    diagonalUp, diagonalDown)

proc borderPropStyle*(style = bsNone, color = ""): BorderProp =
  BorderProp(edit: true, style: style, color: color)

proc borderProp*(style = bsNone, color = ""): BorderProp
  {. deprecated: "use borderPropStyle" .} =
  borderPropStyle(style, color)

proc fillStyle*(pattern = PatternFill(), gradient = GradientFill()): Fill =
  Fill(edit: true, pattern: pattern, gradient: gradient)

proc patternFillStyle*(fgColor = $colWhite; patternType = ptNone): PatternFill =
  PatternFill(edit: true, fgColor: fgColor, bgColor: "",
    patternType: patternType)

proc patternFill*(fgColor = $colWhite; patternType = ptNone): PatternFill
  {. deprecated: "use patternFillStyle" .} =
  patternFillStyle(fgColor, patternType)

proc gradientFillStyle*(stop = GradientStop(), `type` = gtLinear,
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

proc gradientFill*(stop = GradientStop(), `type` = gtLinear,
  degree, left, right, top, bottom = 0.0): GradientFill 
  {. deprecated: "use gradientFillStyle" .} =
  gradientFillStyle(stop, `type`, degree, left, right, top, bottom)

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
  let fillmode = try: parseEnum[CellFill](row.body.attr "cellfill") except ValueError: cfSparse
  let sparse = cfSparse == fillmode
  let rnum = row.rowNum
  var pos = -1
  let refnum = fmt"{col}{rnum}"
  var c =
    if not sparse:
      pos = col.toNum
      row.body[pos]
    else:
      var x: XmlNode
      for node in row.body:
        inc pos
        if refnum == node.attr "r" :
          x = node
          break
      x
  if c == nil:
    pos = row.body.len
    row[col] = ""
    c = row.body[pos]

  let styles = row.fetchStyles
  let (fontId, applyFont) = styles.addFont font
  let styleId = try: parseInt(c.attr "s") except ValueError: 0
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
    let cxfscount = try: parseInt(cxfs.attr "count") except ValueError: 0
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

  let mcells = row.sheet.body.child "mergeCells"
  if mcells == nil: return
  for m in mcells:
    if m.attr("ref").startsWith refnum:
      var topleft, bottomright: string
      if not scanf(m.attr "ref", "$w:$w", topleft, bottomright):
        return
      styleRange(row.sheet, (topleft, bottomright), shareStyle)

proc shareStyle*(sheet: Sheet, source: string, targets: varargs[string]) =
  ## Share style from source {col}{row} to targets {col}{row},
  ## i.e. `sheet.shareStyle("A1", "B2", "C3")`
  ## which shared the style in cell A1 to B2 and C3.
  let (sourceCol, sourceRow) = source.colrow
  let row = sheet.row sourceRow
  row.shareStyle sourceCol, targets

proc copyStyle*(sheet: Sheet, source: string, targets: varargs[string]) =
  ## Copy style from source {col}{row} to targets {col}{row},
  ## i.e. `sheet.shareStyle("A1", "B2", "C3")`
  ## which copied style from cell A1 to B2 and C3.
  let (sourceCol, sourceRow) = source.colrow
  let row = sheet.row sourceRow
  row.copyStyle sourceCol, targets

proc resetStyle*(sheet: Sheet, targets: varargs[string]) =
  ## Reset any styling to default.
  for cr in targets:
    let (tgcol, tgrow) = cr.colrow
    let ctgt = sheet.row(tgrow).retrieveCell tgcol
    if ctgt == nil: continue
    ctgt.attrs["s"] = $0

proc resetStyle*(row: Row, targets: varargs[string]) =
  ## Reset any styling to default.
  row.sheet.resetStyle targets

proc toRgbColorStr(node: XmlNode): string =
  let color = node.child "color"
  if color != nil:
    let rgb = color.attr "rgb"
    if rgb.len > 1:
      result = "#" & rgb[2..^1]

{.hint[ConvFromXtoItselfNotNeeded]: off.}

proc toFont(node: XmlNode): Font =
  result = Font(size: 1)
  if node.tag != "font": return

  template fetchElem(nodename: string, field, fetch, default: untyped) =
    let nn = node.child nodename
    if nn != nil:
      let val {.inject.} = nn.attr "val"
      result.`field` = try: `fetch`(val) except CatchableError: `default`
    else:
      result.`field` = `default`

  fetchElem "name", name, string, ""
  fetchElem "family", family, parseInt, -1
  fetchElem "charset", charset, parseInt, -1
  fetchElem "sz", size, parseInt, 1
  fetchElem "b", bold, parseBool, false
  fetchElem "i", italic, parseBool, false
  fetchElem "strike", strike, parseBool, false
  fetchElem "outline", outline, parseBool, false
  fetchElem "shadow", shadow, parseBool, false
  fetchElem "condense", condense, parseBool, false
  fetchElem "extend", extend, parseBool, false
  result.color = node.toRgbColorStr
  fetchElem "u", underline, parseEnum[Underline], uNone
  fetchElem "vertAlign", verticalAlign, parseEnum[VerticalAlign], vaBaseline

proc toFill(node: XmlNode): Fill =
  result.edit = true
  let pattern = node.child "patternFill"
  if pattern != nil:
    result.pattern = PatternFill(edit: true)
    let fgnode = pattern.child "fgColor"
    if fgnode != nil:
      let fgColorRgb = fgnode.attr "rgb"
      result.pattern.fgColor = if fgColorRgb.len > 1: "#" & fgColorRgb[2..^1] else: ""
    let bgnode = pattern.child "bgColor"
    if bgnode != nil:
      let bgColorRgb = bgnode.attr "rgb"
      result.pattern.bgColor = if bgColorRgb.len > 1: "#" & bgColorRgb[2..^1] else: ""
    result.pattern.patternType = try: parseEnum[PatternType](pattern.attr "patternType") except ValueError: ptNone
  let gradient = node.child "gradientFill"
  if gradient != nil:
    let stop = gradient.child "stop"
    result.gradient = GradientFill(
      edit: true,
      `type`: try: parseEnum[GradientType]gradient.attr "type" except ValueError: gtLinear,
    )
    if stop != nil:
      result.gradient.stop = GradientStop(
          position: try: parseFloat(stop.attr "position") except ValueError: 0.0,
          color: stop.toRgbColorStr,
      )
    macro addelemfloat(field: untyped): untyped =
      let strfield = $field
      result = quote do:
        result.gradient.`field` = try: parseFloat(gradient.attr `strfield`) except ValueError: 0.0
    addelemfloat degree
    addelemfloat left
    addelemfloat right
    addelemfloat top
    addelemfloat bottom

template retrieveStyleId(row: Row, col, styleAttr, child: string, conv: untyped): untyped =
  var cnode: XmlNode
  let cr = fmt"{col}{row.rowNum}"
  retrieveCol(row.body, 0,
    cr == n.attr "r", cnode, (discard colstr; nil))
  if cnode == nil: return
  const stylename = "styles.xml"
  if stylename notin row.sheet.parent.otherfiles: return
  let (path, style) = row.sheet.parent.otherfiles[stylename]
  discard path
  let cxfs = style.retrieveChildOrNew "cellXfs"
  let sid = try: parseInt(cnode.attr "s") except ValueError: 0
  if sid >= cxfs.len: return
  let theid = try: parseInt(cxfs[sid].attr styleAttr) except ValueError: 0
  let childnode = style.retrieveChildOrNew child
  if theid >= childnode.len: return
  childnode[theid].`conv`

proc toBorder(node: XmlNode): Border =
  result = Border(edit: true)

  macro retrieveField(field: untyped): untyped =
    let fname = $field
    result = quote do:
      let child = node.child `fname`
      var b: BorderProp
      if child != nil:
        b.edit = true
        b.style = try: parseEnum[BorderStyle](child.attr "style") except ValueError: bsNone
        b.color = child.toRgbColorStr
      result.`field` = b
  retrieveField start
  retrieveField `end`
  retrieveField top
  retrieveField bottom
  retrieveField vertical
  retrieveField horizontal
  result.diagonalDown = try: parseBool(node.attr "diagonalDown") except ValueError: false
  result.diagonalUp = try: parseBool(node.attr "diagonalUp") except ValueError: false


proc styleFont*(row: Row, col: string): Font =
  ## Get the style font from the cell in the row.
  result = retrieveStyleId(row, col, "fontId", "fonts", toFont)

proc styleFont*(sheet: Sheet, colrow: string): Font =
  ## Get the style font from the sheet to specified cell.
  let (c, r) = colrow.colrow
  sheet.row(r).styleFont(c)

proc styleFill*(row: Row, col: string): Fill =
  ## Get the style fill from the cell in the row.
  result = retrieveStyleId(row, col, "fillId", "fills", toFill)

proc styleFill*(sheet: Sheet, colrow: string): Fill =
  ## Get the style fill from the sheet to specified cell.
  let (c, r) = colrow.colrow
  sheet.row(r).styleFill c

proc styleBorder*(row: Row, col: string): Border =
  ## Get the style border from the cell in the row.
  result = retrieveStyleId(row, col, "borderId", "borders", toBorder)

proc styleBorder*(sheet: Sheet, colrow: string): Border =
  ## Get the style fill from the border to specified cell.
  let (c, r) = colrow.colrow
  sheet.row(r).styleBorder c
