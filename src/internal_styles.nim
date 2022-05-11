include internal_rows

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

proc addFont(styles: XmlNode, font: Font): (int, bool) =
  if font.name == "": return
  let fontnode = font.toXmlNode
  let applyFont = true

  let fonts = styles.retrieveChildOrNew "fonts"
  let fontCount = try: parseInt(fonts.attr "count") except: 0
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

# Had to add for API style consistency.
proc fontStyle*(name: string, size = 10,
  family, charset = -1,
  bold, italic, strike, outline, shadow, condense, extend = false,
  color = "", underline = uNone, verticalAlign = vaBaseline): Font =
  Font(
    name: name,
    size: size,
    family: family,
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
  let sparse = $cfSparse == row.body.attr "cellfill"
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