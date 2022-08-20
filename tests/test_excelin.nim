from std/times import now, DateTime, Time, toTime, parse, Month,
    month, year, monthday, toUnix, `$`
from std/strformat import fmt
from std/sugar import `->`, `=>`
from std/strscans import scanf
from std/os import fileExists, `/`, parentDir
from std/sequtils import repeat
from std/strutils import join

const v15Up = NimMajor >= 1 and NimMinor >= 5
when v15up:
  from std/math import isNaN, cbrt
else:
  from std/math import classify, FloatClass, cbrt
import std/[unittest, with, colors]

import excelin

suite "Excelin unit test":
  var
    excel, excelG: Excel
    sheet1, sheet1G: Sheet
    row1, row5, row1G: Row

  type ForExample = object
    a: string
    b: int

  proc `$`(f: ForExample): string = fmt"[{f.a}:{f.b}]"

  let
    nao = now()
    generatefname = "excelin_generated.xlsx"
    invalidexcel = currentSourcePath.parentDir() / "Book1-no-rels.xlsx"
    fexob = ForExample(a: "A", b: 200)
    row1cellA = "this is string"
    row1cellC = 256
    row1cellE = 42.42
    row1cellB = nao
    row1cellD = "2200/12/01"
    row1cellF = $fexob
    row1cellH = -111
    tobeShared = "the brown fox jumps over the lazy dog"
      .repeat(5).join(";")

  test "can create empty excel and sheet":
    (excel, sheet1) = newExcel()
    check excel != nil
    check sheet1 != nil

  test "can check sheet name and edit to new name":
    check sheet1.name == "Sheet1"
    sheet1.name = "excelin-example"
    check sheet1.name == "excelin-example"

  test "can fetch row 1 and 5":
    row1 = sheet1.row 1
    row5 = sheet1.row 5
    check row1 != nil
    check row5 != nil

  test "can check row number":
    check row1.rowNum == 1
    check row5.rowNum == 5

  test "can put values to row 1":
    row1["A"] = row1cellA
    row1["C"] = row1cellC
    row1["E"] = row1cellE
    row1["B"] = row1cellB
    row1["D"] = row1cellD
    row1["F"] = row1cellF
    row1["H"] = row1cellH

  test "can create shared strings for long size (more than 64 chars)":
    row1["J"] = tobeShared
    row1["K"] = tobeShared

  test "can fetch values inputted from row 1":
    check row1["A", string] == row1cellA
    check row1.getCell[:uint]("C") == row1cellC.uint
    check row1["H", int] == row1cellH
    check row1["B", DateTime].toTime.toUnix == row1cellB.toTime.toUnix
    check row1["B", Time].toUnix == row1cellB.toTime.toUnix
    check row1["E", float] == row1cellE
    checkpoint "fetching other values"
    check row1["J", string] == tobeShared
    check row1["K", string] == tobeShared
    checkpoint "fetching shared string done"

  test "can fetch with custom converter":
    let dt = row1.getCell[:DateTime]("D",
      (s: string) -> DateTime => parse(s, "yyyy/MM/dd"))
    check dt.year == 2200
    check dt.month == mDec
    check dt.monthday == 1

    let dt2 = row1.getCellIt[:DateTime]("D", parse(it, "yyyy/MM/dd"))
    check dt2.year == 2200
    check dt2.month == mDec
    check dt2.monthday == 1

    let fex = row1.getCell[:ForExample]("F", func(s: string): ForExample =
        discard scanf(s, "[$w:$i]", result.a, result.b)
    )
    check fex.a == fexob.a
    check fex.b == fexob.b

    let fex2 = row1.getCellIt[:ForExample]("F", (
      discard it.scanf("[$w:$i]", result.a, result.b)))
    check fex2.a == fexob.a
    check fex2.b == fexob.b

  test "can fill and fetch cell with formula format":
    let row3 = sheet1.row 3
    let row4 = sheet1.row(4, cfFilled)
    var sum3, sum4: int
    for i in 0 .. 9:
      let col = i.toCol
      row3[col] = i
      row4[col] = i
      sum3 += i
      sum4 += i
    row3[10.toCol] = Formula(equation: "SUM(A3:J3)", valueStr: $sum3)
    row4["K"] = Formula(equation: "SUM(A4:J4)", valueStr: $sum4)
    let f3 = row3["K", Formula]
    let f4 = row4[10.toCol, Formula]
    check f3.equation == "SUM(A3:J3)"
    check f4.equation == "SUM(A4:J4)"
    check f3.valueStr == f4.valueStr

    let cube3 = cbrt(float64 sum3)
    let cube4 = cbrt(float64 sum4)
    row3["L"] = Formula(equation: "CUBE(K3)", valueStr: $cube3)
    row4["L"] = Formula(equation: "CUBE(K4)", valueStr: $cube4)
    let fl3 = row3["L", Formula]
    let fl4 = row4["L", Formula]
    check fl3.equation == "CUBE(K3)"
    check fl4.equation == "CUBE(K4)"
    check fl3.valueStr == fl4.valueStr

  test "can give the result to string and write to file":
    excel.writeFile generatefname
    check fileExists generatefname
    check ($excel).len > 0

  test fmt"can read excel file from {generatefname} and assign sheet 1 and row 1":
    excelG = readExcel generatefname
    check excelG != nil
    let names = excelG.sheetNames
    check names.len == 1
    check names == @["excelin-example"]
    sheet1G = excelG.getSheet "excelin-example"
    check sheet1G != nil
    row1G = sheet1G.row 1
    check row1G != nil

  test "can add new sheets with default name or specified name":
    let sheet2 = excelG.addSheet
    require sheet2 != nil
    check excelG.sheetNames == @["excelin-example", "Sheet2"]
    check sheet2.name == "Sheet2"
    let sheet3modif = excelG.addSheet "new-sheet"
    require sheet3modif != nil
    check sheet3modif.name == "new-sheet"
    check excelG.sheetNames == @["excelin-example", "Sheet2", "new-sheet"]
    let s4 = excelG.addSheet
    require s4 != nil
    check s4.name == "Sheet4"
    check excelG.sheetNames == @["excelin-example", "Sheet2", "new-sheet", "Sheet4"]

  test "can add duplicate sheet name, delete the older, do nothing if no sheet to delete":
    let s5 = excelG.addSheet "new-sheet"
    require s5 != nil
    check s5.name == "new-sheet"
    check excelG.sheetNames == @["excelin-example", "Sheet2", "new-sheet", "Sheet4", "new-sheet"]
    excelG.deleteSheet "new-sheet"
    check excelG.sheetNames == @["excelin-example", "Sheet2", "Sheet4", "new-sheet"]
    excelG.deleteSheet "not-exists"
    check excelG.sheetNames == @["excelin-example", "Sheet2", "Sheet4", "new-sheet"]

  test "can refetch row from read file":
    check row1G["A", string] == row1cellA
    check row1G.getCell[:uint]("C") == row1cellC.uint
    check row1G["H", int] == row1cellH
    check row1G["B", DateTime].toTime.toUnix == row1cellB.toTime.toUnix
    check row1G["B", Time].toUnix == row1cellB.toTime.toUnix
    check row1G["E", float] == row1cellE
    check row1G["J", string] == tobeShared
    check row1G["K", string] == tobeShared

  test "can convert column string to int vice versa":
    let colnum = [("A", 0), ("AA", 26), ("AB", 27), ("ZZ", 701), ("AAA", 702),
      ("AAB", 703), ("ZZZ", 18277), ("AAAA", 18278)]
    for cn in colnum:
      check cn[0].toNum == cn[1]
      check cn[1].toCol == cn[0]

  test "can fetch arbitrary row number in new/empty sheet":
    let (_, sheet) = newExcel()
    let row2 = sheet.row 2
    check row2.rowNum == 2
    let row3 = sheet.row(3, cfFilled)
    check row3.rowNum == 3

  test "can clear out all cells in row":
    clear row1
    check row1["A", string] == ""
    check row1.getCell[:uint]("C") == uint.high
    check row1["H", int] == int.low
    check row1["B", DateTime].toTime.toUnix == 0
    check row1["B", Time].toUnix == 0
    when v15up:
      check row1["E", float].isNaN
    else:
      check row1["E", float].classify == fcNan
    checkpoint "fetching other values"
    check row1["J", string] == ""
    check row1["K", string] == ""
    checkpoint "fetching shared string done"

  test "can initialize border and its properties for style":
    var b = borderStyle(diagonalUp = true)
    with b:
      start = borderPropStyle(style = bsMedium) # border style
      diagonalDown = true

    check b.diagonalUp
    check b.diagonalDown
    check b.start.style == bsMedium

    let b2 = borderStyle(
      start = borderPropStyle(style = bsThick),
      `end` = borderPropStyle(style = bsThick),
      top = borderPropStyle(style = bsDotted),
      bottom = borderPropStyle(style = bsDotted))

    check not b2.diagonalUp
    check not b2.diagonalDown
    check b2.start.style == bsThick
    check b2.`end`.style == bsThick
    check b2.top.style == bsDotted
    check b2.bottom.style == bsDotted

  test "can refetch cell styling":
    let (excel, sheet) = newExcel()
    sheet.row(5).style("G",
      font = fontStyle(name = "Cambria Explosion", size = 1_000_000, color = $colGreen),
      border = borderStyle(
        top = borderPropStyle(style = bsThick, color = $colRed),
        `end` = borderPropStyle(style = bsDotted, color = $colNavy),
      ),
      fill = fillStyle(
        pattern = patternFillStyle(patternType = ptDarkDown),
        gradient = gradientFillStyle(
          stop = GradientStop(
            position: 0.75,
            color: $colGray,
          )
        ),
      )
    )
    const fteststyle = "test-style.xlsx"
    excel.writeFile fteststyle

    let excel2 = readExcel fteststyle

    let newsheet = excel2.getSheet "Sheet1"
    let font = newsheet.styleFont("G5")
    check font.name == "Cambria Explosion"
    check font.size == 1_000_000
    check font.color == $colGreen
    check font.family == -1
    check font.charset == -1
    check not font.bold
    check not font.strike
    check not font.italic
    check not font.condense
    check not font.extend
    check not font.outline
    check font.underline == uNone
    check font.verticalAlign == vaBaseline

    let fill = newsheet.styleFill("G5")
    check fill.pattern.patternType == ptDarkDown
    check fill.pattern.fgColor == $colWhite
    check fill.gradient.stop.color == $colGray
    check fill.gradient.stop.position == 0.75
    check fill.gradient.`type` == gtLinear
    check fill.gradient.degree == 0.0
    check fill.gradient.left == 0.0
    check fill.gradient.right == 0.0
    check fill.gradient.top == 0.0
    check fill.gradient.bottom == 0.0

    let border = newsheet.styleBorder "G5"
    check not border.diagonalDown
    check not border.diagonalUp
    check border.top.style == bsThick
    check border.top.color == $colRed
    check border.`end`.style == bsDotted
    check border.`end`.color == $colNavy
    check border.start.style == bsNone
    check border.start.color == ""
    check border.bottom.style == bsNone
    check border.bottom.color == ""

  test "can throw ExcelError when invalid excel without workbook relations found":
    expect ExcelError:
      discard readExcel(invalidexcel)

  test "can fetch last row in sheet":
    var unused: Excel
    (unused, sheet1) = newExcel()
    discard sheet1.row(1)        # add row empty
    sheet1.row(5)["D"] = "test"  # row 5 not empty and not hidden
    let row2 = sheet1.row 2      # not empty and hidden
    row2[100.toCol] = 0xb33f
    row2.hide = true
    sheet1.row(10).hide = true   # empty and hidden

    check sheet1.lastRow.rowNum == 5
    check sheet1.lastRow(getEmpty = true).rowNum == 5
    check sheet1.lastRow(getHidden = true).rowNum == 5
    check sheet1.lastRow(getEmpty = true, getHidden = true).rowNum == 10

  test "can check whether sheet empty and iterating the rows":
    var rowiter = rows
    var r = sheet1.rowiter
    check r.rowNum == 1
    check r.empty
    r = sheet1.rowiter
    check r.rowNum == 2
    check not r.empty
    check r.hidden
    r = sheet1.rowiter
    check r.rowNum == 5
    check not r.empty
    check not r.hidden
    r = sheet1.rowiter
    check r.rowNum == 10
    check r.empty
    check r.hidden
    discard sheet1.rowiter  # because iterator will only be true finished
                            # one more iteration after it's emptied.
    check rowiter.finished

  test "can get last cell in row":
    var rowiter = rows
    var r = sheet1.rowiter
    check r.rowNum == 1
    check r.lastCol == ""
    r = sheet1.rowiter
    check r.rowNum == 2
    check r.lastCol == 100.toCol
    r = sheet1.rowiter
    check r.rowNum == 5
    check r.lastCol == "D"
    r = sheet1.rowiter
    check r.rowNum == 10
    check r.lastCol == ""
    discard sheet1.rowiter
