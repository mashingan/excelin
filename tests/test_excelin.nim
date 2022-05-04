from std/times import now, DateTime, Time, toTime, parse, Month,
    month, year, monthday, toUnix, `$`
from std/strformat import fmt
from std/sugar import `->`, `=>`
from std/strscans import scanf
from std/os import fileExists
from std/sequtils import repeat
from std/strutils import join
from std/math import cbrt

const v15Up = NimMajor >= 1 and NimMinor >= 5
when v15up:
  from std/math import isNaN
else:
  from std/math import classify, FloatClass
import std/unittest

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
