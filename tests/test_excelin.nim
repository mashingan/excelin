from std/times import now, DateTime, Time, toTime, parse, Month,
    month, year, monthday, toUnix, `$`
from std/strformat import fmt
from std/sugar import `->`, `=>`
from std/strscans import scanf
from std/os import fileExists
from std/sequtils import repeat
from std/strutils import join
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
