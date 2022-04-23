# Excelin - create and read Excel pure Nim

A library to work with Excel file and/or data.

# Examples

All operations available are illustrated in below:

```nim
from std/times import now, DateTime, Time, toTime, parse, Month,
    month, year, monthday, toUnix, `$`
from std/strformat import fmt
from std/sugar import `->`, `=>`, dump
from std/strscans import scanf
import excelin

# `newExcel` returns Excel and Sheet object to immediately work
# when creating an Excel data.
let (excel, sheet) = newExcel()

# we of course can also read from Excel file directly using `readExcel`
# we comment this out because the path is imaginary
#let (excelTemplate, sheetFromTemplate) = readExcel("path/to/template.xlsx")

doAssert sheet.name == "Sheet1"
# by default the name sheet is Sheet1

# let's change it to other name
sheet.name = "excelin-example"
doAssert sheet.name == "excelin-example"

# let's add/fetch some row to our sheet
let row1 = sheet.row 1

# excelin.row is immediately creating when the rows if it's not available
# and if it's available, it's returning the existing.
# With excelin.rowNum, we can check its row number.
doAssert row1.rowNum == 1

# let's add another row, this time it's row 5
let row5 = sheet.row 5
doAssert row5.rowNum == 5
# in this case, we immediately get the row 5 even though the existing
# rows in the sheet are only one.
# PS: addRow proc(Sheet): Row and addRow proc(Sheet, Positive): Row is deprecated.

type
    ForExample = object
        a: string
        b: int

proc `$`(f: ForExample): string = fmt"[{f.a}:{f.b}]"

# let's put some values in row cells
let nao = now()
row1["A"] = "this is string"
row1["C"] = 256
row1["E"] = 42.42
row1["B"] = nao # Excelin support DateTime or Time and
                # by default it will be formatted as yyyy-MM-dd'T'HH:mm:dd.fff'.'zz
                # e.g.: 2200-12-01T22:10:23.456+01
row1["D"] = "2200/12/01" # we put the date as string for later example when fetching
                         # using supplied converter function from cell value
row1["F"] = $ForExample(a: "A", b: 2)
row1["H"] = -111

# notice above example we arbitrarily chose the column and by current implementation
# Excel data won't add unnecessary empty cells. In other words, sparse row cells.
# At later, we will add the implementation to fill all cells because it will be
# more efficient when working with large columns.
 
# now let's fetch the data we inputted
doAssert row1["A", string] == "this is string"
doAssert row1.getCell[:uint]("C") == 256
doAssert row1["H", int] == -111
doAssert row1["B", DateTime].toTime.toUnix == nao.toTime.toUnix
doAssert row1["B", Time].toUnix == nao.toTime.toUnix
doAssert row1["E", float] == 42.42
# in above example, we fetched various values from its designated cell position
# using the two kind of function, `getCell` and `[]`. `[]` often used for
# elementary/primitive types those supported by Excelin by default. `getCell`
# has 3rd parameter, a closure with signature `string -> R`, which default to `nil`,
# that will give users the flexibility to read the string value representation
# to the what intended to convert. We'll see it below
#
# note also that we need to compare to its second for DateTime|Time instead of directly using
# times.`==` because the comparison precision up to nanosecond, something we
# can't provide in this example

let dt = row1.getCell[:DateTime]("D",
  (s: string) -> DateTime => (
    dump s; result = parse(s, "yyyy/MM/dd"); dump result))
doAssert dt.year == 2200
doAssert dt.month == mDec
doAssert dt.monthday == 1

let fex = row1.getCell[:ForExample]("F", func(s: string): ForExample =
    discard scanf(s, "[$w:$i]", result.a, result.b)
)
doAssert fex.a == "A"
doAssert fex.b == 2
# above examples we provide two example of using closure for converting
# string representation of cell value to our intended object. With this,
# users can roll their own conversion way to interpret the cell data.

# finally, we have 2 options to access the binary Excel data, using `$` and
# `writeFile`. Both of procs are the usual which `$` is stringify (that's
# to return the string of Excel) and `writeFile` is accepting string path
# to where the Excel data will be written.
let toSendToWire = $excel
excel.writeFile("to/any/path/we/choose.xlsx")

# note that the current excelin.`$` is using the `writeFile` first to temporarily
# write to file in $TEMP dir because the current zip lib dependency doesn't
# provide the `$` to get the raw data from built zip directly.
```

Another example here we work directly with `Sheet` instead of the `Rows` and/or cells.

```nim
import excelin

# prepare our excel
let (excel, _) = newExcel()
doAssert excel.sheetNames == @["Sheet1"]

# above we see that our excel has seq string with a member
# "Sheet1". The "Sheet1" is the default sheet when creating
# a new Excel file.
# Let's add a sheet to our Excel.

let newsheet = excel.addSheet "new-sheet"
doAssert newsheet.name == "new-sheet"
doAssert excel.sheetNames == @["Sheet1", "new-sheet"]

# above, we add a new sheet with supplied of the new-sheet name.
# By checking with `sheetNames` proc, we see two sheets' name.

# Let's see what happen when we add a sheet without supplying the name
let sheet3 = excel.addSheet
doAssert sheet3.name == "Sheet3"
doAssert excel.sheetNames == @["Sheet1", "new-sheet", "Sheet3"]

# While the default name quite unexpected, we can guess the "num" part
# for default sheet naming is related to how many we added/invoked
# the `addSheet` proc. We'll see below example why it's done like this.

# Let's add again
let anewsheet = excel.addSheet "new-sheet"
doAssert anewsheet.name == "new-sheet"
doAssert excel.sheetNames == @["Sheet1", "new-sheet", "Sheet3", "new-sheet"]

# Here, we added a new sheet using existing sheet name.
# This can be done because internally Excel workbook holds the reference of
# sheets is by using its id instead of the name. Hence adding a same name
# for new sheet is possible.
# For the consquence, let's see what happens when we delete a sheet below

# work fine case
excel.deleteSheet "Sheet1"
doAssert excel.sheetNames == @["new-sheet", "Sheet3", "new-sheet"]

# deleting sheet with name "new-sheet"
excel.deleteSheet "new-sheet"
doAssert excel.sheetNames == @["Sheet3", "new-sheet"]
# will delete the older one since it's the first the sheet found with "new-sheet" name

# when there's name available, Excel file will do nothing.
excel.deleteSheet "aww-sheet"
doAssert excel.sheetNames == @["Sheet3", "new-sheet"]
# still same as before.

# Below example we illustrate how to get by sheet name.
anewsheet.row(1)["A"] = "temptest"
doAssert anewsheet.row(1)["A", string] == "temptest"
discard excel.addSheet "new-sheet" # add a new to make it duplicate
let foundOlderSheet = excel.getSheet "new-sheet"
doAssert foundOlderSheet.row(1)["A", string] == "temptest"
# Here we get sheet by name, and like deleting the sheet, fetching/getting
# the sheet also returning the older sheet of the same name.

doAssert excel.sheetNames == @["Sheet3", "new-sheet", "new-sheet"]
excel.writeFile ("many-sheets.xlsx")
# Write it to file and open it with our favorite Excel viewer to see 2 sheets:
# Sheet3 and new-sheet, new-sheet.
# Using libreoffice to view the Excel file, the duplicate name will be appended with
# format {sheetName}-{numDuplicated}.
# We can replicate that behaviour too but currently we support duplicate sheet name.
```

# Install

Excelin requires minimum Nim version of `v1.4.0`.  

For installation, we can choose several methods will be mentioned below.

Using Nimble package (when it's available):

```
nimble install excelin
```

Or to install it locally

```
git clone https://github.com/mashingan/excelin
cd excelin
nimble develop
```

or directly from Github repo

```
nimble install https://github.com/mashingan/excelin 
```

to install the `#head` branch

```
nimble install https://github.com/mashingan/excelin@#head
#or
nimble install excelin@#head
```

# License

MIT
