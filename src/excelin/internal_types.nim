# Excelin
# Library to read and create Excel file purely in Nim
# MIT License Copyright (c) 2022 Rahmatullah

## Excelin
## *******
##
## A library to work with spreadsheet file (strictly .xlsx) without dependency
## outside of Nim compiler (and its requirement) and its development environment.
##

from std/xmltree import XmlNode, findAll, `$`, child, items, attr, `<>`,
     newXmlTree, add, newText, toXmlAttributes, delete, len, xmlHeader,
     attrs, `attrs=`, innerText, `[]`, insert, clear, XmlNodeKind, kind,
     tag
from std/tables import TableRef, newTable, `[]`, `[]=`, contains, pairs,
     keys, del, values, initTable, len


const
  excelinVersion* = "0.4.9"

type
  Excel* = ref object
    ## The object the represent as Excel file, mostly used for reading the Excel
    ## and writing it to Zip file and to memory buffer (at later version).
    content: XmlNode
    rels: XmlNode
    workbook: Workbook
    sheets: TableRef[FilePath, Sheet]
    sharedStrings: SharedStrings
    otherfiles: TableRef[string, FileRep]
    embedfiles: TableRef[string, EmbedFile]
    sheetCount: int

  InternalBody = object of RootObj
    body: XmlNode

  Workbook* = ref object of InternalBody
    ## The object that used for managing package information of Excel file.
    ## Most users won't need to use this.
    path: string
    sheetsInfo: seq[XmlNode]
    rels: FileRep
    parent: Excel

  Sheet* = ref object of InternalBody
    ## The main object that will be used most of the time for many users.
    ## This object will represent a sheet in Excel file such as adding row
    ## getting row, and/or adding cell directly.
    parent: Excel
    rid: string
    privName: string
    filename: string

  Row* = ref object of InternalBody
    ## The object that will be used for working with values within cells of a row.
    ## Users can get the value within cell and set its value.
    sheet: Sheet
  FilePath = string
  FileRep = (FilePath, XmlNode)
  EmbedFile = (FilePath, string)

  SharedStrings = ref object of InternalBody
    path: string
    strtables: TableRef[string, int]
    count: Natural
    unique: Natural

  ExcelError* = object of CatchableError
    ## Error when the Excel file read is invalid, specifically Excel file
    ## that doesn't have workbook.

  CellFill* = enum
    cfSparse = "sparse"
    cfFilled = "filled"

  Formula* = object
    ## Object exclusively working with formula in a cell.
    ## The equation is simply formula representative and
    ## the valueStr is the value in its string format,
    ## which already calculated beforehand.
    equation*: string
    valueStr*: string

  Hyperlink* = object
    ## Object that will be used to fill cell with external link.
    target*: string
    text*: string
    tooltip*: string

  Font* = object
    ## Cell font styling. Provide name if it's intended to style the cell.
    ## If no name is supplied, it will ignored. Field `family` and `charset`
    ## are optionals but in order to be optional, provide it with negative value
    ## because there's value for family and charset 0. Since by default int is 0,
    ## it could be yield different style if the family and charset are not intended
    ## to be filled but not assigned with negative value.
    name*: string
    family*: int
    charset*: int
    size*: Positive
    bold*: bool
    italic*: bool
    strike*: bool
    outline*: bool
    shadow*: bool
    condense*: bool
    extend*: bool
    color*: string
    underline*: Underline
    verticalAlign*: VerticalAlign

  Underline* = enum
    uNone = "none"
    uSingle = "single"
    uDouble = "double"
    uSingleAccounting = "singleAccounting"
    uDoubleAccounting = "doubleAccounting"

  VerticalAlign* = enum
    vaBaseline = "baseline"
    vaSuperscript = "superscript"
    vaSubscript = "subscript"

  Border* = object
    ## The object that will define the border we want to apply to cell.
    ## Use `border <#border,BorderProp,BorderProp,BorderProp,BorderProp,BorderProp,BorderProp,bool,bool>`_
    ## to initialize working border instead because the indication whether border can be edited is private.
    edit: bool
    start*: BorderProp # left
    `end`*: BorderProp # right
    top*: BorderProp
    bottom*: BorderProp
    vertical*: BorderProp
    horizontal*: BorderProp
    diagonalUp*: bool
    diagonalDown*: bool

  BorderProp* = object
    ## The object that will define the style and color we want to apply to border
    ## Use `borderProp<#borderProp,BorderStyle,string>`_
    ## to initialize working border prop instead because the indication whether
    ## border properties filled is private.
    edit: bool ## indicate whether border properties is filled
    style*: BorderStyle
    color*: string #in RGB

  BorderStyle* = enum
    bsNone = "none"
    bsThin = "thin"
    bsMedium = "medium"
    bsDashed = "dashed"
    bsDotted = "dotted"
    bsThick = "thick"
    bsDouble = "double"
    bsHair = "hair"
    bsMediumDashed = "mediumDashed"
    bsDashDot = "dashDot"
    bsMediumDashDot = "mediumDashDot"
    bsDashDotDot = "dashDotDot"
    bsMediumDashDotDot = "mediumDashDotDot"
    bsSlantDashDot = "slantDashDot"

  Fill* = object
    ## Fill cell style. Use `fillStyle <#fillStyle,PatternFill,GradientFill>`_
    ## to initialize this object to indicate cell will be edited with this Fill.
    edit: bool
    pattern*: PatternFill
    gradient*: GradientFill

  PatternFill* = object
    ## Pattern to fill the cell. Use `patternFill<#patternFill,string,string,PatternType>`_
    ## to initialize.
    edit: bool
    fgColor*: string
    bgColor: string
    patternType*: PatternType

  PatternType* = enum
    ptNone = "none"
    ptSolid = "solid"
    ptMediumGray = "mediumGray"
    ptDarkGray = "darkGray"
    ptLightGray = "lightGray"
    ptDarkHorizontal = "darkHorizontal"
    ptDarkVertical = "darkVertical"
    ptDarkDown = "darkDown"
    ptDarkUp = "darkUp"
    ptDarkGrid = "darkGrid"
    ptDarkTrellis = "darkTrellis"
    ptLightHorizontal = "lightHorizontal"
    ptLightVertical = "lightVertical"
    ptLightDown = "lightDown"
    ptLightUp = "lightUp"
    ptLightGrid = "lightGrid"
    ptLightTrellis = "lightTrellis"
    ptGray125 = "gray125"
    ptGray0625 = "gray0625"

  GradientFill* = object
    ## Gradient to fill the cell. Use
    ## `gradientFill<#gradientFill,GradientStop,GradientType,float,float,float,float,float>`_
    ## to initialize.
    edit: bool
    stop*: GradientStop
    `type`*: GradientType
    degree*: float
    left*: float
    right*: float
    top*: float
    bottom*: float

  GradientStop* = object
    ## Indicate where the gradient will stop with its color at stopping position.
    color*: string
    position*: float

  GradientType* = enum
    gtLinear = "linear"
    gtPath = "path"

  Range* = (string, string)
    ## Range of table which consist of top left cell and bottom right cell.

  FilterType* = enum
    ftFilter
    ftCustom

  Filter* = object
    ## Filtering that supplied to column id in sheet range. Ignored if the sheet
    ## hasn't set its auto filter range.
    case kind*: FilterType
    of ftFilter:
      valuesStr*: seq[string]
    of ftCustom:
      logic*: CustomFilterLogic
      customs*: seq[(FilterOperator, string)]

  FilterOperator* = enum
    foEq = "equal"
    foLt = "lessThan"
    foLte = "lessThanOrEqual"
    foNeq = "notEqual"
    foGte = "greaterThanOrEqual"
    foGt = "greaterThan"

  CustomFilterLogic* = enum
    cflAnd = "and"
    cflOr = "or"
    cflXor = "xor"