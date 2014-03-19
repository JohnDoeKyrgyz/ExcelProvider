namespace SpreasheetTypes
open FSharp.ExcelProvider
type BookTest = ExcelFile<"BookTest.xls", "Sheet1", true>
type HeaderTest = ExcelFile<"BookTestWithHeader.xls", "A2", true>
type MultipleRegions = ExcelFile<"MultipleRegions.xlsx", "A1:C5,E3:G5", true>
type DifferentMainSheet = ExcelFile<"DifferentMainSheet.xlsx">