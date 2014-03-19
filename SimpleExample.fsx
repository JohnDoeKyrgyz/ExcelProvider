printfn __SOURCE_DIRECTORY__

#I __SOURCE_DIRECTORY__
#r @"bin\ExcelProvider.dll"

open FSharp.ExcelProvider

type DataTypesTest = ExcelFile<"tests\ExcelProvider.Tests\DataTypes.xlsx">
let file = new DataTypesTest()
let row = file.Data |> Seq.head

printfn "\nTypes:\n"
printfn "%A" (row.String)
printfn "%A" (row.Float)
printfn "%A" (row.Boolean)
printfn "ToString() %O" row