module ExcelProvider.Model

open System
open System.Collections.Generic
open System.Reflection

open ExcelProvider.ExcelAddressing
open ICSharpCode.SharpZipLib.Zip

// Represents a row in a provided ExcelFileInternal
type Row(rowIndex, getCellValue: int -> int -> obj, columns: Map<string, int>) = 
    member this.GetValue columnIndex = getCellValue rowIndex columnIndex

    override this.ToString() =
        let columnValueList =                    
            [for column in columns do
                let value = getCellValue rowIndex column.Value
                let columnName, value = column.Key, string value
                yield sprintf "\t%s = %s" columnName value]
            |> String.concat Environment.NewLine

        sprintf "Row %d%s%s" rowIndex Environment.NewLine columnValueList

// get the type, and implementation of a getter property based on a template value
let private propertyImplementation columnIndex (value : obj) =
    match value with
    | :? float -> typeof<double>, (fun [row] -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> v :?> double) @@>)
    | :? bool -> typeof<bool>, (fun [row] -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> v :?> bool) @@>)
    | :? DateTime -> typeof<DateTime>, (fun [row] -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> v :?> DateTime) @@>)
    | :? string -> typeof<string>, (fun [row] -> <@@ (%%row: Row).GetValue columnIndex |> (fun v -> v :?> string) @@>)
    | _ -> typeof<obj>, (fun [row] -> <@@ (%%row: Row).GetValue columnIndex @@>)

// gets a list of column definition information for the columns in a view         
let internal getColumnDefinitions (data : View) forcestring =
    let getCell = getCellValue data
    [for columnIndex in 0 .. data.ColumnMappings.Count - 1 do
        let columnName = getCell 0 columnIndex |> string
        if not (String.IsNullOrWhiteSpace(columnName)) then
            let cellType, getter =
                if forcestring then
                    let getter = (fun [row] -> 
                    <@@ 
                        let value = (%%row: Row).GetValue columnIndex |> string
                        if String.IsNullOrEmpty value then null
                        else value
                    @@>)
                    typedefof<string>, getter
                else
                    let cellValue = getCell 1 columnIndex
                    propertyImplementation columnIndex cellValue             
            yield (columnName, (columnIndex, cellType, getter))]

// Simple type wrapping Excel data
type ExcelFileInternal(filename, range) =
    
    let data = 
        let view = openWorkbookView filename range
        let columns = [for (columnName, (columnIndex, _, _)) in getColumnDefinitions view true -> columnName, columnIndex] |> Map.ofList
        let buildRow rowIndex = new Row(rowIndex, getCellValue view, columns)        
        seq{ 1 .. view.RowCount}
        |> Seq.map buildRow

    member __.Data = data

// Set up assembly reference forwarding for the ICSharpCode.SharpZipLib assembly.
// This would normally be done in an app.config file, but this is not convenient, as this assembly will be dynmically generated.
do 
    let sharpZipLibAssemblyName = 
        let zipFileType = typedefof<ZipFile>
        zipFileType.Assembly.GetName()

    // Keep track of assemblies that have already been loaded to prevent stack overflows
    let loadedAssemblies = new HashSet<string>()
   
    let resolveAssembly sender (resolveEventArgs : ResolveEventArgs) =        
        let assemblyName = resolveEventArgs.Name
        if loadedAssemblies.Add( assemblyName ) then         
            let assemblyName =
                if assemblyName.StartsWith(sharpZipLibAssemblyName.Name)
                then sharpZipLibAssemblyName.FullName
                else assemblyName
            Assembly.Load( assemblyName )
        else null

    let handler = new ResolveEventHandler( resolveAssembly )
    AppDomain.CurrentDomain.add_AssemblyResolve handler