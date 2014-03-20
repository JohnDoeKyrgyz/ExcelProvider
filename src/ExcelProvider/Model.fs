module ExcelProvider.Model

open System
open System.Collections.Generic
open System.Reflection

open ExcelProvider.ExcelAddressing
open ICSharpCode.SharpZipLib.Zip

// Represents a row in a provided ExcelFileInternal
type RowInternal(rowIndex, getCellValue: int -> int -> obj, columns: Map<string, int>) = 
    member this.GetValue columnIndex = getCellValue rowIndex columnIndex

    member private this.CoerceToString value =        
        let value = value |> string
        if String.IsNullOrWhiteSpace(value) then null
        else value

    override this.ToString() =
        let columnValueList =                    
            [for column in columns do
                let value = getCellValue rowIndex column.Value
                let columnName, value = column.Key, string value
                yield sprintf "\t%s = %s" columnName value]
            |> String.concat Environment.NewLine

        sprintf "Row %d%s%s" rowIndex Environment.NewLine columnValueList

let private getColumnTypeFromValue (value : obj) =
    match value with
    | :? float -> typeof<double>
    | :? bool -> typeof<bool>
    | :? DateTime -> typeof<DateTime>
    | :? string -> typeof<string>
    | _ -> typeof<obj>      

// gets a list of column definition information for the columns in a view         
let internal getColumnDefinitions (data : View) forcestring =
    let getCell = getCellValue data
    [for columnIndex in 0 .. data.ColumnMappings.Count - 1 do
        let columnName = getCell 0 columnIndex |> string
        if not (String.IsNullOrWhiteSpace(columnName)) then
            let cellType =
                if forcestring then typedefof<string>
                else
                    let cellValue = getCell 1 columnIndex
                    getColumnTypeFromValue cellValue
            yield (columnName, (columnIndex, cellType))]

// Simple type wrapping Excel data
type ExcelFileInternal(filename, range) =
    
    let data = 
        let view = openWorkbookView filename range
        let columns = [for (columnName, (columnIndex, _)) in getColumnDefinitions view true -> columnName, columnIndex] |> Map.ofList
        let buildRow rowIndex = new RowInternal(rowIndex, getCellValue view, columns)        
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