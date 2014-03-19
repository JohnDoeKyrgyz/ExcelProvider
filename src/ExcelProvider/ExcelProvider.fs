module ExcelProvider.ExcelProvider

open System
open System.Collections.Generic
open System.IO
open System.Reflection

open ExcelProvider.ExcelAddressing
open ExcelProvider.Helper
open ExcelProvider.Model
open Microsoft.FSharp.Core.CompilerServices
open Samples.FSharp.ProvidedTypes

type internal GlobalSingleton private () =
    static let mutable instance = Dictionary<_, _>()
    static member Instance = instance

let internal memoize f x =
    if (GlobalSingleton.Instance).ContainsKey(x) then (GlobalSingleton.Instance).[x]
    else 
        let res = f x
        (GlobalSingleton.Instance).[x] <- res
        res

let internal typExcel (cfg:TypeProviderConfig) =

    // Create an assembly to host the generated types
    let executingAssembly = Assembly.GetExecutingAssembly()

    // Create the main provided type
    let excelFileProvidedType = ProvidedTypeDefinition(executingAssembly, rootNamespace, "ExcelFile", Some(typeof<ExcelFileInternal>))
    
    // Parameterize the type by the file to use as a template
    let filename = ProvidedStaticParameter("filename", typeof<string>)
    let range = ProvidedStaticParameter("sheetname", typeof<string>, "")
    let forcestring = ProvidedStaticParameter("forcestring", typeof<bool>, false)
    let staticParams = [ filename; range; forcestring ]

    do excelFileProvidedType.DefineStaticParameters(staticParams, fun tyName paramValues ->
        let (filename, range, forcestring) = 
            match paramValues with
            | [| :? string  as filename;   :? string as range;  :? bool as forcestring|] -> (filename, range, forcestring)
            | [| :? string  as filename;   :? bool as forcestring |] -> (filename, String.Empty, forcestring)
            | [| :? string  as filename|] -> (filename, String.Empty, false)
            | _ -> ("no file specified to type provider", String.Empty,  true)

        // resolve the filename relative to the resolution folder
        let resolvedFilename = Path.Combine(cfg.ResolutionFolder, filename)

        let ProvidedTypeDefinitionExcelCall (filename, range, forcestring)  =         
            let data = openWorkbookView resolvedFilename range

            // define a provided type for each row, erasing to a int -> obj
            let providedRowType = ProvidedTypeDefinition("Row", Some(typeof<Row>))

            // add one property per Excel field
            let columnProperties = getColumnDefinitions data forcestring
            for (columnName, (columnIndex, propertyType, getter)) in columnProperties do            

                let prop = ProvidedProperty(columnName, propertyType, GetterCode = getter)
                // Add metadata defining the property's location in the referenced file
                prop.AddDefinitionLocation(1, columnIndex, filename)
                providedRowType.AddMember(prop)

            // define the provided type, erasing to an seq<int -> obj>
            let providedExcelFileType = ProvidedTypeDefinition(executingAssembly, rootNamespace, tyName, Some(typeof<ExcelFileInternal>))

            // add a parameterless constructor which loads the file that was used to define the schema
            providedExcelFileType.AddMember(ProvidedConstructor([], InvokeCode = fun [] -> <@@ ExcelFileInternal(resolvedFilename, range) @@>))

            // add a constructor taking the filename to load
            providedExcelFileType.AddMember(ProvidedConstructor([ProvidedParameter("filename", typeof<string>)], InvokeCode = fun [filename] -> <@@ ExcelFileInternal(%%filename, range) @@>))

            // add a new, more strongly typed Data property (which uses the existing property at runtime)
            providedExcelFileType.AddMember(ProvidedProperty("Data", typedefof<seq<_>>.MakeGenericType(providedRowType), GetterCode =fun [excFile] -> <@@ (%%excFile:ExcelFileInternal).Data @@>))

            // add the row type as a nested type
            providedExcelFileType.AddMember(providedRowType)

            // add the provided types to the generation assembly            
            let providedAssemblyFilePath = Path.Combine( Path.GetTempPath(), "ExcelProvider.ProvidedTypes-" + Guid.NewGuid().ToString() + ".dll" )

            let providedAssembly = new ProvidedAssembly(providedAssemblyFilePath)
            let providedTypes = [providedExcelFileType; providedRowType]
            for providedType in providedTypes do
                providedType.IsErased <- false
                providedType.SuppressRelocation <- false
            providedAssembly.AddTypes(providedTypes)

            providedExcelFileType

        (memoize ProvidedTypeDefinitionExcelCall)(filename, range, forcestring))

    // add the type to the namespace
    excelFileProvidedType

[<TypeProvider>]
type public ExcelProvider(cfg:TypeProviderConfig) as this =
    inherit TypeProviderForNamespaces()

    do 
        let providedType = typExcel cfg        
        this.AddNamespace(rootNamespace,[providedType])

[<TypeProviderAssembly>]
do ()