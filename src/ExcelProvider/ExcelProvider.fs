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
    
// creates an expression that can serve as the body for a property getter
let internal buildPropertyGetter (rowType : Type) propertyType (columnIndex : int) parameters =
    let getValueMethod = rowType.GetMethod("GetValue")
    let this = parameters |> List.head
    Quotations.Expr.Coerce( Quotations.Expr.Call( this, getValueMethod, [ <@@ columnIndex @@> ] ), propertyType )

// creates an expression that can serve as the body for a property getter that forces the output to be a string
let internal buildForceStringPropertyGetter (rowType : Type) (columnIndex : int) parameters =
    let standardGetter = buildPropertyGetter rowType typedefof<obj> columnIndex parameters
    let coerceMethod = rowType.GetMethod("CoerceToString")
    let this = parameters |> List.head
    Quotations.Expr.Call(this, coerceMethod, [standardGetter])
    
let dummyGetter (rowType : Type) [parameters] = <@@ null @@>

let internal typExcel (cfg:TypeProviderConfig) =

    // Create an assembly to host the generated types
    let executingAssembly = Assembly.GetExecutingAssembly()

    // Create the main provided type
    let excelFileProvidedType = ProvidedTypeDefinition(executingAssembly, rootNamespace, "ExcelFile", Some(typeof<obj>))
    
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

            //create the row type
            let providedRowType = ProvidedTypeDefinition("Row", Some(typeof<RowInternal>))

            //create the file type
            let genericFileBaseType = typedefof<ExcelFileInternal<_>>
            let fileBaseType = genericFileBaseType.MakeGenericType(providedRowType)
            let providedExcelFileType = ProvidedTypeDefinition(executingAssembly, rootNamespace, tyName, Some(fileBaseType))

            //define the row type constructor
            let baseConstructor = typedefof<RowInternal>.GetConstructors() |> Seq.head
            let rowConstructor = 
                ProvidedConstructor(
                    [ProvidedParameter("excelFile", providedExcelFileType); ProvidedParameter("rowIndex", typeof<int>)], 
                    BaseConstructorCall = (fun [excelFile; rowIndex] -> baseConstructor, [ <@@ RowInternal( %%excelFile, %%rowIndex ) @@>]))
            providedRowType.AddMember(rowConstructor)

            //define the file type constructors
            let rowBuilder = <@@ (fun file rowIndex -> Quotations.Expr.NewObject(rowConstructor, [file; rowIndex])) @@>            

            // add a parameterless constructor which loads the file that was used to define the schema
            let excelFileTypeConstructor = genericFileBaseType.GetConstructors() |> Seq.head
            let parameterlessConstructor = 
                ProvidedConstructor([],
                    BaseConstructorCall = fun [] -> excelFileTypeConstructor, [ <@@ ExcelFileInternal(resolvedFilename, range, %%rowBuilder) @@> ])
            providedExcelFileType.AddMember( parameterlessConstructor )

            // add a constructor taking the filename to load
            let filenameParameter = ProvidedParameter("filename", typeof<string>)
            let filenameConstructor = 
                ProvidedConstructor([filenameParameter], 
                    BaseConstructorCall = (fun [filename] -> baseConstructor, [ <@@ ExcelFileInternal(%%filename, range, %%rowBuilder) @@> ]) )
            providedExcelFileType.AddMember( filenameConstructor )

            // add one property per Excel field
            let columnProperties = getColumnDefinitions data forcestring
            for (columnName, (columnIndex, propertyType)) in columnProperties do

                let getter = dummyGetter providedRowType
//                    if forcestring then buildForceStringPropertyGetter providedRowType columnIndex
//                    else buildPropertyGetter providedRowType propertyType columnIndex

                let prop = ProvidedProperty(columnName, propertyType, GetterCode = getter)
                // Add metadata defining the property's location in the referenced file
                prop.AddDefinitionLocation(1, columnIndex, filename)
                providedRowType.AddMember prop

            // add the row type as a nested type
            providedExcelFileType.AddMember(providedRowType)
            
            // add the provided types to the generation assembly            
            let providedAssemblyFilePath = Path.Combine( Path.GetTempPath(), "ExcelProvider.ProvidedTypes-" + Guid.NewGuid().ToString() + ".dll" )

            let providedAssembly = new ProvidedAssembly(providedAssemblyFilePath)
            let providedTypes = [providedRowType; excelFileProvidedType]
            for providedType in providedTypes do
                providedType.IsErased <- false
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