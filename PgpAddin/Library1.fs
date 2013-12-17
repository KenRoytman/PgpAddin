module PgpAddin

open ExcelDna.Integration
open Microsoft.Office.Interop.Excel
open System

type Skills = {Id : int ; Name : string}

type ExcelInterface() = 
  member val Handle = ExcelDnaUtil.Application :?> Application with get

  interface IExcelAddIn with
    member x.AutoOpen() = ExcelAsyncUtil.Initialize()
    member x.AutoClose() = ExcelAsyncUtil.Uninitialize()

let RefDataSheet = "Reference Data"
let SkillsTable = "Skills"

let xl = ExcelInterface()

let xlExec action = ExcelAsyncUtil.QueueAsMacro action

let loadXlLst name =
    let sheet = xl.Handle.Sheets.Item(RefDataSheet) :?> Worksheet
    let lst = sheet.ListObjects.Item(SkillsTable)
    let enumerator = lst.ListRows.GetEnumerator()
    seq { while enumerator.MoveNext()
            do yield enumerator.Current :?> ListRow }
    
let getSkills =
    let rows = loadXlLst SkillsTable
    rows |> Seq.map ( fun x -> let rng = x.Range.Value2 :?> obj[,]
                               let id = rng.[1, 1]
                               let name = rng.[1,2]
                               { Id = Convert.ToInt32(id); Name = Convert.ToString(name) } )
    
[<ExcelFunction>]
let poop (x : int) = getSkills |> Seq.length

[<ExcelFunction>]
let poop1 x y = x + y