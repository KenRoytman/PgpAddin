module PgpAddin

open ExcelDna.Integration
open Microsoft.Office.Interop.Excel

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
    lst.ListRows

//let getSkills =
//    let rows = loadXlLst SkillsTable


//[<ExcelFunction()>]
//let poop x y = ExcelAsyncUtil.QueueAsMacro 