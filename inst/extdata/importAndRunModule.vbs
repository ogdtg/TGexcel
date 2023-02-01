'Target Excel file to import BAS file to
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")


Filepath = WScript.Arguments(0)
pathToBASfile = WScript.Arguments(1)
moduleName = WScript.Arguments(2)

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(Filepath)

objExcel.Visible = True

'objExcel.DisplayAlerts = False

'Imports BAS module, but using a filepath

objExcel.VBE.ActiveVBProject.VBComponents.Import pathToBASfile

objExcel.DisplayAlerts = False

objExcel.Run moduleName


objWorkbook.Save

objWorkbook.Close 

