' Abrir libro
Set objExcel = CreateObject("Excel.Application") ' crea instancia
Set objWorkBook = objExcel.WorkBooks.Open(pathFile) ' se abre el archivo
Set objSheet = objWorkBook.WorkSheets(nameSheet) ' se le asigna hoja a un objeto

' Cerrar archivo
objWorkBook.Save
objWorkBook.Close
objExcel.Quit

Set objWorkBook = Nothing
Set objSheet = Nothing
Set objExcel = Nothing

' Cells y Range
objSheet.Cells(Row, Col)
objSheet.Range("A1:A1")

' Auto Filter
objSheet.Range("A1:Z1").AutoFilter numColumna, datoBuscar ' filtra con un solo dato
objSheet.Range("A1:Z1").AutoFilter numColumna, datoBuscar_1, 2, datoBuscar2 ' filtra dos datos

' Delay
Wscript.Sleep(2000) ' tiene milisegundos

' Obtener el primer elemento de un filtro
objSheet.AutoFilter.Range.Offset(6).SpecialCells(12).Row

'   Quitar filtro aplicado
objSheet.Activate
objSheet.Cells.Select
objSheet.Selection.AutoFilter

' Try Catch
On Error Resume Next
' cuerpo del codigo
if Err.Number <> 0 then
	On Error Resum Next
	Err.Clear
Else

End If

' Elimina botones
objExcel.ActiveSheet.Shapes.Range(Array("Nombre del boton")).Select
objExcel.Selection.Delete

'Eliminar columnas
objSheet.Columns(6).Delete

' Filtro en correo electronico por fechas y subject
timeStart = " 00:00 AM"
timeEnd = " 23:59 PM"
fechaStart = fechaStart & timeStart 
fechaEnd = fechaEnd & timeEnd 

fechaStart = FormatDateTime(CDate(fechaStart), 2) &" "& FormatDateTime(CDate(fechaStart), 4) 
fechaEnd = FormatDateTime(CDate(fechaEnd), 2) &" "& FormatDateTime(CDate(fechaEnd), 4) 
subject = "ejemplo"

filter = "[ReceivedTime] >= '" & fechaStart & "'" & _
				 " And [ReceivedTime] <= '" & fechaEnd & "'" & _
				 " And [Subject] = '" & subject & "'"	



' Para obtener la primer celda visible disponible y la ultima fila disponible
fristRow = objSheet.Autofilter.Range.Offset(1).SpecialCells(12).Row
lastRow = objSheet.UsedRange.SpecialCells(11).Row

' pegado especial
objSheet.Range("A1").PasteSpecial -4163, -4142, 0, 0

' funcion para encryptar informacion de excel
Function EncryptFile(objSheet, objExcel, password)
    objSheet.Activate
    objExcel.ActivateSheet.Protect password, True. True, True
    EncryptFile = True
End Function ' EncryptFile

' funcinalidad para proteger con contrase;a el libro xlsx
Dim path, pass
Const sheetName = 1
pass = "1234"
path = "C:\test.xlsx"

Set objExcel = CreateObject("Excel.Application")
Set objBook = objExcel.Workbooks.Open(path)
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objSheet = objBook.WorkSheets(sheetName)

objSheet.Activate
objSheet.Cells(1,1).Value = "hola mundo"

objBook.Save
objBook.SaveAs path, , pass
objBook.Close

objExcel.Quit

' funcionalidad para abrir un archivo de excel y mapear las intacias que se van generando
Function formatDTYyyyymmdd(dateTime)
    formatDTYyyyymmdd = "(" & FormatDateTime(dateTime) & ")"
End Function ' formatDTYyyyymmdd

Function WriteLog(logLocation, lineType, message)
    If logLocation <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(logLocation) = False Then
            Set logFile = fso.CreateTextFile(logLocation)
        Else
            Set logFile = fso.OpenTextFile(logLocation, 8, False, 0)
        End If
        logFile.WriteLine(formatDTYyyyymmdd(Now) & " - " & lineType & " - " & message)
        logFile.Close
    End If
End Function ' WriteLog

Function CloseExcelInstance(infoLogFile, errorLogFile, nameScript)
    Const strComputer = "."
    Const findProc = "EXCEL.EXE"

    Set objWMIService = GetObject("Winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set processList = objWMIService.ExecQuery("Select Name from Win32_Process WHERE Name='" & findProc & "'")

    If processList.Count > 0 Then
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "taskkill /f /im excel.exe"
        If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Killing any lining Excel instance"
    Else
        If errorLogFile <> "" Then WriteLog errorLogFile, "INFO", nameScript & "It was not possible to close the Excel instance"
    End If
    Set objShell = Nothing
    Set objWMIService = Nothing
    Set processList = Nothing
End Function ' CloseExcelInstance

Function CreateExcelInstance(search, objShell, nameScript)
    Dim excelObj, counter
    counter = 0
    Do While (counter < 5)
        If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Attempt number " & counter & "CreateExcelInstance"
        Err.Clear
        On Error Resume Next
        Set excelObj = CreateObject("Excel.Application")
        If Err.Number = 0 Then
            If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Success Excel instance creation"
            Exit Do
        Else
            If infoLogFile <> "" Then WriteLog infoLogFile, "WARNING", nameScript & "Excel instance creation fail"
            Dim oShell : Set oShell = CreateObject("WScript.Shell")
            If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "Killing any lining Excel instance"
            oShell.Run "taskkill /f /im excel.exe"
            counter = counter + 1
        End If
    Loop
    Set CreateExcelInstance = excelObj
End Function ' CreateExcelInstance

Function OpenWorkBookFile(infoLogFile, errorLogFile, pathFile, nameSheet, excelObj, workBook, nameScript)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pathFile) =  False Then
        If errorLogFile <> "" Then WriteLog errorLogFile, "ERROR", nameScript & "File not found"
        WScript.Stdout.WriteLine "##Error##File not found"
    Else
        If infoLogFile <> "" Then WriteLog infoLogFile, "INFO", nameScript & "File fount it"
    End If

    ' create instance
    Set excelObj = CreateExcelInstance(infoLogFile, errorLogFile)
    Err.Clear
    excelObj.Visible = True
    excelObj.DisplayAlert = False

    ' open file
    Set workBook = excelObj.WorkBooks.Open(pathFile)
    excelObj.DisplayAlert = False

    ' assignamet obj in sheet
    Set objSheet = workBook.WorkSheet(nameSheet)
End Function ' CloseExcelFileOpen

' -----------------------------------------------
' the function to manupulate excel applications
' -----------------------------------------------

On Error Resume Next

' Dim nameScript
' nameScript = " "&Wscript.ScriptName
' MsgBox nameScript
If Err.Number <> 0 Then
    ' return
    ' WScript.StdOut.WriteLine "value retunr" 
    'WScript.StdOut.WriteLine Err.Description
Else
    'WScript.StdOut.WriteLine "value retunr"
End If


' convertir un archivo de txt a excel
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 

strFile = "archivo.txt"

Set objWorkbook = objExcel.Workbooks.Open(strFile) 

objWorkbook.SaveAs Replace(strFile, ".txt", ".xlsx") 
objWorkbook.Close 

objExcel.Quit

' formas de transformar una url del sp a local
'Note: We convert the URL of the document library into an UNC form in 
'the ConvertPath method. That means, it converts a 
'URL like http://YourSharePoint/DocLib into \\YourSharePoint\DocLib. However, if you 
'have configured HTTPS for your SharePoint, you need to convert 
'the URL into this form: \\YourSharePoint@SSL\DavWWWRoot\DocLib. In this case, 
'you should either extend the ConvertPath method, or simply use a fix path in your 
'code as a quick and dirty solution.

'for example

"https://company.sharepoint.com/sites/nameGrup/Lists/nameLista/AllItems.aspx"

'se remplaza

' convierte path
convertPath = Replace(path, " ", "%20")
convertPath = Replace(convertPath, "/", "\")
convertPath = Replace(convertPath, "https:", "")



'y que da como resultado

\\company.sharepoint.com@SSL\DavWWWRoot\sites\nameGrup