''
' Install Excel-REST
'
' Run: cscript install.vbs
'
' (c) Tim Hall - https://github.com/timhall/Excel-REST
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Dim Args
Set Args = WScript.Arguments

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim Excel
Dim ExcelWasOpen
Set Excel = Nothing

Main

Sub Main()
  On Error Resume Next

  Print "----------------------------------------------------------------" & vbNewLine & _
    "Excel-REST" & vbNewLine & vbNewLine & _
    "Welcome to the Excel-REST installer!" & vbNewLine & _
    "This will walk you through installing Excel-REST in your project" & vbNewLine & _
    "----------------------------------------------------------------" & vbNewLine

  ExcelWasOpen = OpenExcel(Excel)

  If Not Excel Is Nothing Then
    Install
  ElseIf Err.Number <> 0 Then
    Print "ERROR: Failed to open Excel" & vbNewLine & Err.Description
  End If

  CloseExcel Excel, ExcelWasOpen

  Input vbNewLine & "All Finished! Press Enter to exit..."
End Sub

Sub Install
  Dim Path
  Path = Input("In what Workbook that you would like to install Excel-REST?" & vbNewLine & "(e.g. C:\Users\Tim\DownloadStuff.xlsm)")
  Path = FullPath(Path)

  Dim Workbook
  Dim WorkbookWasOpen
  Set Workbook = Nothing
  WorkbookWasOpen = OpenWorkbook(Excel, Path, Workbook)

  If Not Workbook Is Nothing Then
    ' TODO Install
  ElseIf Err.Number <> 0 Then
    Print "ERROR: Failed to open Workbook" & vbNewLine & Err.Description
  End If

  CloseWorkbook Workbook, WorkbookWasOpen
End Sub


''
' Excel helpers
' ------------------------------------ '

''
' Open Workbook and return whether Workbook was already open
'
' @param {Object} Excel
' @param {String} Path
' @param {Object} Workbook object to load Workbook into
' @return {Boolean} Workbook was already open
Function OpenWorkbook(Excel, Path, ByRef Workbook)
  On Error Resume Next

  Path = FullPath(Path)
  Set Workbook = Excel.Workbooks(GetFilename(Path))

  If Workbook Is Nothing Or Err.Number <> 0 Then
    Err.Clear

    If FileExists(Path) Then
      Set Workbook = Excel.Workbooks.Open(Path)
    Else
      Print "Workbook not found at " & Path
      ' TODO Create workbook if it doesn't exist
      'Dim CreateWorkbook
      'CreateWorkbook = Input("Workbook not found at " & Path & vbNewLine & "Would you like to create it, yes or no? (yes)")
      '
      'If UCase(CreateWorkbook) = "YES" Or CreateWorkbook = "" Then
      '  Print "Create workbook..."
      'End If
    End If
    OpenWorkbook = False
  Else
    OpenWorkbook = True
  End If
End Function

''
' Close Workbook and save changes 
' (keep open without saving changes if previously open)
'
' @param {Object} Workbook
' @param {Boolean} KeepWorkbookOpen
Sub CloseWorkbook(ByRef Workbook, KeepWorkbookOpen)
  If Not KeepWorkbookOpen And Not Workbook Is Nothing Then
    Workbook.Close True
  End If

  Set Workbook = Nothing
End Sub

''
' Open Excel and return whether Excel was already open
'
' @param {Object} Excel object to load Excel into
' @return {Boolean} Excel was already open
Function OpenExcel(ByRef Excel)
  On Error Resume Next

  Set Excel = GetObject(, "Excel.Application")

  If Excel Is Nothing Or Err.Number <> 0 Then
    Err.Clear

    Set Excel = CreateObject("Excel.Application")
    OpenExcel = False
  Else
    OpenExcel = True
  End If
End Function

''
' Close Excel (keep open if previously open)
'
' @param {Object} Excel
' @param {Boolean} KeepExcelOpen
Sub CloseExcel(ByRef Excel, KeepExcelOpen)
  If Not KeepExcelOpen And Not Excel Is Nothing Then
    Excel.Quit  
  End If

  Set Excel = Nothing
End Sub

''
' Filesystem helpers
' ------------------------------------ '

Function FullPath(Path)
  FullPath = FSO.GetAbsolutePathName(Path)
End Function

Function GetFilename(Path)
  Dim Parts
  Parts = Split(Path, "\")

  GetFilename = Parts(UBound(Parts))
End Function

Function RemoveExtension(Name)
    Dim Parts
    Parts = Split(Name, ".")
    
    If UBound(Parts) > LBound(Parts) Then
        ReDim Preserve Parts(UBound(Parts) - 1)
    End If
    
    RemoveExtension = Join(Parts, ".")
End Function

Function FileExists(Path)
  FileExists = FSO.FileExists(Path)
End Function

''
' General helpers
' ------------------------------------ '

Sub Print(Message)
  WScript.Echo Message  
End Sub

Function Input(Prompt)
  If Prompt <> "" Then
    WScript.StdOut.Write Prompt & " "
  End If

  Input = WScript.StdIn.ReadLine 
End Function