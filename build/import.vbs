Option Explicit

Dim Args
Dim Workbooks
Dim Modules
Dim Excel
Dim Workbook
Dim i
Dim KeepExcelOpen
Dim KeepWorkbookOpen

' Setup workbooks for import
' Optionally, pass workbook for import as argument
Set Args = Wscript.Arguments
If Args.Length > 0 Then
  Workbooks = Array(Args(0))
Else
  Workbooks = Array("Excel-REST - Blank.xlsm", "examples\Excel-REST - Example.xlsm", "specs\Excel-REST - Specs.xlsm")
End If

' Include all standard Excel-REST modules
Modules = Array("RestHelpers.bas", "IAuthenticator.cls", "RestClient.cls", "RestRequest.cls", "RestResponse.cls", "RestClientBase.bas")

' Open Excel
KeepExcelOpen = OpenExcel(Excel)
Excel.Visible = True
Excel.DisplayAlerts = False

For i = LBound(Workbooks) To UBound(Workbooks)
  WScript.Echo "Importing Excel-REST into " & Workbooks(i)
  KeepWorkbookOpen = OpenWorkbook(Excel, FullPath(Workbooks(i)), Workbook)
  ImportModules Workbook, ".\src\", Modules
  CloseWorkbook Workbook, KeepWorkbookOpen
Next

CloseExcel Excel, KeepExcelOpen

Set Workbook = Nothing
Set Excel = Nothing


''
' Module helpers
' ------------------------------------ '

Function RemoveModule(Workbook, Name)
  Dim Module
  Set Module = GetModule(Workbook, Name)

  If Not Module Is Nothing Then
    Workbook.VBProject.VBComponents.Remove Module
  End If
End Function

Function GetModule(Workbook, Name)
  Dim Module
  Set GetModule = Nothing

  For Each Module In Workbook.VBProject.VBComponents
    If Module.Name = Name Then
      Set GetModule = Module
      Exit Function
    End If
  Next
End Function

Sub ImportModule(Workbook, Folder, Filename)
  If VarType(Workbook) = vbObject Then
    RemoveModule Workbook, RemoveExtension(Filename)
    Workbook.VBProject.VBComponents.Import FullPath(Folder & Filename)
  End If
End Sub

Sub ImportModules(Workbook, Folder, Filenames)
  Dim i
  For i = LBound(Filenames) To UBound(Filenames)
    ImportModule Workbook, Folder, Filenames(i)
  Next
End Sub


''
' Excel helpers
' ------------------------------------ '

Function OpenWorkbook(Excel, Path, ByRef Workbook)
  On Error Resume Next

  Set Workbook = Excel.Workbooks(GetFilename(Path))

  If Workbook Is Nothing Or Err.Number <> 0 Then
    Set Workbook = Excel.Workbooks.Open(Path)
    OpenWorkbook = False
  Else
    OpenWorkbook = True
  End If

  Err.Clear
End Function

Function OpenExcel(Excel)
  On Error Resume Next
  
  Set Excel = GetObject(, "Excel.Application")

  If Excel Is Nothing Or Err.Number <> 0 Then
    Set Excel = CreateObject("Excel.Application")
    OpenExcel = False
  Else
    OpenExcel = True
  End If

  Err.Clear
End Function

Sub CloseWorkbook(ByRef Workbook, KeepWorkbookOpen)
  If Not KeepWorkbookOpen And VarType(Workbook) = vbObject Then
    Workbook.Close True
  End If

  Set Workbook = Nothing
End Sub

Sub CloseExcel(ByRef Excel, KeepExcelOpen)
  If Not KeepExcelOpen Then
    Excel.Quit
  End If

  Set Excel = Nothing
End Sub


''
' Filesystem helpers
' ------------------------------------ '

Function FullPath(Path)
  Dim FSO
  Set FSO = CreateObject("Scripting.FileSystemObject")
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
