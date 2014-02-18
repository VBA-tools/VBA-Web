' Export specs from given workbook to given folder
'
' Example:
' (From cmd pointed to Excel-REST folder)
' cscript build\export-specs.vbs "specs\Excel-REST - Specs.xlsm" specs\
Option Explicit

Dim Args
Dim WBPath
Dim OutputPath
Dim Excel
Dim Workbook
Dim Modules
Dim ExcelWasOpen
Dim WorkbookWasOpen

Set Args = Wscript.Arguments
If Args.Length > 0 Then
  WBPath = Args(0)
  OutputPath = Args(1)
Else
  WBPath = "specs\Excel-REST - Specs.xlsm"
  OutputPath = "specs\"
End If

' Setup modules to export
Modules = Array(_
  "RestClientSpecs.bas", _
  "RestClientAsyncSpecs.bas", _
  "RestRequestSpecs.bas", _
  "RestHelpersSpecs.bas", _
  "RestClientBaseSpecs.bas" _
)

If WBPath <> "" And OutputPath <> "" Then
  WScript.Echo "Exporting Excel-REST specs from " & WBPath & " to " & OutputPath

  ExcelWasOpen = OpenExcel(Excel)
  Excel.Visible = True
  Excel.DisplayAlerts = False

  ' Get workbook path relative to root Excel-REST project
  WBPath = FullPath(WBPath)
  OutputPath = FullPath(OutputPath)

  If Right(OutputPath, 1) <> "\" Then
    OutputPath = OutputPath & "\"
  End If
  
  ' Open workbook
  WorkbookWasOpen = OpenWorkbook(Excel, WBPath, Workbook)

  Dim i
  Dim Module
  For i = LBound(Modules) To UBound(Modules)
    Set Module = GetModule(Workbook, RemoveExtension(Modules(i)))

    If Not Module Is Nothing Then
      Module.Export OutputPath & Modules(i)
    End If
  Next

  CloseWorkbook Workbook, WorkbookWasOpen
  CloseExcel Excel, ExcelWasOpen

  Set Workbook = Nothing
  Set Excel = Nothing
End If


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
