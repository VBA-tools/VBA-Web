Option Explicit

Dim Args
Dim WBPath
Dim OutputPath
Dim Excel
Dim Workbook
Dim Modules

Set Args = Wscript.Arguments
WBPath = Args(0)
OutputPath = Args(1)

' Setup modules to export
Modules = Array("RestHelpers.bas", "IAuthenticator.cls", "RestClient.cls", "RestRequest.cls", "RestResponse.cls")

If WBPath <> "" And OutputPath <> "" Then
  Set Excel = CreateObject("Excel.Application")
  Excel.Visible = True
  Excel.DisplayAlerts = False

  ' Get workbook path relative to root Excel-REST project
  Dim FSO
  Set FSO = CreateObject("Scripting.FileSystemObject")
  WBPath = FSO.GetAbsolutePathName(WBPath)
  OutputPath = FSO.GetAbsolutePathName(OutputPath)

  If Right(OutputPath, 1) <> "\" Then
    OutputPath = OutputPath & "\"
  End If
  
  ' Open workbook
  Set Workbook = Excel.Workbooks.Open(WBPath)

  Dim i
  Dim Module
  For i = LBound(Modules) To UBound(Modules)
    Set Module = GetModule(Workbook, RemoveExtension(Modules(i)))

    If Not Module Is Nothing Then
      Module.Export OutputPath & Modules(i)
    End If
  Next

  Workbook.Close True
  Excel.Quit

  Set Workbook = Nothing
  Set Excel = Nothing
End If

Function GetModule(Workbook, Name)
  Dim Module

  For Each Module In Workbook.VBProject.VBComponents
    If Module.Name = Name Then
      Set GetModule = Module
      Exit Function
    End If
  Next
End Function

Function RemoveExtension(Name)
    Dim Parts
    Parts = Split(Name, ".")
    
    If UBound(Parts) > LBound(Parts) Then
        ReDim Preserve Parts(UBound(Parts) - 1)
    End If
    
    RemoveExtension = Join(Parts, ".")
End Function
