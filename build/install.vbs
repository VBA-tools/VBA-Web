''
' Install v3.0.7
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Install Excel-REST and authenticators
' Run: cscript install.vbs
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Dim Args
Set Args = WScript.Arguments

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Excel
Dim ExcelWasOpen
Set Excel = Nothing
Dim Workbook
Dim WorkbookWasOpen
Set Workbook = Nothing
Dim Path

Dim ModulesFolder
Dim AuthenticatorsFolder
ModulesFolder = ".\src\"
AuthenticatorsFolder = ".\authenticators\"

Dim Modules
Modules = Array( _
  "RestHelpers.bas", _
  "IAuthenticator.cls", _
  "RestClient.cls", _
  "RestRequest.cls", _
  "RestResponse.cls" _
)

Dim Authenticators
Authenticators = Array( _
  "EmptyAuthenticator.cls", _
  "HttpBasicAuthenticator.cls", _
  "OAuth1Authenticator.cls", _
  "OAuth2Authenticator.cls", _
  "GoogleAuthenticator.cls", _
  "TwitterAuthenticator.cls", _
  "FacebookAuthenticator.cls", _
  "DigestAuthenticator.cls" _
)

Main

Sub Main()
  On Error Resume Next

  Print "Welcome to Excel-REST v3.0.7, let's get started!"
  
  ExcelWasOpen = OpenExcel(Excel)

  If Not Excel Is Nothing Then
    Install

    CloseExcel Excel, ExcelWasOpen
  ElseIf Err.Number <> 0 Then
    Print vbNewLine & "ERROR: Failed to open Excel" & vbNewLine & Err.Description
  End If

  Input vbNewLine & "All finished, thanks for using Excel-REST! Press any key to exit..."
End Sub

Sub Install
  Dim Success
  Path = Input(vbNewLine & _
    "In what Workbook would you like to install or update Excel-REST?" & vbNewLine & _
    "(e.g. C:\Users\Tim\DownloadStuff.xlsm)")
  Path = FullPath(Path)

  WorkbookWasOpen = OpenWorkbook(Excel, Path, Workbook)

  If Not Workbook Is Nothing Then
    If Not VBAIsTrusted(Workbook) Then
      Print vbNewLine & _
        "ERROR: In order to install Excel-REST," & vbNewLine & _
        "access to the VBA project object model needs to be trusted in Excel." & vbNewLine & vbNewLine & _
        "To enable:" & vbNewLine & _
        "Options > Trust Center > Trust Center Settings > Macro Settings > " & vbnewLine & _
        "Trust access to the VBA project object model"
    Else
      If Not AlreadyInstalled(Workbook) Then
        Success = InstallModules
      Else
        Dim ShouldUpgrade
        ShouldUpgrade = Input(vbNewLine & "Excel-REST appears to already be installed." & vbNewLine & vbNewLine & _
          "Warning: The currently installed Excel-REST files will be removed" & vbNewLine & _
          "and any previously made changes to those files will be lost" & vbNewLine & vbNewLine & _
          "Would you like to upgrade to v3.0.7? [yes/no]")

        If UCase(ShouldUpgrade) = "YES" Then
          Success = InstallModules
        Else
          Success = True
        End If
      End If
    End If

    If Success Then
      If UCase(Input(vbNewLine & "Would you like to install an authenticator (e.g. OAuth2)? [yes/no]")) = "YES" Then
        InstallAuthenticator
      End If
    End If

    CloseWorkbook Workbook, WorkbookWasOpen

    If UCase(Input(vbNewLine & "Would you like to install Excel-REST in another Workbook? [yes/no]")) = "YES" Then
      Install
    End If
  ElseIf Err.Number <> 0 Then
    Print vbNewLine & "ERROR: Failed to open Workbook" & vbNewLine & Err.Description
  End If
End Sub

Function InstallModules
  On Error Resume Next
  Dim i
  Dim Module
  Dim Backup
  Dim Backups
  ReDim Backups(UBound(Modules))

  WScript.StdOut.Write vbNewLine & "Installing Excel-REST..."

  For i = LBound(Modules) To UBound(Modules)
    ' Check for existing module and create backup if found
    Set Backups(i) = BackupModule(Workbook, RemoveExtension(Modules(i)), "backup__")
    
    If Err.Number <> 0 Then
      WScript.StdOut.Write "ERROR" & vbNewLine
      Print "Failed to backup previous version of Excel-REST" & vbNewLine & _
        "Please manually remove any existing Excel-REST files and try again"
      Exit For
    Else
      ' Import module
      ImportModule Workbook, ModulesFolder, Modules(i)

      If Err.Number <> 0 Then
        WScript.StdOut.Write "ERROR" & vbNewLine
        Print "Failed to install new version of Excel-REST" & vbNewLine & _
          "Any existing Excel-REST files will be now be attempted to be restored."
        Exit For
      End If
    End If
  Next

  If Err.Number <> 0 Then
    Err.Clear

    ' Restore backups
    For i = LBound(Modules) To UBound(Modules)
      RestoreModule Workbook, Modules(i), "backup__"
    Next
  Else
    ' Remove backups
    For i = LBound(Backups) To UBound(Backups)
      If Not Backups(i) Is Nothing Then
        Workbook.VBProject.VBComponents.Remove Backups(i)
      End If
    Next

    If Err.Number <> 0 Then
      WScript.StdOut.Write "ERROR" & vbNewLine
      Print "Excel-REST installed correctly," & vbNewLine & _
          "but failed to remove backups of the previous version" & vbNewLine & vbNewLine & _
          "It is safe to remove these files manually (backup__...)"
    End If
  End If

  If Err.Number = 0 Then
    WScript.StdOut.Write "Done!" & vbNewLine
    InstallModules = True
  End If
End Function

Sub InstallAuthenticator
  On Error Resume Next
  Dim i
  Dim Message
  Dim Install
  Dim Another
  Dim Backup

  Message = vbNewLine & "Which authenticator would you like to install?"
  For i = LBound(Authenticators) To UBound(Authenticators)
    Message = Message & vbNewLine & "- " & Replace(RemoveExtension(Authenticators(i)), "Authenticator", "")
  Next

  Install = Input(Message & vbNewLine & "[authenticator.../cancel]")
  If Install <> "" And UCase(Install) <> "CANCEL" Then
    For i = LBound(Authenticators) To UBound(Authenticators)
      If UCase(Install) = UCase(Replace(RemoveExtension(Authenticators(i)), "Authenticator", "")) Then
        WScript.StdOut.Write vbNewLine & "Installing " & Authenticators(i) & "..."

        Set Backup = BackupModule(Workbook, Authenticators(i), "backup__")

        If Err.Number <> 0 Then
          Err.Clear
          WScript.StdOut.Write "ERROR" & vbNewLine
          Print "Failed to backup previous version of " & Authenticators(i) & vbNewLine & _
            "Please manually remove it and try again"
        Else
          ImportModule Workbook, AuthenticatorsFolder, Authenticators(i)

          If Err.Number <> 0 Then
            Err.Clear
            WScript.StdOut.Write "ERROR" & vbNewLine
            Print "Failed to install new version of " & Authenticators(i) & vbNewLine & Err.Description

            RestoreModule Workbook, Authenticators(i), "backup__"
          Else
            If Not Backup Is Nothing Then
              Workbook.VBProject.VBComponents.Remove Backup
            End If

            If Err.Number <> 0 Then
              WScript.StdOut.Write "ERROR" & vbNewLine
              Print "Authenticator installed correctly," & vbNewLine & _
                  "but failed to remove the backup of the previous version" & vbNewLine & vbNewLine & _
                  "It is safe to remove this file manually (backup__...)"
            Else
              WScript.StdOut.Write "Done!" & vbNewLine
              Another = Input(vbNewLine & "Would you like to install another authenticator? [yes/no]")
              If UCase(Another) = "YES" Then
                InstallAuthenticator
              End If
            End If
          End If  
        End If
      End If
    Next
  End If
End Sub

Function AlreadyInstalled(ByRef Workbook)
  Dim i
  Dim Module
  For i = LBound(Modules) To UBound(Modules)
    Set Module = GetModule(Workbook, RemoveExtension(Modules(i)))
    If Not Module Is Nothing Then
      AlreadyInstalled = True
      Exit Function
    End If
  Next
End Function

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
      Path = Input(vbNewLine & _
        "Workbook not found at " & Path & vbNewLine & _
        "Would you like to try another location? [path.../cancel]")

      If UCase(Path) <> "CANCEL" And Path <> "" Then
        OpenWorkbook = OpenWorkbook(Excel, Path, Workbook)
      End If
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
' Check if VBA is trusted
'
' @param {Object} Workbook
' @param {Boolean}
Function VBAIsTrusted(Workbook)
  On Error Resume Next
  Dim Count
  Count = Workbook.VBProject.VBComponents.Count

  If Err.Number <> 0 Then
    Err.Clear
    VBAIsTrusted = False
  Else
    VBAIsTrusted = True
  End If
End Function

''
' Get module
'
' @param {Object} Workbook
' @param {String} Name
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

''
' Import module
'
' @param {Object} Workbook
' @param {String} Folder
' @param {String} Filename
Sub ImportModule(Workbook, Folder, Filename)
  Dim Module
  If Not Workbook Is Nothing Then
    ' Check for existing and remove
    Set Module = GetModule(Workbook, RemoveExtension(Filename))
    If Not Module Is Nothing Then
      Workbook.VBProject.VBComponents.Remove Module
    End If

    ' Import module
    Workbook.VBProject.VBComponents.Import FullPath(Folder & Filename)
  End If
End Sub

''
' Get module and backup (if found)
'
' @param {Object} Workbook
' @param {String} Name
' @param {String} Prefix
Function BackupModule(Workbook, Name, Prefix)
  Dim Backup
  Dim Existing
  Set Backup = GetModule(Workbook, Name)

  If Not Backup Is Nothing Then
    ' Remove any previous backups
    Set Existing = GetModule(Workbook, Prefix & Name)
    If Not Existing Is Nothing Then
      Workbook.VBProject.VBComponents.Remove Existing
    End If

    Backup.Name = Prefix & Name
  End If

  Set BackupModule = Backup
End Function

''
' Restore module from backup (if found)
'
' @param {Object} Workbook
' @param {String} Name
' @param {String} Prefix
Sub RestoreModule(Workbook, Name, Prefix)
  Dim Backup
  Dim Module
  Set Backup = GetModule(Workbook, Prefix & Name)

  If Not Backup Is Nothing Then
    ' Find upgraded module (and remove if found)
    Set Module = GetModule(Workbook, Name)
    If Not Module Is Nothing Then
      Workbook.VBProject.VBComponents.Remove Module
    End If

    ' Restore backup
    Backup.Name = Name
  End If
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