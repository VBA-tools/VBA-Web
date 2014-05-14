VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' RestResponse v3.0.5
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Wrapper for http responses
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public StatusCode As StatusCodes
Public StatusDescription As String
Public Content As String
Public Data As Variant
Public Body As Variant
Public Headers As Collection
Public Cookies As Dictionary

Private Sub Class_Initialize()
    Body = Array()
    Set Headers = New Collection
    Set Cookies = New Dictionary
End Sub
