VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpBasicAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IAuthenticator
''
' HttpBasicAuthenticator v1.0.0
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Utilize http basic authentication
'
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Username As String
Public Password As String


' ============================================= '
' Public Methods
' ============================================= '

''
' Helper for setting required parameters at once
'
' @param {String} Username
' @param {String} Password
' --------------------------------------------- '

Public Sub Setup(Username As String, Password As String)
    Me.Username = Username
    Me.Password = Password
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub IAuthenticator_BeforeExecute(Request As RestRequest)
    Request.AddHeader "Authorization", CreateHeader()
End Sub

Private Sub IAuthenticator_HttpOpen(Http As Object, Request As RestRequest, BaseUrl As String, Optional UseAsync As Boolean = False)
    ' Use http open with username and password values set
    ' (This is used in addition to setting request header, as some services required this)
    Http.Open Request.MethodName(), Request.FullUrl(BaseUrl), UseAsync, Me.Username, Me.Password
End Sub

Private Function CreateHeader() As String
    ' Create Basic Auth request header
    Dim Header As String
    Header = "Basic " & EncodeStringToBase64(Me.Username & ":" & Me.Password)
    CreateHeader = Header
End Function

