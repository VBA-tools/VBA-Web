VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EH_HttpBasicAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements EH_IAuthenticator
''
' EH_HttpBasicAuthenticator v0.2.2
' (c) Tim Hall - https://github.com/timhall/ExcelHelpers
'
' Utilize http basic authentication
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @version 0.2.2
' @date 20121024
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

Private Sub EH_IAuthenticator_BeforeExecute(request As EH_RestRequest)
    Call request.AddHeader("Authorization", CreateHeader())
End Sub

Private Sub EH_IAuthenticator_HttpOpen(http As MSXML2.IXMLHTTPRequest, request As EH_RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Use http open with username and password values set
    ' (This is used in addition to setting request header, as some services required this)
    Call http.Open(request.MethodName(), request.FullUrl(BaseUrl), useAsync, Me.Username, Me.Password)
End Sub

Private Function CreateHeader() As String
    ' Create Basic Auth request header
    Dim header As String
    header = "Basic " & EncodeStringToBase64(Me.Username & ":" & Me.Password)
    CreateHeader = header
End Function
