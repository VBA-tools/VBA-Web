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
' HttpBasicAuthenticator v0.2.1
' (c) Tim Hall - https://github.com/timhall/ExcelExperiments
'
' Utilize http basic authentication
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @version 0.2.0
' @date 20120724
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
Public Sub Setup(Username As String, Password As String)
    Me.Username = Username
    Me.Password = Password
End Sub

' ============================================= '
' Private Methods
' ============================================= '

Private Sub IAuthenticator_BeforeExecute(request As RestRequest)
    Call request.AddHeader("Authorization", CreateHeader())
End Sub

Private Sub IAuthenticator_HttpOpen(http As MSXML2.IXMLHTTPRequest, request As RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Open http
    Call http.Open(request.MethodName(), request.FullUrl(BaseUrl), useAsync, Me.Username, Me.Password)
End Sub

Private Function CreateHeader() As String
    Dim header As String
    header = "Basic " & EncodeStringToBase64(Me.Username & ":" & Me.Password)
    CreateHeader = header
End Function
