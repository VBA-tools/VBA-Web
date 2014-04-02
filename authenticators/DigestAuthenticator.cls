VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DigestAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Digest Authenticator v2.0.3
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Custom IAuthenticator for Digest Authentication
' http://en.wikipedia.org/wiki/Digest_access_authentication
'
' @implements: IAuthenticator v3.*
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IAuthenticator
Option Explicit

Private Const qop As String = "auth"
Private pClientNonce As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Username As String
Public Password As String
Public Realm As String
Public ServerNonce As String
Public RequestCount As Long
Public Opaque As String

Public Property Get ClientNonce() As String
    If pClientNonce = "" Then
        pClientNonce = RestHelpers.CreateNonce
    End If
    ClientNonce = pClientNonce
End Property
Public Property Let ClientNonce(Value As String)
    pClientNonce = Value
End Property

Public Property Get IsAuthenticated() As Boolean
    If ServerNonce <> "" Then
        IsAuthenticated = True
    End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup authenticator
' --------------------------------------------- '

Public Sub Setup(Username As String, Password As String)
    Me.Username = Username
    Me.Password = Password
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {RestClient} Client The client that is about to execute the request
' @param {RestRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IAuthenticator_BeforeExecute(ByVal Client As RestClient, ByRef Request As RestRequest)
    ' Add headers, cookies, etc to `Request` before it is executed
    ' (Leave blank to pass Request through unmodified)
    If Me.IsAuthenticated Then
        Me.RequestCount = Me.RequestCount + 1
        Request.AddHeader "Authorization", CreateHeader(Client, Request)
    End If
End Sub

''
' Hook for taking action after request has been executed
'
' @param {RestClient} Client The client that executed request
' @param {RestRequest} Request The request that was just executed
' @param {RestResponse} Response to request
' --------------------------------------------- '

Private Sub IAuthenticator_AfterExecute(ByVal Client As RestClient, ByVal Request As RestRequest, ByRef Response As RestResponse)
    If Response.StatusCode = 401 And Not Me.IsAuthenticated Then
        ExtractAuthenticateInformation Response
        
        Request.AddHeader "Authorization", CreateHeader(Client, Request)
        RestHelpers.UpdateResponse Response, Client.Execute(Request)
    End If
End Sub

''
' Hook for overriding standard http open (used for HTTP Basic)
'
' @param {MSXML2.IXMLHTTPRequest} http
' @parma {RestClient} Client The client that is about to open request
' @param {RestRequest} Request The request about to be opened
' @param {String} BaseUrl
' @param {Boolean} [useAsync=False]
' --------------------------------------------- '

Private Sub IAuthenticator_HttpOpen(ByRef Http As Object, ByVal Client As RestClient, ByRef Request As RestRequest, BaseUrl As String, Optional UseAsync As Boolean = False)
    ' Perform standard http open
    Call Http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), UseAsync)
End Sub

Public Function CreateHeader(Client As RestClient, Request As RestRequest) As String
    Dim Uri As String
    Uri = RestHelpers.UrlParts(Request.FullUrl(Client.BaseUrl))("Uri")

    CreateHeader = "Digest" & _
        " username=""" & Me.Username & """" & _
        ", realm=""" & Me.Realm & """" & _
        ", nonce=""" & Me.ServerNonce & """" & _
        ", uri=""" & Uri & """" & _
        ", qop=" & qop & _
        ", nc=" & FormattedRequestCount & _
        ", cnonce=""" & Me.ClientNonce & """" & _
        ", response=""" & CalculateResponse(Client, Request) & """" & _
        ", opaque=""" & Me.Opaque & """"
End Function

Public Function CalculateResponse(Client As RestClient, Request As RestRequest) As String
    Dim HA1 As String
    Dim HA2 As String
    Dim Uri As String
    Uri = RestHelpers.UrlParts(Request.FullUrl(Client.BaseUrl))("Uri")
    
    HA1 = CalculateHA1
    HA2 = CalculateHA2(Request.MethodName, Uri)
    
    CalculateResponse = RestHelpers.MD5(HA1 & ":" & Me.ServerNonce & ":" & FormattedRequestCount & ":" & Me.ClientNonce & ":" & qop & ":" & HA2)
End Function

' Extract authentication information from 401 response headers
Public Sub ExtractAuthenticateInformation(Response As RestResponse)
    Dim Header As Dictionary
    For Each Header In Response.Headers
        ' Find authentication header
        If Header("key") = "WWW-Authenticate" Then
            ' Make sure using Digest authentication
            If Left(Header("value"), 6) = "Digest" Then
                Dim Lines As Variant
                Lines = Split(Mid$(Header("value"), 7), vbCrLf)
                
                Dim i As Integer
                Dim Key As String
                Dim Value As String
                For i = LBound(Lines) To UBound(Lines)
                    Key = LCase(Trim(Mid$(Lines(i), 1, InStr(1, Lines(i), "=") - 1)))
                    Value = Trim(Mid$(Lines(i), InStr(1, Lines(i), "=") + 1, Len(Lines(i))))
                    
                    ' Remove quotes and trailing comma
                    Value = Replace(Value, """", "")
                    If Right(Value, 1) = "," Then Value = Left(Value, Len(Value) - 1)
                    
                    ' Find realm, nonce, and opaque
                    If Key = "realm" Then Me.Realm = Value
                    If Key = "nonce" Then Me.ServerNonce = Value
                    If Key = "opaque" Then Me.Opaque = Value
                Next i
            End If
            
            Exit Sub
        End If
    Next Header
End Sub

Public Function CalculateHA1() As String
    CalculateHA1 = MD5(Me.Username & ":" & Me.Realm & ":" & Me.Password)
End Function

Public Function CalculateHA2(Method As String, Uri As String) As String
    CalculateHA2 = MD5(Method & ":" & Uri)
End Function

' Pad request count to 8 places
Public Function FormattedRequestCount() As String
    FormattedRequestCount = Right("00000000" & Me.RequestCount, 8)
End Function
