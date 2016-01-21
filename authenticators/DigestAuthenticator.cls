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
' Digest Authenticator v3.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for Digest Authentication
' http://en.wikipedia.org/wiki/Digest_access_authentication
'
' @class DigestAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_Qop As String = "auth"
Private auth_pClientNonce As String

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
    If auth_pClientNonce = "" Then
        auth_pClientNonce = WebHelpers.CreateNonce
    End If
    ClientNonce = auth_pClientNonce
End Property
Public Property Let ClientNonce(Value As String)
    auth_pClientNonce = Value
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
'
' @param {String} Username
' @param {String} Password
''
Public Sub Setup(Username As String, Password As String)
    Me.Username = Username
    Me.Password = Password
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    If Me.IsAuthenticated Then
        Me.RequestCount = Me.RequestCount + 1
        Request.SetHeader "Authorization", CreateHeader(Client, Request)
    End If
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    If Response.StatusCode = 401 And Not Me.IsAuthenticated Then
        WebHelpers.LogDebug "Extract Authenticate and retry 401 request " & Client.GetFullUrl(Request), "Digest.AfterExecute"
        ExtractAuthenticateInformation Response
        
        Request.SetHeader "Authorization", CreateHeader(Client, Request)
        Response.Update Client.Execute(Request)
    End If
End Sub

''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
    ' e.g. Update option, headers, etc.
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' http://curl.haxx.se/docs/manpage.html#--digest
    Curl = Curl & " --digest --user " & WebHelpers.PrepareTextForShell(Me.Username) & ":" & WebHelpers.PrepareTextForShell(Me.Password)
End Sub

''
' Create digest header for given Client and Request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function CreateHeader(Client As WebClient, Request As WebRequest) As String
    Dim auth_Uri As String
    auth_Uri = WebHelpers.GetUrlParts(Client.GetFullUrl(Request))("Path")

    CreateHeader = "Digest " & _
        "username=""" & Me.Username & """, " & _
        "realm=""" & Me.Realm & """, " & _
        "nonce=""" & Me.ServerNonce & """, " & _
        "uri=""" & auth_Uri & """, " & _
        "qop=" & auth_Qop & ", " & _
        "nc=" & web_FormattedRequestCount & ", " & _
        "cnonce=""" & Me.ClientNonce & """, " & _
        "response=""" & web_CalculateResponse(Client, Request) & """, " & _
        "opaque=""" & Me.Opaque & """"
        
    WebHelpers.LogDebug CreateHeader, "DigestAuthenticator.CreateHeader"
End Function

''
' Extract authentication information from 401 response headers
'
' @internal
' @param {WebResponse} Response
''
Public Sub ExtractAuthenticateInformation(Response As WebResponse)
    Dim auth_Header As String
    Dim web_CrLf As String
    
    auth_Header = WebHelpers.FindInKeyValues(Response.Headers, "WWW-Authenticate")
    web_CrLf = VBA.Chr$(13) & VBA.Chr$(10)
    
    If auth_Header <> "" And VBA.Left$(auth_Header, 6) = "Digest" Then
        Dim auth_Lines As Variant
        auth_Lines = VBA.Split(VBA.Mid$(auth_Header, 7), web_CrLf)
        
        Dim auth_i As Integer
        Dim auth_Key As String
        Dim auth_Value As String
        For auth_i = LBound(auth_Lines) To UBound(auth_Lines)
            auth_Key = VBA.LCase$(VBA.Trim$(VBA.Mid$(auth_Lines(auth_i), 1, VBA.InStr(1, auth_Lines(auth_i), "=") - 1)))
            auth_Value = VBA.Trim$(VBA.Mid$(auth_Lines(auth_i), VBA.InStr(1, auth_Lines(auth_i), "=") + 1, VBA.Len(auth_Lines(auth_i))))
            
            ' Remove quotes and trailing comma
            auth_Value = VBA.Replace(auth_Value, """", "")
            If VBA.Right$(auth_Value, 1) = "," Then
                auth_Value = VBA.Left$(auth_Value, VBA.Len(auth_Value) - 1)
            End If
            
            ' Find realm, nonce, and opaque
            If auth_Key = "realm" Then Me.Realm = auth_Value
            If auth_Key = "nonce" Then Me.ServerNonce = auth_Value
            If auth_Key = "opaque" Then Me.Opaque = auth_Value
        Next auth_i
        
        WebHelpers.LogDebug "realm=" & Me.Realm & ", nonce=" & Me.ServerNonce & ", opaque=" & Me.Opaque, "DigestAuthenticator.ExtractAuthenticateInformation"
    End If
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Calculate digest response fro given Client and Request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Private Function web_CalculateResponse(web_Client As WebClient, web_Request As WebRequest) As String
    Dim auth_HA1 As String
    Dim auth_HA2 As String
    Dim auth_Uri As String
    
    auth_Uri = WebHelpers.GetUrlParts(web_Client.GetFullUrl(web_Request))("Path")
    auth_HA1 = web_CalculateHA1
    auth_HA2 = web_CalculateHA2(WebHelpers.MethodToName(web_Request.Method), auth_Uri)
    
    web_CalculateResponse = WebHelpers.MD5(auth_HA1 & ":" & Me.ServerNonce & ":" & web_FormattedRequestCount & ":" & Me.ClientNonce & ":" & auth_Qop & ":" & auth_HA2)
End Function

''
' Calculate HA1 portion of digest response
'
' @internal
' @return {String}
''
Private Function web_CalculateHA1() As String
    web_CalculateHA1 = WebHelpers.MD5(Me.Username & ":" & Me.Realm & ":" & Me.Password)
End Function

''
' Calculate HA1 portion of digest response
'
' @internal
' @return {String}
''
Private Function web_CalculateHA2(web_Method As String, web_Uri As String) As String
    web_CalculateHA2 = WebHelpers.MD5(web_Method & ":" & web_Uri)
End Function

''
' Pad request count to 8 places
'
' @internal
' @return {String}
''
Private Function web_FormattedRequestCount() As String
    web_FormattedRequestCount = Right("00000000" & Me.RequestCount, 8)
End Function
