VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuth1Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' OAuth1 Authenticator v3.0.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' OAuth1 authenticator
'
' @class OAuth1Authenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_SignatureMethod As String = "HMAC-SHA1"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ConsumerKey As String
Public ConsumerSecret As String
Public Token As String
Public TokenSecret As String
Public Realm As String
Public Nonce As String
Public Timestamp As String

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup
'
' @param {String} ConsumerKey
' @param {String} ConsumerSecret
' @param {String} Token
' @param {String} TokenSecret
' @param {String} [Realm]
''
Public Sub Setup(ConsumerKey As String, ConsumerSecret As String, _
    Token As String, TokenSecret As String, Optional Realm As String = "")
    
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
    Me.Token = Token
    Me.TokenSecret = TokenSecret
    Me.Realm = Realm
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' Add authorization header to request
    Request.SetHeader "Authorization", CreateHeader(Client, Request)
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
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
    ' e.g. Add flags to cURL
End Sub

''
' Create header for given client and request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function CreateHeader(auth_Client As WebClient, auth_Request As WebRequest) As String
    Dim auth_Nonce As String
    Dim auth_Timestamp As String
    Dim auth_Base As String
    Dim auth_SigningKey As String
    Dim auth_Signature As String
    
    ' Load or create nonce and timestamp
    If Me.Nonce <> "" Then
        auth_Nonce = Me.Nonce
    Else
        auth_Nonce = WebHelpers.CreateNonce()
    End If
    If Me.Timestamp <> "" Then
        auth_Timestamp = Me.Timestamp
    Else
        auth_Timestamp = auth_CreateTimestamp
    End If
    
    ' Create needed parts of authorization header
    auth_Base = CreateBaseString(auth_Nonce, auth_Timestamp, auth_Client, auth_Request)
    auth_SigningKey = auth_CreateSigningKey()
    auth_Signature = CreateSignature(auth_Base, auth_SigningKey)
    
    ' Generate header
    CreateHeader = "OAuth "
    
    ' Add realm (if exists)
    If Me.Realm <> "" Then
        CreateHeader = CreateHeader & "realm=""" & Me.Realm & """, "
    End If
    
    ' Construct header parts
    ' [OAuth Core 1.0 Revision A](http://oauth.net/core/1.0a/)
    CreateHeader = CreateHeader & "oauth_consumer_key=""" & Me.ConsumerKey & """, "
    CreateHeader = CreateHeader & "oauth_nonce=""" & auth_Nonce & """, "
    CreateHeader = CreateHeader & "oauth_signature=""" & WebHelpers.UrlEncode(auth_Signature) & """, "
    CreateHeader = CreateHeader & "oauth_signature_method=""" & auth_SignatureMethod & """, "
    CreateHeader = CreateHeader & "oauth_timestamp=""" & auth_Timestamp & """, "
    CreateHeader = CreateHeader & "oauth_token=""" & Me.Token & """, "
    CreateHeader = CreateHeader & "oauth_version=""" & "1.0" & """"
    
    CreateHeader = CreateHeader
End Function

''
' Create base string for given parameters
'
' @internal
' @param {String} Nonce
' @param {String} Timestamp
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function CreateBaseString(auth_Nonce As String, auth_Timestamp As String, auth_Client As WebClient, auth_Request As WebRequest) As String
    Dim auth_Base As String
    Dim auth_Parameters As String
    
    ' Check for parameters and add to auth_Base if present
    auth_Parameters = GetRequestParameters(auth_Client, auth_Request)
    If auth_Parameters <> "" Then
        auth_Base = auth_Parameters & "&"
    End If
    
    auth_Base = auth_Base & "oauth_consumer_key" & "=" & Me.ConsumerKey
    auth_Base = auth_Base & "&" & "oauth_nonce" & "=" & auth_Nonce
    auth_Base = auth_Base & "&" & "oauth_signature_method" & "=" & auth_SignatureMethod
    auth_Base = auth_Base & "&" & "oauth_timestamp" & "=" & auth_Timestamp
    auth_Base = auth_Base & "&" & "oauth_token" & "=" & Me.Token
    auth_Base = auth_Base & "&" & "oauth_version=1.0"
    
    CreateBaseString = WebHelpers.MethodToName(auth_Request.Method) & "&" & _
        WebHelpers.UrlEncode(GetRequestUrl(auth_Client, auth_Request), EncodeUnsafe:=False) & "&" & _
        WebHelpers.UrlEncode(auth_Base)
End Function

''
' Create signature with given parameters
'
' @internal
' @param {String} Base
' @param {String} SigningKey
' @return {String}
''
Public Function CreateSignature(auth_Base As String, auth_SigningKey As String) As String
    CreateSignature = WebHelpers.HMACSHA1(auth_Base, auth_SigningKey, "Base64")
End Function

''
' Create request url for given client and request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function GetRequestUrl(auth_Client As WebClient, auth_Request As WebRequest) As String
' From OAuth 1.0 Docs
' http://oauth.net/core/1.0/#anchor14
'
' The Signature Base String includes the request absolute URL, tying the signature to a specific endpoint.
' The URL used in the Signature Base String MUST include the scheme, authority, and path, and MUST exclude the query and fragment as defined by [RFC3986] section 3.
'
' If the absolute request URL is not available to the Service Provider (it is always available to the Consumer),
' it can be constructed by combining the scheme being used, the HTTP Host header, and the relative HTTP request URL.
' If the Host header is not available, the Service Provider SHOULD use the host name communicated to the Consumer in the documentation or other means.
'
' The Service Provider SHOULD document the form of URL used in the Signature Base String to avoid ambiguity due to URL normalization.
' Unless specified, URL scheme and authority MUST be lowercase and include the port number; http default port 80 and https default port 443 MUST be excluded.

    Dim auth_Parts As Dictionary
    Set auth_Parts = WebHelpers.GetUrlParts(auth_Client.GetFullUrl(auth_Request))
    
    ' Url scheme and authority MUST be lowercase
    GetRequestUrl = LCase(auth_Parts("Protocol") & "://" & auth_Parts("Host"))
    
    ' Include port (80 and 443 MUST be excluded)
    If auth_Parts("Port") <> 80 And auth_Parts("Port") <> 443 Then
        GetRequestUrl = GetRequestUrl & ":" & auth_Parts("Port")
    End If
    
    ' Include path
    GetRequestUrl = GetRequestUrl + auth_Parts("Path")
    
    ' MUST exclude query and fragment
End Function

''
' Create request parameters for given client and request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function GetRequestParameters(auth_Client As WebClient, auth_Request As WebRequest) As String
    ' TODO Sort parameters by key then value
    
    Dim auth_Parts As Dictionary
    Set auth_Parts = WebHelpers.GetUrlParts(auth_Client.GetFullUrl(auth_Request))
    
    ' Remove leading ?
    GetRequestParameters = auth_Parts("Querystring")
    
    ' Replace + for spaces with %20
    GetRequestParameters = Replace(GetRequestParameters, "+", "%20")
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function auth_CreateSigningKey() As String
    auth_CreateSigningKey = Me.ConsumerSecret & "&" & Me.TokenSecret
End Function

Private Function auth_CreateTimestamp() As String
    auth_CreateTimestamp = VBA.CStr(VBA.DateDiff("s", #1/1/1970#, WebHelpers.ConvertToUtc(VBA.Now)))
End Function
