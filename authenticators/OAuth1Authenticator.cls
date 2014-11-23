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
' OAuth1 Authenticator v2.0.6
' (c) Tim Hall - https://github.com/timhall/VBA-Web
'
' Utilize OAuth1 authentication
'
' @implements: IWebAuthenticator v3.*
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' Declarations
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME)

' Constants and private variables
Private Const SignatureMethod As String = "HMAC-SHA1"

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private Enum TIME_ZONE
    TIME_ZONE_ID_INVALID = 0
    TIME_ZONE_STANDARD = 1
    TIME_ZONE_DAYLIGHT = 2
End Enum

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
' Helper for setting required parameters at once
'
' @param {String} ConsumerKey
' @param {String} ConsumerSecret
' @param {String} Token
' @param {String} TokenSecret
' @param {String} [Realm]
' --------------------------------------------- '

Public Sub Setup(ConsumerKey As String, ConsumerSecret As String, _
    Token As String, TokenSecret As String, Optional Realm As String = "")
    
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
    Me.Token = Token
    Me.TokenSecret = TokenSecret
    Me.Realm = Realm
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param {WebRequest} Request The request about to be executed
' --------------------------------------------- '

Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' Add authorization header to request
    Request.AddHeader "Authorization", CreateHeader(Client, Request)
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param {WebResponse} Response to request
' --------------------------------------------- '

Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)

End Sub

Public Function CreateHeader(Client As WebClient, Request As WebRequest) As String
    Dim Header As String
    Dim Nonce As String
    Dim Timestamp As String
    Dim base As String
    Dim signingKey As String
    Dim signature As String
    
    ' Load or create nonce and timestamp
    If Me.Nonce <> "" Then
        Nonce = Me.Nonce
    Else
        Nonce = WebHelpers.CreateNonce()
    End If
    If Me.Timestamp <> "" Then
        Timestamp = Me.Timestamp
    Else
        Timestamp = CreateTimestamp
    End If
    
    ' Create needed parts of authorization header
    base = CreateBaseString(Nonce, Timestamp, Client, Request)
    signingKey = CreateSigningKey()
    signature = CreateSignature(base, signingKey)
    
    ' Add realm (if exists)
    If Me.Realm <> "" Then
        Header = "OAuth realm=" & Chr(34) & Me.Realm & Chr(34) & ", "
    Else
        Header = "OAuth "
    End If
    
    ' Construct header parts
    ' [OAuth Core 1.0 Revision A](http://oauth.net/core/1.0a/)
    Header = Header & "oauth_consumer_key=" & Chr(34) & Me.ConsumerKey & Chr(34) & ", "
    Header = Header & "oauth_nonce=" & Chr(34) & Nonce & Chr(34) & ", "
    Header = Header & "oauth_signature=" & Chr(34) & WebHelpers.UrlEncode(signature) & Chr(34) & ", "
    Header = Header & "oauth_signature_method=" & Chr(34) & SignatureMethod & Chr(34) & ", "
    Header = Header & "oauth_timestamp=" & Chr(34) & Timestamp & Chr(34) & ", "
    Header = Header & "oauth_token=" & Chr(34) & Me.Token & Chr(34) & ", "
    Header = Header & "oauth_version=" & Chr(34) & "1.0" & Chr(34)
    CreateHeader = Header
    
    WebHelpers.LogDebug CreateHeader, "OAuth1Authenticator.CreateHeader"
End Function

Public Function CreateBaseString(Nonce As String, Timestamp As String, Client As WebClient, Request As WebRequest) As String
    Dim base As String
    
    ' Check for parameters and add to base if present
    Dim Parameters As String
    Parameters = RequestParameters(Client, Request)
    If Parameters <> "" Then
        base = Parameters & "&"
    End If
    
    base = base & "oauth_consumer_key" & "=" & Me.ConsumerKey
    base = base & "&" & "oauth_nonce" & "=" & Nonce
    base = base & "&" & "oauth_signature_method" & "=" & SignatureMethod
    base = base & "&" & "oauth_timestamp" & "=" & Timestamp
    base = base & "&" & "oauth_token" & "=" & Me.Token
    base = base & "&" & "oauth_version=1.0"
    
    CreateBaseString = WebHelpers.MethodToName(Request.Method) & "&" & WebHelpers.UrlEncode(RequestUrl(Client, Request)) & "&" & WebHelpers.UrlEncode(base)
End Function

Public Function RequestUrl(Client As WebClient, Request As WebRequest) As String
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

    Dim Parts As Dictionary
    Set Parts = WebHelpers.UrlParts(Client.GetFullRequestUrl(Request))
    
    ' Url scheme and authority MUST be lowercase
    RequestUrl = LCase(Parts("Protocol") & "://" & Parts("Host"))
    
    ' Include port (80 and 443 MUST be excluded)
    If Parts("Port") <> 80 And Parts("Port") <> 443 Then
        RequestUrl = RequestUrl & ":" & Parts("Port")
    End If
    
    ' Include path
    RequestUrl = RequestUrl + Parts("Path")
    
    ' MUST exclude query and fragment
End Function

Public Function RequestParameters(Client As WebClient, Request As WebRequest) As String
' TODO Sort parameters by key then value
    
    Dim Parts As Dictionary
    Set Parts = WebHelpers.UrlParts(Client.GetFullRequestUrl(Request))
    
    ' Remove leading ?
    RequestParameters = Parts("Querystring")
    
    ' Replace + for spaces with %20
    RequestParameters = Replace(RequestParameters, "+", "%20")
End Function

Public Function CreateSigningKey() As String
    CreateSigningKey = Me.ConsumerSecret & "&" & Me.TokenSecret
End Function

Public Function CreateSignature(base As String, signingKey As String) As String
    CreateSignature = WebHelpers.HMACSHA1(base, signingKey, "Base64")
End Function

Public Function CreateTimestamp() As String
    CreateTimestamp = CStr(DateDiff("s", #1/1/1970#, GetGMT()))
End Function

' ======================================================================================== '
'
' Time Calculations
'
' ======================================================================================== '

Function GetGMT(Optional LocalTime As Date) As Date
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' http://www.cpearson.com/excel/TimeZoneAndDaylightTime.aspx
    ' ConvertLocalToGMT
    ' This function returns the GMT based on LocalTime, if provided.
    ' If LocalTime is not equal to 0, the GMT corresponding to LocalTime
    ' is returned. If LocalTime is 0, the GMT corresponding to the local
    ' time is returned. Since GMT isn't affected by DST, we need to
    ' subtract 1 hour if we are presently in DST.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim T As Date
    Dim TZI As TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    Dim GMT As Date
    
    If LocalTime <= 0 Then
        T = Now
    Else
        T = LocalTime
    End If
    DST = GetTimeZoneInformation(TZI)
    GMT = T + TimeSerial(0, TZI.Bias, 0) - IIf(DST = TIME_ZONE_DAYLIGHT, TimeSerial(1, 0, 0), 0)
    GetGMT = GMT
End Function
