Attribute VB_Name = "RestHelpers"
''
' RestHelpers v2.1.2
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Common helpers RestClient
'
' @dependencies: Microsoft Scripting Runtime
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' Declare SetTimer and KillTimer
' See [SetTimer and VBA](http://www.mcpher.com/Home/excelquirks/classeslink/vbapromises/timercallbacks)
' and [MSDN Article](http://msdn.microsoft.com/en-us/library/windows/desktop/ms644906(v=vs.85).aspx)
' --------------------------------------------- '
#If VBA7 And Win64 Then
    ' 64-bit
    Public Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal HWnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
    Public Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long) As Long
   
#Else
    '32-bit
    Public Declare Function SetTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long) As Long
  
#End If

Private Const UserAgent As String = "Excel Client v2.1.2 (https://github.com/timhall/Excel-REST)"

' Moved to top from JSONLib
Private Const INVALID_JSON      As Long = 1
Private Const INVALID_OBJECT    As Long = 2
Private Const INVALID_ARRAY     As Long = 3
Private Const INVALID_BOOLEAN   As Long = 4
Private Const INVALID_NULL      As Long = 5
Private Const INVALID_KEY       As Long = 6

Public Enum StatusCodes
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
    RequestTimeout = 408
    UnsupportedMediaType = 415
    InternalServerError = 500
    BadGateway = 502
    ServiceUnavailable = 503
    GatewayTimeout = 504
End Enum


' ============================================= '
' Shared Helpers
' ============================================= '

''
' Parse given JSON string into object (Dictionary or Collection)
'
' @param {String} JSON
' @return {Object}
' --------------------------------------------- '

Public Function ParseJSON(json As String) As Object
    Set ParseJSON = json_parse(json)
End Function

''
' Convert object to JSON string
'
' @param {Object}
' @return {String}
' --------------------------------------------- '

Public Function ConvertToJSON(Obj As Object) As String
    ConvertToJSON = json_toString(Obj)
End Function

''
' URL Encode the given raw values
'
' @param {Variant} rawVal The raw string to encode
' @param {Boolean} [spaceAsPlus=False] Use plus sign for encoded spaces (otherwise %20)
' @return {String} Encoded string
' --------------------------------------------- '

Public Function URLEncode(rawVal As Variant, Optional spaceAsPlus As Boolean = False) As String
    Dim urlVal As String
    Dim stringLen As Long
    
    urlVal = CStr(rawVal)
    stringLen = Len(urlVal)
    
    If stringLen > 0 Then
        ReDim Result(stringLen) As String
        Dim i As Long, charCode As Integer
        Dim char As String, space As String
        
        ' Set space value
        If spaceAsPlus Then
            space = "+"
        Else
            space = "%20"
        End If
        
        ' Loop through string characters
        For i = 1 To stringLen
            ' Get character and ascii code
            char = Mid$(urlVal, i, 1)
            charCode = asc(char)
            Select Case charCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    ' Use original for AZaz09-._~
                    Result(i) = char
                Case 32
                    ' Add space
                    Result(i) = space
                Case 0 To 15
                    ' Convert to hex w/ leading 0
                    Result(i) = "%0" & Hex(charCode)
                Case Else
                    ' Convert to hex
                    Result(i) = "%" & Hex(charCode)
            End Select
        Next i
        URLEncode = Join(Result, "")
    End If
End Function


''
' Join Url with /
'
' @param {String} LeftSide
' @param {String} RightSide
' @return {String} Joined url
' --------------------------------------------- '

Public Function JoinUrl(LeftSide As String, RightSide As String) As String
    If Left(RightSide, 1) = "/" Then
        RightSide = Right(RightSide, Len(RightSide) - 1)
    End If
    If Right(LeftSide, 1) = "/" Then
        LeftSide = Left(LeftSide, Len(LeftSide) - 1)
    End If
    
    JoinUrl = LeftSide & "/" & RightSide
End Function

''
' Combine two objects
'
' @param {Dictionary} OriginalObj Original object to add values to
' @param {Dictionary} NewObj New object containing values to add to original object
' @param {Boolean} [OverwriteOriginal=True] Overwrite any values that already exist in the original object
' @return {Dictionary} Combined object
' --------------------------------------------- '

Public Function CombineObjects(ByVal OriginalObj As Dictionary, ByVal NewObj As Dictionary, _
    Optional OverwriteOriginal As Boolean = True) As Dictionary
    
    Dim Combined As New Dictionary
    
    Dim OriginalKey As Variant
    Dim Key As Variant
    
    If Not OriginalObj Is Nothing Then
        For Each Key In OriginalObj.keys()
            Combined.Add Key, OriginalObj(Key)
        Next Key
    End If
    If Not NewObj Is Nothing Then
        For Each Key In NewObj.keys()
            If Combined.Exists(Key) And OverwriteOriginal Then
                Combined(Key) = NewObj(Key)
            ElseIf Not Combined.Exists(Key) Then
                Combined.Add Key, NewObj(Key)
            End If
        Next Key
    End If
    
    Set CombineObjects = Combined
End Function

''
' Apply whitelist to given object to filter out unwanted key/values
'
' @param {Dictionary} Original model to filter
' @param {Variant} WhiteList Array of values to retain in the model
' @return {Dictionary} Filtered object
' --------------------------------------------- '

Public Function FilterObject(ByVal Original As Dictionary, Whitelist As Variant) As Dictionary
    Dim Filtered As New Dictionary
    Dim i As Integer
    
    If IsArray(Whitelist) Then
        For i = LBound(Whitelist) To UBound(Whitelist)
            If Original.Exists(Whitelist(i)) Then
                Filtered.Add Whitelist(i), Original(Whitelist(i))
            End If
        Next i
    ElseIf VarType(Whitelist) = vbString Then
        If Original.Exists(Whitelist) Then
            Filtered.Add Whitelist, Original(Whitelist)
        End If
    End If
    
    Set FilterObject = Filtered
End Function

''
' Prepare http request for execution
'
' @param {RestRequest} Request
' @param {Integer} TimeoutMS
' @param {Boolean} [UseAsync=False]
' @return {Object} Setup http object
' --------------------------------------------- '

Public Function PrepareHttpRequest(Request As RestRequest, TimeoutMS As Integer, _
    Optional UseAsync As Boolean = False) As Object
    Dim Http As Object
    Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Set timeouts
    Http.setTimeouts TimeoutMS, TimeoutMS, TimeoutMS, TimeoutMS
    
    ' Add general headers to request
    Request.AddHeader "User-Agent", UserAgent
    Request.AddHeader "Content-Type", Request.ContentType
    
    If Request.IncludeContentLength Then
        Request.AddHeader "Content-Length", Request.ContentLength
    Else
        If Request.Headers.Exists("Content-Length") Then
            Request.Headers.Remove "Content-Length"
        End If
    End If
    
    ' Pass http to request and setup onreadystatechange
    If UseAsync Then
        Set Request.HttpRequest = Http
        Http.onreadystatechange = Request
    End If
    
    Set PrepareHttpRequest = Http
End Function

''
' Set headers to http object for given request
'
' @param {Object} Http request
' @param {RestRequest} Request
' --------------------------------------------- '

Public Sub SetHeaders(ByRef Http As Object, Request As RestRequest)
    Dim HeaderKey As Variant
    For Each HeaderKey In Request.Headers.keys()
        Http.setRequestHeader HeaderKey, Request.Headers(HeaderKey)
    Next HeaderKey
End Sub

''
' Prepare proxy for http object
'
' @param {String} ProxyServer
' @param {String} [Username=""]
' @param {String} [Password=""]
' @param {Variant} [BypassList]
' --------------------------------------------- '

Public Sub PrepareProxyForHttpRequest(ByRef Http As Object, ProxyServer As String, _
    Optional Username As String = "", Optional Password As String = "", Optional BypassList As Variant)
    
    If ProxyServer <> "" Then
        Http.SetProxy 2, ProxyServer, BypassList
        
        If Username <> "" Then
            Http.SetProxyCredentials Username, Password
        End If
    End If
End Sub

''
' Execute request synchronously
'
' @param {Object} Http
' @param {RestRequest} Request The request to execute
' @return {RestResponse} Wrapper of server response for request
' --------------------------------------------- '

Public Function ExecuteRequest(ByRef Http As Object, ByRef Request As RestRequest) As RestResponse
    On Error GoTo ErrorHandling
    Dim Response As RestResponse

    ' Send the request and handle response
    Http.Send Request.Body
    Set Response = Request.CreateResponseFromHttp(Http)
    
ErrorHandling:

    If Not Http Is Nothing Then Set Http = Nothing
    If Err.Number <> 0 Then
        If InStr(Err.Description, "The operation timed out") > 0 Then
            ' Return 408
            Set Response = Request.CreateResponse(StatusCodes.RequestTimeout, "Request Timeout")
            Err.Clear
        Else
            ' Rethrow error
            Err.Raise Err.Number, Description:=Err.Description
        End If
    End If
    
    Set ExecuteRequest = Response
End Function

''
' Execute request asynchronously
'
' @param {Object} Http
' @param {RestRequest} Request The request to execute
' @param {String} Callback Name of function to call when request completes (specify "" if none)
' @param {Variant} [CallbackArgs] Variable array of arguments that get passed directly to callback function
' --------------------------------------------- '

Public Sub ExecuteRequestAsync(ByRef Http As Object, ByRef Request As RestRequest, TimeoutMS As Integer, Callback As String, Optional ByVal CallbackArgs As Variant)
    On Error GoTo ErrorHandling

    Request.Callback = Callback
    Request.CallbackArgs = CallbackArgs
    
    ' Send the request
    Request.StartTimeoutTimer TimeoutMS
    Http.Send Request.Body
    
    Exit Sub
    
ErrorHandling:

    ' Close http and rethrow error
    If Not Http Is Nothing Then Set Http = Nothing
    Err.Raise Err.Number, Description:=Err.Description
End Sub

' ======================================================================================== '
'
' Timeout Timing
'
' ======================================================================================== '

''
' Start timeout timer for request
'
' @param {RestRequest} Request
' @param {Long} TimeoutMS
' --------------------------------------------- '
Public Sub StartTimeoutTimer(Request As RestRequest, TimeoutMS As Integer)
    SetTimer Application.HWnd, ObjPtr(Request), TimeoutMS, AddressOf RestHelpers.TimeoutTimerExpired
End Sub

''
' Stop timeout timer for request
'
' @param {RestRequest} Request
' --------------------------------------------- '
Public Sub StopTimeoutTimer(Request As RestRequest)
    KillTimer Application.HWnd, ObjPtr(Request)
End Sub

''
' Handle timeout timers expiring
'
' See [MSDN Article](http://msdn.microsoft.com/en-us/library/windows/desktop/ms644907(v=vs.85).aspx)
' --------------------------------------------- '
#If VBA7 And Win64 Then
Public Sub TimeoutTimerExpired(ByVal HWnd As Long, ByVal Msg As Long, _
        ByVal Request As RestRequest, ByVal dwTimer As Long)
#Else
Sub TimeoutTimerExpired(ByVal HWnd As Long, ByVal uMsg As Long, _
        ByVal Request As RestRequest, ByVal dwTimer As Long)
#End If
    
    StopTimeoutTimer Request
    Request.TimedOut
End Sub

' ======================================================================================== '
'
' Crytography and encoding
'
' ======================================================================================== '

''
' Generate a keyed hash value using the HMAC method and SHA1 algorithm
' [Does VBA have a Hash_HMAC](http://stackoverflow.com/questions/8246340/does-vba-have-a-hash-hmac)
'
' @param {String} sTextToHash
' @param {String} sSharedSecretKey
' @return {String}
' --------------------------------------------- '

Public Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey As String) As String
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = EncodeBase64(bytes)
    Set asc = Nothing
    Set enc = Nothing
End Function

''
' Base64 encode data
'
' @param {Byte Array} arrData
' @return {String} Encoded string
' --------------------------------------------- '

Public Function EncodeBase64(ByRef Data() As Byte) As String
    Dim XML As Object
    Dim Node As Object
    Set XML = CreateObject("MSXML2.DOMDocument")

    ' byte array to base64
    Set Node = XML.createElement("b64")
    Node.DataType = "bin.base64"
    Node.nodeTypedValue = Data
    EncodeBase64 = Node.text

    Set Node = Nothing
    Set XML = Nothing
End Function

''
' Base64 encode string value
'
' @param {String} Data
' @return {String} Encoded string
' --------------------------------------------- '

Public Function EncodeStringToBase64(ByVal Data As String) As String
    Dim asc As Object
    Dim bytes() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    bytes = asc.Getbytes_4(Data)
    EncodeStringToBase64 = Replace(EncodeBase64(bytes), vbLf, "")
    Set asc = Nothing
End Function

''
' Create random alphanumeric nonce
'
' @param {Integer} [NonceLength=32]
' @return {String} Randomly generated nonce
' --------------------------------------------- '

Public Function CreateNonce(Optional NonceLength As Integer = 32) As String
    Dim str As String
    Dim count As Integer
    Dim Result As String
    Dim random As Integer
    
    str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    Result = ""
    
    For count = 1 To NonceLength
        random = Int(((Len(str) - 1) * Rnd) + 1)
        Result = Result + Mid$(str, random, 1)
    Next
    CreateNonce = Result
End Function


' ======================================================================================== '
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' Changes for Excel-REST:
' - Updated json_parseNumber to reduce chance of overflow
' - Swapped Mid for Mid$
' - Handle colon in object key
' - Handle duplicate keys in object parsing
' - Change methods to Private and prefix with json_
'
' ======================================================================================== '

' (Moved to top of file)
'Private Const INVALID_JSON      As Long = 1
'Private Const INVALID_OBJECT    As Long = 2
'Private Const INVALID_ARRAY     As Long = 3
'Private Const INVALID_BOOLEAN   As Long = 4
'Private Const INVALID_NULL      As Long = 5
'Private Const INVALID_KEY       As Long = 6

'
'   parse string and create JSON object (Dictionary or Collection in VB)
'
Private Function json_parse(ByRef str As String) As Object

    Dim index As Long
    index = 1
    
    On Error Resume Next

    Call json_skipChar(str, index)
    Select Case Mid$(str, index, 1)
    Case "{"
        Set json_parse = json_parseObject(str, index)
    Case "["
        Set json_parse = json_parseArray(str, index)
    End Select

End Function

'
'   parse collection of key/value (Dictionary in VB)
'
Private Function json_parseObject(ByRef str As String, ByRef index As Long) As Dictionary

    Set json_parseObject = New Dictionary
    
    ' "{"
    Call json_skipChar(str, index)
    If Mid$(str, index, 1) <> "{" Then Err.Raise vbObjectError + INVALID_OBJECT, Description:="char " & index & " : " & Mid$(str, index)
    index = index + 1
    
    Dim Key As String
    
    Do
        Call json_skipChar(str, index)
        If "}" = Mid$(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid$(str, index, 1) Then
            index = index + 1
            Call json_skipChar(str, index)
        End If
        
        Key = json_parseKey(str, index)
        If Not json_parseObject.Exists(Key) Then
            json_parseObject.Add Key, json_parseValue(str, index)
        Else
            json_parseObject.Item(Key) = json_parseValue(str, index)
        End If
    Loop

End Function

'
'   parse list (Collection in VB)
'
Private Function json_parseArray(ByRef str As String, ByRef index As Long) As Collection

    Set json_parseArray = New Collection
    
    ' "["
    Call json_skipChar(str, index)
    If Mid$(str, index, 1) <> "[" Then Err.Raise vbObjectError + INVALID_ARRAY, Description:="char " & index & " : " + Mid$(str, index)
    index = index + 1
    
    Do
        
        Call json_skipChar(str, index)
        If "]" = Mid$(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid$(str, index, 1) Then
            index = index + 1
            Call json_skipChar(str, index)
        End If
        
        ' add value
        json_parseArray.Add json_parseValue(str, index)
        
    Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function json_parseValue(ByRef str As String, ByRef index As Long)

    Call json_skipChar(str, index)
    
    Select Case Mid$(str, index, 1)
    Case "{"
        Set json_parseValue = json_parseObject(str, index)
    Case "["
        Set json_parseValue = json_parseArray(str, index)
    Case """", "'"
        json_parseValue = json_parseString(str, index)
    Case "t", "f"
        json_parseValue = json_parseBoolean(str, index)
    Case "n"
        json_parseValue = json_parseNull(str, index)
    Case Else
        json_parseValue = json_parseNumber(str, index)
    End Select

End Function

'
'   parse string
'
Private Function json_parseString(ByRef str As String, ByRef index As Long) As String

    Dim quote   As String
    Dim char    As String
    Dim Code    As String
    
    Call json_skipChar(str, index)
    quote = Mid$(str, index, 1)
    index = index + 1
    Do While index > 0 And index <= Len(str)
        char = Mid$(str, index, 1)
        Select Case (char)
        Case "\"
            index = index + 1
            char = Mid$(str, index, 1)
            Select Case (char)
            Case """", "\", "/" ' Before: Case """", "\\", "/"
                json_parseString = json_parseString & char
                index = index + 1
            Case "b"
                json_parseString = json_parseString & vbBack
                index = index + 1
            Case "f"
                json_parseString = json_parseString & vbFormFeed
                index = index + 1
            Case "n"
                json_parseString = json_parseString & vbNewLine
                index = index + 1
            Case "r"
                json_parseString = json_parseString & vbCr
                index = index + 1
            Case "t"
                json_parseString = json_parseString & vbTab
                index = index + 1
            Case "u"
                index = index + 1
                Code = Mid$(str, index, 4)
                json_parseString = json_parseString & ChrW(Val("&h" + Code))
                index = index + 4
            End Select
        Case quote
            
            index = index + 1
            Exit Function
        Case Else
            json_parseString = json_parseString & char
            index = index + 1
        End Select
    Loop

End Function

'
'   parse number
'
Private Function json_parseNumber(ByRef str As String, ByRef index As Long)

    Dim Value   As String
    Dim char    As String
    
    Call json_skipChar(str, index)
    Do While index > 0 And index <= Len(str)
        char = Mid$(str, index, 1)
        If InStr("+-0123456789.eE", char) Then
            Value = Value & char
            index = index + 1
        Else
            json_parseNumber = Val(Value)
            Exit Function
        End If
    Loop


End Function

'
'   parse true / false
'
Private Function json_parseBoolean(ByRef str As String, ByRef index As Long) As Boolean

    Call json_skipChar(str, index)
    If Mid$(str, index, 4) = "true" Then
        json_parseBoolean = True
        index = index + 4
    ElseIf Mid$(str, index, 5) = "false" Then
        json_parseBoolean = False
        index = index + 5
    Else
        Err.Raise vbObjectError + INVALID_BOOLEAN, Description:="char " & index & " : " & Mid$(str, index)
    End If

End Function

'
'   parse null
'
Private Function json_parseNull(ByRef str As String, ByRef index As Long)

    Call json_skipChar(str, index)
    If Mid$(str, index, 4) = "null" Then
        json_parseNull = Null
        index = index + 4
    Else
        Err.Raise vbObjectError + INVALID_NULL, Description:="char " & index & " : " & Mid$(str, index)
    End If

End Function

Private Function json_parseKey(ByRef str As String, ByRef index As Long) As String

    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim char    As String
    
    Call json_skipChar(str, index)
    Do While index > 0 And index <= Len(str)
        char = Mid$(str, index, 1)
        Select Case (char)
        Case """"
            dquote = Not dquote
            index = index + 1
            If Not dquote Then
                Call json_skipChar(str, index)
                If Mid$(str, index, 1) <> ":" Then
                    Err.Raise vbObjectError + INVALID_KEY, Description:="char " & index & " : " & json_parseKey
                End If
            End If
        Case "'"
            squote = Not squote
            index = index + 1
            If Not squote Then
                Call json_skipChar(str, index)
                If Mid$(str, index, 1) <> ":" Then
                    Err.Raise vbObjectError + INVALID_KEY, Description:="char " & index & " : " & json_parseKey
                End If
            End If
        Case ":"
            If Not dquote And Not squote Then
                index = index + 1
                Exit Do
            Else
                ' Colon in key name
                json_parseKey = json_parseKey & char
                index = index + 1
            End If
        Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", char) Then
            Else
                json_parseKey = json_parseKey & char
            End If
            index = index + 1
        End Select
    Loop

End Function

'
'   skip special character
'
Private Sub json_skipChar(ByRef str As String, ByRef index As Long)

    While index > 0 And index <= Len(str) And InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Mid$(str, index, 1))
        index = index + 1
    Wend

End Sub

Private Function json_toString(ByRef Obj As Variant) As String

    Select Case VarType(Obj)
        Case vbNull
            json_toString = "null"
        Case vbEmpty
            json_toString = """"""
        Case vbDate
            json_toString = """" & CStr(Obj) & """"
        Case vbString
            json_toString = """" & json_encode(Obj) & """"
        Case vbObject
            Dim bFI, i
            bFI = True
            If TypeName(Obj) = "Dictionary" Then
                json_toString = json_toString & "{"
                Dim keys
                keys = Obj.keys
                For i = 0 To Obj.count - 1
                    If bFI Then bFI = False Else json_toString = json_toString & ","
                    Dim Key
                    Key = keys(i)
                    json_toString = json_toString & """" & Key & """:" & json_toString(Obj(Key))
                Next i
                json_toString = json_toString & "}"
            ElseIf TypeName(Obj) = "Collection" Then
                json_toString = json_toString & "["
                Dim Value
                For Each Value In Obj
                    If bFI Then bFI = False Else json_toString = json_toString & ","
                    json_toString = json_toString & json_toString(Value)
                Next Value
                json_toString = json_toString & "]"
            End If
        Case vbBoolean
            If Obj Then json_toString = "true" Else json_toString = "false"
        Case vbVariant, vbArray, vbArray + vbVariant
            Dim sEB
            json_toString = json_multiArray(Obj, 1, "", sEB)
        Case Else
            json_toString = Replace(Obj, ",", ".")
    End Select

End Function

Private Function json_encode(str) As String
    
    Dim i, j, aL1, aL2, C, p

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To Len(str)
        p = True
        C = Mid$(str, i, 1)
        For j = 0 To 7
            If C = Chr(aL1(j)) Then
                json_encode = json_encode & "\" & Chr(aL2(j))
                p = False
                Exit For
            End If
        Next

        If p Then
            Dim A
            A = AscW(C)
            If A > 31 And A < 127 Then
                json_encode = json_encode & C
            ElseIf A > -1 Or A < 65535 Then
                json_encode = json_encode & "\u" & String(4 - Len(Hex(A)), "0") & Hex(A)
            End If
        End If
    Next
End Function

Private Function json_multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
    Dim iDU, iDL, i ' Integer DimensionUBound, Integer DimensionLBound
    On Error Resume Next
    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)
    
    Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
    If Err.Number = 9 Then
        sPB1 = sPT & sPS
        For i = 1 To Len(sPB1)
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid$(sPB1, i, 1)
        Next
'        json_multiArray = json_multiArray & json_toString(Eval("aBD(" & sPB2 & ")"))
        json_multiArray = json_multiArray & json_toString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        json_multiArray = json_multiArray & "["
        For i = iDL To iDU
            json_multiArray = json_multiArray & json_multiArray(aBD, iBC + 1, i, sPT)
            If i < iDU Then json_multiArray = json_multiArray & ","
        Next
        json_multiArray = json_multiArray & "]"
        sPT = Left(sPT, iBC - 2)
    End If
    Err.Clear
End Function

