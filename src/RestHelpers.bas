Attribute VB_Name = "RestHelpers"
''
' RestHelpers v1.0.3
' (c) Tim Hall - https://github.com/timhall/Excel-REST
'
' Common helpers RestClient
'
' @dependencies
'   JSONLib (http://code.google.com/p/vba-json/)
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Enum StatusCodes
    Ok = 200
    Created = 201
    NoContent = 204
    NotModified = 304
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
    NotFound = 404
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

Public Function ParseJSON(JSON As String) As Object
    Dim lib As New JSONLib
    Set ParseJSON = lib.parse(JSON)
    Set lib = Nothing
End Function

''
' Convert object to JSON string
'
' @param {Object}
' @return {String}
' --------------------------------------------- '

Public Function ConvertToJSON(Obj As Object) As String
    Dim lib As New JSONLib
    ConvertToJSON = lib.ToString(Obj)
    Set lib = Nothing
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
        ReDim result(stringLen) As String
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
                    result(i) = char
                Case 32
                    ' Add space
                    result(i) = space
                Case 0 To 15
                    ' Convert to hex w/ leading 0
                    result(i) = "%0" & Hex(charCode)
                Case Else
                    ' Convert to hex
                    result(i) = "%" & Hex(charCode)
            End Select
        Next i
        URLEncode = Join(result, "")
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
' @param {Dictionary} origObj Original object to add values to
' @param {Dictionary} newObj New object containing values to add to original object
' @param {Boolean} [overwriteOriginal=True] Overwrite any values that already exist in the original object
' --------------------------------------------- '

Public Function CombineObjects(ByVal OriginalObj As Object, ByVal NewObj As Object, _
    Optional OverwriteOriginal As Boolean = True) As Object
    
    Dim Combined As Object
    Dim NewKey As Variant
    
    If Not OriginalObj Is Nothing Then
        Set Combined = OriginalObj
    Else
        Set Combined = CreateObject("Scripting.Dictionary")
    End If
    For Each NewKey In NewObj.keys()
        If Combined.Exists(NewKey) And OverwriteOriginal Then
            Combined(NewKey) = NewObj(NewKey)
        Else
            Combined.Add NewKey, NewObj(NewKey)
        End If
    Next NewKey

    Set CombineObjects = Combined
End Function

''
' Apply whitelist to given model to filter out unwanted key/values
'
' @param {Dictionary} Original Original model to filter
' @param {Variant} WhiteList Array of values to retain in the model
' --------------------------------------------- '

Public Function FilterModel(ByVal Original As Object, WhiteList As Variant) As Object
    Dim Filtered As Object
    Dim i As Integer
    
    Set Filtered = CreateObject("Scripting.Dictionary")
    
    If IsArray(WhiteList) Then
        For i = LBound(WhiteList) To UBound(WhiteList)
            If Original.Exists(WhiteList(i)) Then
                Filtered.Add WhiteList(i), Original(WhiteList(i))
            End If
        Next i
    ElseIf VarType(WhiteList) = vbString Then
        If Original.Exists(WhiteList) Then
            Filtered.Add WhiteList, Original(WhiteList)
        End If
    End If
    
    Set FilterModel = Filtered
End Function

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
    Dim result As String
    Dim random As Integer
    
    str = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUIVWXYZ"
    result = ""
    
    For count = 1 To NonceLength
        random = Int(((Len(str) - 1) * Rnd) + 1)
        result = result + Mid$(str, random, 1)
    Next
    CreateNonce = result
End Function


