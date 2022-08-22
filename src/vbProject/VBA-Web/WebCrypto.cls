VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' WebCrypto
'
' A class that calls Windows Crypto APIs to create HMAC-SHA1, HMAC-SHA256 signatures & SHA256, MD5 hashes.
'
' Several useful encoding and decoding operations are included, along
' with conversion to and from UTF8 character encoding.
'
' Based on code shared by Bob Riemersma: https://www.vbforums.com/showthread.php?635398-VB6-HMAC-SHA-256-HMAC-SHA-1-Using-Crypto-API/
' Modified to work with Tim Hall's VBA-Web: https://github.com/VBA-tools/VBA-Web
'
' @module WebCrypto
' @author Bob Riemersma
' @edited Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder VBA-Web
'@ignoremodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Windows API Headers
' --------------------------------------------- '

Private Const CP_UTF8 As Long = 65001

Private Const CALG_RC2                   As Long = &H6602&
Private Const CALG_MD5                   As Long = &H8003&
Private Const CALG_SHA1                  As Long = &H8004&
Private Const CALG_SHA_256               As Long = &H800C&
Private Const CALG_HMAC                  As Long = &H8009&

Private Const PROV_RSA_FULL              As Long = 1
Private Const PROV_RSA_AES               As Long = 24
Private Const CRYPT_VERIFYCONTEXT        As Long = &HF0000000
Private Const CRYPT_MACHINE_KEYSET       As Long = 32
Private Const MS_DEFAULT_PROVIDER        As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENH_RSA_AES_PROV        As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"
Private Const MS_ENH_RSA_AES_PROV_XP     As String = "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"

Private Const HP_HASHVAL                 As Long = 2
Private Const HP_HASHSIZE                As Long = 4
Private Const HP_HMAC_INFO               As Long = 5

Private Const CRYPT_STRING_BASE64        As Long = &H1&
Private Const CRYPT_STRING_HEX           As Long = &H4&
Private Const CRYPT_STRING_HEXASCII      As Long = &H5&
Private Const CRYPT_STRING_HEXADDR       As Long = &HA&
Private Const CRYPT_STRING_HEXASCIIADDR  As Long = &HB&
Private Const CRYPT_STRING_HEXRAW        As Long = &HC&       'Requires Vista or later, so we emulate.
Private Const CRYPT_STRING_NOCR          As Long = &H80000000
Private Const CRYPT_STRING_NOCRLF        As Long = &H40000000 'Requires Vista or later!

Private Const CRYPT_IPSEC_HMAC_KEY       As Long = &H100&

Private Const PLAINTEXTKEYBLOB           As Byte = &H8
Private Const CUR_BLOB_VERSION           As Byte = &H2

Private Type HMAC_INFO
    HashAlgId As Long
    pbInnerString As Long
    cbInnerString As Long
    pbOuterString As Long
    cbOuterString As Long
End Type

Private Type BLOBHEADER
    bType As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
End Type

Private Type KEYBLOB
    hdr As BLOBHEADER
    cbKeySize As Long
    'rgbKeyData() As Byte 'We'll actually append this when we build the Byte array copy.
End Type

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Any, _
    ByVal Source As Any, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function CryptAcquireContext Lib "Advapi32" Alias "CryptAcquireContextW" ( _
    ByRef phProv As LongPtr, _
    ByVal pszContainer As LongPtr, _
    ByVal pszProvider As LongPtr, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptCreateHash Lib "Advapi32" ( _
    ByVal hProv As LongPtr, _
    ByVal AlgId As Long, _
    ByVal hKey As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef phHash As LongPtr) As Long
    
Private Declare PtrSafe Function CryptDestroyHash Lib "Advapi32" ( _
    ByVal hHash As LongPtr) As Long
    
Private Declare PtrSafe Function CryptDestroyKey Lib "Advapi32" ( _
    ByVal hKey As LongPtr) As Long

Private Declare PtrSafe Function CryptGetHashParam Lib "Advapi32" ( _
    ByVal hHash As LongPtr, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptHashData Lib "Advapi32" ( _
    ByVal hHash As LongPtr, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptImportKey Lib "Advapi32" ( _
    ByVal hProv As LongPtr, _
    ByVal pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal hPubKey As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef phKey As LongPtr) As Long

Private Declare PtrSafe Function CryptReleaseContext Lib "Advapi32" ( _
    ByVal hProv As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptSetHashParam Lib "Advapi32" ( _
    ByVal hHash As LongPtr, _
    ByVal dwParam As Long, _
    ByRef pbData As HMAC_INFO, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptBinaryToString Lib "Crypt32" _
    Alias "CryptBinaryToStringW" ( _
    ByRef pbBinary As Byte, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As LongPtr, _
    ByRef pcchString As Long) As Long
    
Private Declare PtrSafe Function CryptStringToBinary Lib "Crypt32" _
    Alias "CryptStringToBinaryW" ( _
    ByVal pszString As LongPtr, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, _
    ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, _
    ByRef pdwFlags As Long) As Long

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private hBaseProvider As LongPtr
Private hAdvProvider As LongPtr
Private hKey As LongPtr
Private hHmacHash As LongPtr
Private TypeNameOfMe As String
Private IsWin5_1 As Boolean 'XP or Windows 2003 Server.

' --------------------------------------------- '
' Enums
' --------------------------------------------- '

''
' @property EncDecFormat
' @param edfBase64          Standard Base64 encoding, with line-wrap, no headers.
' @param edfHex             Hexadecimal encoding with line-wrap.
' @param edfHexAscii        Hex plus ASCII encoding with line-wrap.
' @param edfHexAddr         Hex plus an address/count at the left of each wrapped line (with line-wrap).
' @param edfHexAsciiAddr    Hex plus ASCII plus an address/count, with line-wrap.
' @param edfHexRaw          Hexadecimal encoding with NO line-wrap.
''
Public Enum EncDecFormat
    edfBase64
    edfHex
    edfHexAscii
    edfHexAddr
    edfHexAsciiAddr
    edfHexRaw
End Enum

''
' @property EncodeFolding
' @param efCrLf             Standard CrLf line-wrap.
' @param efLf               Lf-only line-wrap.
' @param efNoFolding        Suppress folding (requires Vista or later).
''
Public Enum EncodeFolding
    efCrLf
    efLf
    efNoFolding
End Enum

' ============================================= '
' Public Methods
' ============================================= '

''
' Accepts a String of data in any of several supported formats
' and returns a decoded binary Byte array result.
'
' @method Decode
' @param {String} Encoded
' @param {EncDecFormat} [Format=edfHexRaw]
' @return {Byte}
''
Public Function Decode(ByVal Encoded As String, Optional ByVal Format As EncDecFormat = edfHexRaw) As Byte()
    
    Dim dwFlags As Long
    Dim bytBuf() As Byte
    Dim lngOutLen As Long
    Dim dwActualUsed As Long
    
    Select Case Format
        Case edfBase64
            dwFlags = CRYPT_STRING_BASE64
        Case edfHex
            dwFlags = CRYPT_STRING_HEX
        Case edfHexAscii
            dwFlags = CRYPT_STRING_HEXASCII
        Case edfHexAddr
            dwFlags = CRYPT_STRING_HEXADDR
        Case edfHexAsciiAddr
            dwFlags = CRYPT_STRING_HEXASCIIADDR
        Case edfHexRaw
            dwFlags = CRYPT_STRING_HEXRAW
    End Select
    
    If IsWin5_1 And (Format = edfHexRaw) Then
        'Emulate missing format.
        Dim i As Long
        
        ReDim bytBuf(Len(Encoded) \ 2 - 1)
        For i = 1 To Len(Encoded) Step 2
            bytBuf((i - 1) \ 2) = VBA.CLng("&H0" & VBA.Mid$(Encoded, i, 2))
        Next
        Decode = bytBuf
    ElseIf CryptStringToBinary(StrPtr(Encoded), _
                              VBA.Len(Encoded), _
                              dwFlags, _
                              0, _
                              lngOutLen, _
                              0, _
                              dwActualUsed) = 0 Then
        Err.Raise vbObjectError Or &HC302&, _
                  TypeNameOfMe, _
                  "Failed to determine decoded length, system error " _
                & VBA.CStr(Err.LastDllError)
    Else
        ReDim bytBuf(lngOutLen - 1)
        If CryptStringToBinary(StrPtr(Encoded), _
                               VBA.Len(Encoded), _
                               dwFlags, _
                               VarPtr(bytBuf(0)), _
                               lngOutLen, _
                               0, _
                               dwActualUsed) = 0 Then
            Err.Raise vbObjectError Or &HC304&, _
                      TypeNameOfMe, _
                      "Failed to decode value, system error " _
                    & VBA.CStr(Err.LastDllError)
        Else
            Decode = bytBuf
        End If
    End If
End Function

''
' Accepts a Byte array of binary data and returns a String encoded in
' any of several formats with various line folding options (efNoFolding
' requires Vista or later). The efNoFolding option does not apply to the
' edfHexRaw format.
'
' @method Encode
' @param {Byte} Bytes
' @param {EncDecFormat} [Format=edfHexRaw]
' @param {EncodeFolding} [Folding=efCrLf]
' @return {String}
''
Public Function Encode(ByRef Bytes() As Byte, Optional ByVal Format As EncDecFormat = edfHexRaw, Optional ByVal Folding As EncodeFolding = efCrLf) As String
    
    Dim dwFlags As Long
    Dim lngOutLen As Long
    Dim strEncoded As String
    
    Select Case Format
        Case edfBase64
            dwFlags = CRYPT_STRING_BASE64
        Case edfHex
            dwFlags = CRYPT_STRING_HEX
        Case edfHexAscii
            dwFlags = CRYPT_STRING_HEXASCII
        Case edfHexAddr
            dwFlags = CRYPT_STRING_HEXADDR
        Case edfHexAsciiAddr
            dwFlags = CRYPT_STRING_HEXASCIIADDR
        Case edfHexRaw
            dwFlags = CRYPT_STRING_HEXRAW
    End Select
    
    Select Case Folding
        Case efNoFolding
            dwFlags = dwFlags Or CRYPT_STRING_NOCRLF
        Case efLf
            dwFlags = dwFlags Or CRYPT_STRING_NOCR
    End Select

    If IsWin5_1 And (Format = edfHexRaw) Then
        'Emulate missing format.
        Dim i As Long

        strEncoded = VBA.String$(2 * (UBound(Bytes) - LBound(Bytes) + 1), 0)
        For i = 1 To VBA.Len(strEncoded) Step 2
            Mid$(strEncoded, i, 2) = VBA.Right$("0" & VBA.Hex$(Bytes((i - 1) \ 2 + LBound(Bytes))), 2)
        Next
        Encode = strEncoded
    ElseIf CryptBinaryToString(Bytes(LBound(Bytes)), _
                               UBound(Bytes) - LBound(Bytes) + 1, _
                               dwFlags, _
                               0&, _
                               lngOutLen) = 0 Then
        Err.Raise vbObjectError Or &HC30A&, _
                  TypeNameOfMe, _
                  "Failed to determine encoded length, system error " _
                & VBA.CStr(Err.LastDllError)
    Else
        strEncoded = VBA.String$(lngOutLen - 1, 0)
        If CryptBinaryToString(Bytes(LBound(Bytes)), _
                               UBound(Bytes) - LBound(Bytes) + 1, _
                               dwFlags, _
                               StrPtr(strEncoded), _
                               lngOutLen) = 0 Then
            Err.Raise vbObjectError Or &HC30C&, _
                      TypeNameOfMe, _
                      "Failed to encode value, system error " _
                    & VBA.CStr(Err.LastDllError)
        Else
            If IsWin5_1 Then
                Select Case Folding
                    Case efNoFolding
                        Encode = VBA.Replace(strEncoded, vbNewLine, "")
                    Case Else
                        Encode = strEncoded
                End Select
            Else
                Encode = strEncoded
            End If
        End If
    End If
End Function

''
' Accepts a Byte array of UTF8 text and returns a String in UTF16 ("Unicode") encoding.
'
' @method FromUTF8
' @param {Byte} UTF8
' @return {String}
''
Public Function FromUTF8(ByRef UTF8() As Byte) As String
    Dim lngOutLen As Long
    Dim strWide As String
    
    lngOutLen = MultiByteToWideChar(CP_UTF8, _
                                    0, _
                                    VarPtr(UTF8(LBound(UTF8))), _
                                    UBound(UTF8) - LBound(UTF8) + 1, _
                                    0, _
                                    0)
    If lngOutLen = 0 Then
        Err.Raise vbObjectError Or &HC312&, _
                  TypeNameOfMe, _
                  "Failed to decode string, system error " _
                & VBA.CStr(Err.LastDllError)
    Else
        strWide = VBA.String$(lngOutLen, 0)
        lngOutLen = MultiByteToWideChar(CP_UTF8, _
                                        0, _
                                        VarPtr(UTF8(LBound(UTF8))), _
                                        UBound(UTF8) - LBound(UTF8) + 1, _
                                        StrPtr(strWide), _
                                        lngOutLen)
        If lngOutLen = 0 Then
            Err.Raise vbObjectError Or &HC312&, _
                      TypeNameOfMe, _
                      "Failed to decode string, system error " _
                    & VBA.CStr(Err.LastDllError)
        Else
            FromUTF8 = strWide
        End If
    End If
End Function

''
' Accepts a String of UTF16 ("Unicode") text and returns a Byte array in UTF8 encoding.
' @method ToUTF8
' @param {String} Text
' @return {Byte}
''
Public Function ToUTF8(ByVal Text As String) As Byte()
    Dim lngOutLen As Long
    Dim UTF8() As Byte
    
    lngOutLen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), Len(Text), _
                                    0, 0, 0, 0)
    ReDim UTF8(lngOutLen - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(Text), Len(Text), _
                        VarPtr(UTF8(0)), lngOutLen, 0, 0
    ToUTF8 = UTF8
End Function

''
' Creates an HMAC-SHA1 signature using the previously imported key
' and the data provided as a Byte array, and returns a Byte array result.
'
' InitHmac() must be called first, and called again whenever a different key is to be used.
' @method HMACSHA1
' @param {Byte} Data
' @return {Byte}
''
Public Function HMACSHA1(ByRef Data() As Byte) As Byte()
    Dim lngErr As Long
    Dim HmacInfo As HMAC_INFO
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hKey = 0 Then
        Err.Raise vbObjectError Or &HD322&, _
                  "HS1.HmacSha1", _
                  "No key set, call InitHmac first"
    ElseIf CryptCreateHash(hAdvProvider, CALG_HMAC, hKey, 0, hHmacHash) = 0 Then
        lngErr = Err.LastDllError
        DestroyHandles
        Err.Raise vbObjectError Or &HD32A&, _
                  "HS1.HmacSha1", _
                  "Failed to create HMAC hash object, system error " _
                & CStr(lngErr)
    Else
        HmacInfo.HashAlgId = CALG_SHA1
        If CryptSetHashParam(hHmacHash, HP_HMAC_INFO, HmacInfo, 0&) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HD32C&, _
                      "HS1.HmacSha1", _
                      "Failed to set HMAC_INFO hash param, system error " _
                    & CStr(lngErr)
        ElseIf CryptHashData(hHmacHash, _
                             Data(LBound(Data)), _
                             UBound(Data) - LBound(Data) + 1, _
                             0&) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HD32E&, _
                      "HS1.HmacSha1", _
                      "Failed to hash data, system error " _
                    & CStr(lngErr)
        Else
            lngDataLen = 4 '4 bytes for Long length.
            If CryptGetHashParam(hHmacHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
                lngErr = Err.LastDllError
                DestroyHandles
                Err.Raise vbObjectError Or &HD332&, _
                          "HS1.HmacSha1", _
                          "Failed to obtain hash value length, system error " _
                        & CStr(lngErr)
            Else
                lngDataLen = lngHashSize
                ReDim bytHashValue(lngDataLen - 1)
                If CryptGetHashParam(hHmacHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                    lngErr = Err.LastDllError
                    DestroyHandles
                    Err.Raise vbObjectError Or &HD334&, _
                              "HS1.HmacSha1", _
                              "Failed to obtain hash value, system error " _
                            & CStr(lngErr)
                Else
                    DestroyHandles
                    HMACSHA1 = bytHashValue
                End If
            End If
        End If
    End If
End Function

''
' Creates an HMAC-SHA256 signature using the previously imported key
' and the data provided as a Byte array, and returns a Byte array result.
'
' InitHmac() must be called first, and called again whenever a different key is to be used.
'
' @method HMACSHA256
' @param {Byte} Data
' @return {Byte}
''
Public Function HMACSHA256(ByRef Data() As Byte) As Byte()
    Dim lngErr As Long
    Dim HmacInfo As HMAC_INFO
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hKey = 0 Then
        Err.Raise vbObjectError Or &HC322&, _
                  TypeNameOfMe, _
                  "No key set, call InitHmac first"
    ElseIf CryptCreateHash(hAdvProvider, CALG_HMAC, hKey, 0, hHmacHash) = 0 Then
        lngErr = Err.LastDllError
        DestroyHandles
        Err.Raise vbObjectError Or &HC32A&, _
                  TypeNameOfMe, _
                  "Failed to create HMAC hash object, system error " _
                & VBA.CStr(lngErr)
    Else
        HmacInfo.HashAlgId = CALG_SHA_256
        If CryptSetHashParam(hHmacHash, HP_HMAC_INFO, HmacInfo, 0) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HC32C&, _
                      TypeNameOfMe, _
                      "Failed to set HMAC_INFO hash param, system error " _
                    & VBA.CStr(lngErr)
        ElseIf CryptHashData(hHmacHash, _
                             Data(LBound(Data)), _
                             UBound(Data) - LBound(Data) + 1, _
                             0&) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HC32E&, _
                      TypeNameOfMe, _
                      "Failed to hash data, system error " _
                    & VBA.CStr(lngErr)
        Else
            lngDataLen = 4 '4 bytes for Long length.
            If CryptGetHashParam(hHmacHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
                lngErr = Err.LastDllError
                DestroyHandles
                Err.Raise vbObjectError Or &HC332&, _
                          TypeNameOfMe, _
                          "Failed to obtain hash value length, system error " _
                        & VBA.CStr(lngErr)
            Else
                lngDataLen = lngHashSize
                ReDim bytHashValue(lngDataLen - 1)
                If CryptGetHashParam(hHmacHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                    lngErr = Err.LastDllError
                    DestroyHandles
                    Err.Raise vbObjectError Or &HC334&, _
                              TypeNameOfMe, _
                              "Failed to obtain hash value, system error " _
                            & VBA.CStr(lngErr)
                Else
                    DestroyHandles
                    HMACSHA256 = bytHashValue
                End If
            End If
        End If
    End If
End Function

''
' Prepares to create HMACs by importing the key supplied as a Byte array.
'
' @method InitHMAC
' @param {Byte} Key
''
Public Sub InitHMAC(ByRef Key() As Byte)
    Dim kbKey As KEYBLOB
    Dim bytKbKey() As Byte
    Dim lngErr As Long

    DestroyHandles
    If hAdvProvider = 0 Then
        Err.Raise vbObjectError Or &HC342&, _
                  TypeNameOfMe, _
                  "No cryptographic RSA AES provider context"
    Else
        With kbKey
            With .hdr
                .bType = PLAINTEXTKEYBLOB
                .bVersion = CUR_BLOB_VERSION
                .aiKeyAlg = CALG_RC2
            End With
            .cbKeySize = UBound(Key) - LBound(Key) + 1
            ReDim bytKbKey(LenB(kbKey) + .cbKeySize - 1)
            CopyMemory VarPtr(bytKbKey(0)), VarPtr(kbKey), LenB(kbKey)
            CopyMemory VarPtr(bytKbKey(LenB(kbKey))), VarPtr(Key(LBound(Key))), .cbKeySize
        End With
        If CryptImportKey(hAdvProvider, _
                          VarPtr(bytKbKey(0)), _
                          UBound(bytKbKey) + 1, _
                          0, _
                          CRYPT_IPSEC_HMAC_KEY, _
                          hKey) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HC344&, _
                      TypeNameOfMe, _
                      "Failed to import key, system error " _
                    & VBA.CStr(lngErr)
        End If
    End If
End Sub

''
' Calculates the MD5 hash of the provided data in the form of
' a Byte array and returns it as a Byte array.
'
' @method MD5
' @param {Byte} Data
' @return {Byte}
''
Public Function MD5(ByRef Data() As Byte) As Byte()
    Dim hHash As LongPtr
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hBaseProvider = 0 Then
        Err.Raise vbObjectError Or &HC352&, _
                  TypeNameOfMe, _
                  "No cryptographic Base provider context"
    ElseIf CryptCreateHash(hBaseProvider, CALG_MD5, 0&, 0&, hHash) = 0 Then
        Err.Raise vbObjectError Or &HC354&, _
                  TypeNameOfMe, _
                  "Failed to create CryptoAPI Hash object, system error " _
                & VBA.CStr(Err.LastDllError)
    ElseIf CryptHashData(hHash, _
                         Data(LBound(Data)), _
                         UBound(Data) - LBound(Data) + 1, _
                         0&) = 0 Then
        CryptDestroyHash hHash
        Err.Raise vbObjectError Or &HC356&, _
                  TypeNameOfMe, _
                  "Failed to hash data, system error " _
                & VBA.CStr(Err.LastDllError)
    Else
        lngDataLen = 4 '4 bytes for Long length.
        If CryptGetHashParam(hHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
            CryptDestroyHash hHash
            Err.Raise vbObjectError Or &HC358&, _
                      TypeNameOfMe, _
                      "Failed to obtain hash value length, system error " _
                    & VBA.CStr(Err.LastDllError)
        Else
            lngDataLen = lngHashSize
            ReDim bytHashValue(lngDataLen - 1)
            
            If CryptGetHashParam(hHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                CryptDestroyHash hHash
                Err.Raise vbObjectError Or &HC35A&, _
                          TypeNameOfMe, _
                          "Failed to obtain hash value, system error " _
                        & VBA.CStr(Err.LastDllError)
            Else
                CryptDestroyHash hHash
                MD5 = bytHashValue
            End If
        End If
    End If
End Function

''
' Calculates the SHA256 hash of the provided data in the form of
' a Byte array and returns it as a Byte array.
'
' @method SHA256
' @param {Byte} Data
' @return {Byte}
''
Public Function SHA256(ByRef Data() As Byte) As Byte()
    Dim hHash As LongPtr
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hBaseProvider = 0 Then
        Err.Raise vbObjectError Or &HC352&, _
                  TypeNameOfMe, _
                  "No cryptographic Base provider context"
    ElseIf CryptCreateHash(hAdvProvider, CALG_SHA_256, 0&, 0&, hHash) = 0 Then
        Err.Raise vbObjectError Or &HC354&, _
                  TypeNameOfMe, _
                  "Failed to create CryptoAPI Hash object, system error " _
                & VBA.CStr(Err.LastDllError)
    ElseIf CryptHashData(hHash, _
                         Data(LBound(Data)), _
                         UBound(Data) - LBound(Data) + 1, _
                         0&) = 0 Then
        CryptDestroyHash hHash
        Err.Raise vbObjectError Or &HC356&, _
                  TypeNameOfMe, _
                  "Failed to hash data, system error " _
                & VBA.CStr(Err.LastDllError)
    Else
        lngDataLen = 4 '4 bytes for Long length.
        If CryptGetHashParam(hHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
            CryptDestroyHash hHash
            Err.Raise vbObjectError Or &HC358&, _
                      TypeNameOfMe, _
                      "Failed to obtain hash value length, system error " _
                    & VBA.CStr(Err.LastDllError)
        Else
            lngDataLen = lngHashSize
            ReDim bytHashValue(lngDataLen - 1)
            
            If CryptGetHashParam(hHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                CryptDestroyHash hHash
                Err.Raise vbObjectError Or &HC35A&, _
                          TypeNameOfMe, _
                          "Failed to obtain hash value, system error " _
                        & VBA.CStr(Err.LastDllError)
            Else
                CryptDestroyHash hHash
                SHA256 = bytHashValue
            End If
        End If
    End If
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub DestroyHandles(Optional ByVal Release As Boolean = False)
    On Error Resume Next 'Handle all exceptions here!
    If hHmacHash <> 0 Then CryptDestroyHash hHmacHash: hHmacHash = 0
    If hKey <> 0 Then CryptDestroyKey hKey: hKey = 0
    If Release And (hBaseProvider <> 0) Then CryptReleaseContext hBaseProvider, 0&: hBaseProvider = 0
    If Release And (hAdvProvider <> 0) Then CryptReleaseContext hAdvProvider, 0&: hAdvProvider = 0
    Err.Clear
End Sub

' ============================================= '
' Initialize/Terminate Methods
' ============================================= '

Private Sub Class_Terminate()
    DestroyHandles Release:=True
End Sub

Private Sub Class_Initialize()
    Dim strProvider As String
    
    TypeNameOfMe = TypeName(Me)
    
    'NOTE: Version probe hacks below.  These should defeat any use of
    '      version-lie appcompat shims by naive who try to use this class.
    '
    '      We need these because (a.) Windows 5.1 (XP and Server 2003) do
    '      not have support for CRYPT_STRING_NOCRLF or CRYPT_STRING_HEXRAW,
    '      and (b.) Windows XP does not support MS_ENH_RSA_AES_PROV and so
    '      we must request MS_ENH_RSA_AES_PROV_XP instead.
    
    On Error GoTo 0
    If CryptAcquireContext(hBaseProvider, _
                           0&, _
                           StrPtr(MS_DEFAULT_PROVIDER), _
                           PROV_RSA_FULL, _
                           CRYPT_VERIFYCONTEXT Or CRYPT_MACHINE_KEYSET) = 0 Then
        Err.Raise vbObjectError Or &HC366&, _
                  TypeNameOfMe, _
                  "Failed to obtain CryptoAPI Base context, system error " _
                & CStr(Err.LastDllError)
    ElseIf CryptAcquireContext(hAdvProvider, _
                               0&, _
                               StrPtr(MS_ENH_RSA_AES_PROV), _
                               PROV_RSA_AES, _
                               CRYPT_VERIFYCONTEXT Or CRYPT_MACHINE_KEYSET) = 0 Then
        If CryptAcquireContext(hAdvProvider, _
                                   0&, _
                                   StrPtr(MS_ENH_RSA_AES_PROV_XP), _
                                   PROV_RSA_AES, _
                                   CRYPT_VERIFYCONTEXT Or CRYPT_MACHINE_KEYSET) = 0 Then
            Err.Raise vbObjectError Or &HC368&, _
                      TypeNameOfMe, _
                      "Failed to obtain CryptoAPI RSA AES context, system error " _
                    & CStr(Err.LastDllError)
        End If
    End If
End Sub
