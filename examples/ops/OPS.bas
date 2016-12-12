Attribute VB_Name = "OPS"
Private pOPSConsumerKey As String
Private pOPSConsumerSecret As String
Private pClient As WebClient

Private Property Get OPSConsumerKey() As String
    If pOPSConsumerKey = "" Then
        If Credentials.Loaded Then
            pOPSConsumerKey = Credentials.Values("OPS")("consumer_key")
        Else
            pOPSConsumerKey = InputBox("Please Enter OPS Consumer Key")
        End If
    End If
    
    OPSConsumerKey = pOPSConsumerKey
End Property
Private Property Get OPSConsumerSecret() As String
    If pOPSConsumerSecret = "" Then
        If Credentials.Loaded Then
            pOPSConsumerSecret = Credentials.Values("OPS")("consumer_secret")
        Else
            pOPSConsumerSecret = InputBox("Please Enter OPS Consumer Secret")
        End If
    End If
    
    OPSConsumerSecret = pOPSConsumerSecret
End Property

Public Property Get Client() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "https://ops.epo.org/3.1/"
        
        ' Setup authenticator (note: provide consumer key and secret here
        Dim Auth As New OPSAuthenticator
        Auth.Setup OPSConsumerKey, OPSConsumerSecret
        
        ' If there are issues automatically getting the token with consumer key / secret
        ' the token can be found in the developer console and manually entered here
        ' Auth.Token = "AUTH_TOKEN"
        
        Set pClient.Authenticator = Auth
        
        ' Add XML converter
        WebHelpers.RegisterConverter "xml", "application/xml", "OPS.ConvertToXml", "OPS.ParseXml"
    End If
    
    Set Client = pClient
End Property

Public Function Search(Query As String) As Collection
#If Mac Then
    Err.Raise 11099, Description:="XML services (such as the OPS example) are not currently supported on the Mac (Note: OPS supports JSON, but XML is used for this example)"
#Else
    Dim Request As New WebRequest
    Request.Resource = "rest-services/published-data/search"
    Request.CustomResponseFormat = "xml"
    Request.AddQuerystringParam "q", Query
    
    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    If Response.StatusCode = WebStatusCode.Ok Then
        Set Search = GetBiblioData(GetDocNumbers(Response.Data))
    End If
#End If
End Function

Public Function GetBiblioData(DocNumbers As Variant) As Collection
    Dim Request As New WebRequest
    Request.Resource = "rest-services/published-data/publication/epodoc/{number}/biblio"
    Request.AddUrlSegment "number", VBA.Join(DocNumbers, ",")
    Request.CustomResponseFormat = "xml"
    
    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
    
    If Response.StatusCode = WebStatusCode.Ok Then
        Dim Documents As Object
        Dim Doc As Object
        Dim Results As New Collection
        Dim Result As Dictionary
        Dim Child As Object
        Dim Title As String
        Dim Index As Long
        
        Set Documents = GetChild(GetChild(Response.Data, "ops:world-patent-data"), "exchange-documents")
        Index = 0
        For Each Doc In Documents.ChildNodes
            ' Get English title
            For Each Child In GetChildren(GetChild(Doc, "bibliographic-data"), "invention-title")
                If GetAttribute(Child, "lang") = "en" Then
                    Title = Child.Text
                    Exit For
                End If
            Next Child
            
            Set Result = New Dictionary
            Result.Add "title", Title
            Result.Add "number", DocNumbers(Index)
            
            Results.Add Result
            
            Index = Index + 1
        Next Doc
        
        Set GetBiblioData = Results
    End If
End Function

Private Function GetDocNumbers(SearchData As Object) As Variant
    Dim Results As Object
    Dim DocNumbers() As String
    Dim Child As Object
    Dim Index As Long
    Dim Country As String
    Dim Num As String
    Dim Kind As String
    
    Set Results = GetChild(GetChild(GetChild(SearchData, "ops:world-patent-data"), "ops:biblio-search"), "ops:search-result").ChildNodes
    ReDim DocNumbers(Results.Length - 1)
    Index = 0
    For Each Child In Results
        Country = GetChild(GetChild(Child, "document-id"), "country").Text
        Num = GetChild(GetChild(Child, "document-id"), "doc-number").Text
        Kind = GetChild(GetChild(Child, "document-id"), "kind").Text
        
        DocNumbers(Index) = Country & Num & "." & Kind
        Index = Index + 1
    Next Child
    
    GetDocNumbers = DocNumbers
End Function

' Enable XML parsing/converting
' https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0
Public Function ParseXml(Value As String) As Object
    Set ParseXml = CreateObject("MSXML2.DOMDocument")
    ParseXml.Async = False
    ParseXml.LoadXML Value
End Function

Public Function ConvertToXml(Value As Variant) As String
    ConvertToXml = VBA.Trim$(VBA.Replace(Value.Xml, vbCrLf, ""))
End Function

Private Function GetChildren(Node As Object, Name As String) As Collection
    Dim Child As Object
    Dim Children As New Collection
    For Each Child In Node.ChildNodes
        If Child.nodeName = Name Then
            Children.Add Child
        End If
    Next Child
    
    Set GetChildren = Children
End Function

Private Function GetChild(Node As Object, Name As String) As Object
    Dim Child As Object
    For Each Child In Node.ChildNodes
        If Child.nodeName = Name Then
            Set GetChild = Child
            Exit Function
        End If
    Next Child
End Function

Private Function GetAttribute(Node As Object, Name As String) As String
    Dim Attr As Object
    For Each Attr In Node.Attributes
        If Attr.Name = Name Then
            GetAttribute = Attr.Text
            Exit Function
        End If
    Next Attr
End Function
