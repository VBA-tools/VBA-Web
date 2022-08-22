Attribute VB_Name = "XmlConverter"
''
' VBA-XML v0.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-XML
'
' XML Converter for VBA
'
' Design:
' The goal is to have the general form of MSXML2.DOMDocument (albeit not feature complete)
'
' ParseXML(<messages><message id="1">A</message><message id="2">B</message></messages>) ->
'
' {Dictionary}
' - nodeName: {String} "#document"
' - attributes: {Collection} (Nothing)
' - childNodes: {Collection}
'   {Dictionary}
'   - nodeName: "messages"
'   - attributes: (empty)
'   - childNodes:
'     {Dictionary}
'     - nodeName: "message"
'     - attributes:
'       {Collection of Dictionary}
'       nodeName: "id"
'       text: "1"
'     - childNodes: (empty)
'     - text: A
'     {Dictionary}
'     - nodeName: "message"
'     - attributes:
'       {Collection of Dictionary}
'       nodeName: "id"
'       text: "2"
'     - childNodes: (empty)
'     - text: B
'
' Errors:
' 10101 - XML parse error
'
' References:
' - http://www.w3.org/TR/REC-xml/
'
' @author tim.hall.engr@gmail.com
' @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php
' @depend UtcConverter
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' VBA-Git Annotations
' https://github.com/VBA-Tools-v2/VBA-Git | https://radiuscore.co.nz
'
' @excludeobfuscation
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder VBA-Web.Helpers
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Private Module

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private Type xml_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-XML will use String for numbers longer than 15 characters that contain only digits
    ' to override set this option to `True`.
    UseDoubleForLargeNumbers As Boolean
    
    ' Use this option to include Node mapping (`parentNode`, `firstChild`, `lastChild`) in parsed object.
    ' Performance suffers (slightly) when including node mapping in object structure.
    IncludeNodeMapping As Boolean
    
    ' Internal VBA-XML parser is much slower than using `MSXML2.DOMDocument`. By default on Windows
    ' machines `MSXML2.DOMDocument` is used. Set this option to `True` to force use of VBA-XML.
    ' Not recommended if dealing with large XML strings (>1,000,000 char).
    '
    ' This option has no effect on Mac machines.
    ForceVbaXml As Boolean
End Type

Public XmlOptions As xml_Options

' ============================================= '
' Public Methods
' ============================================= '

' --------------------------------------------- '
' Helper Methods
' --------------------------------------------- '

''
' Helper for use with VBA-XML.
'
' Create a basic Node in dictionary structure as required by `ConvertToXml`.
' Can be used when building a XML body for `WebRequest`.
'
' @method CreateNode
' @param {String} Name | `nodeName`
' @param {Variant} Value | `nodeValue` Leave blank if void element.
' @param {Collection} ChildNodes | `childNodes` A collection of Node dictionaries, as created by `CreateNode`.
' @param {Collection} Attributes | `attributes` a collection of Attribute dictionaries, as created by `CreateAttribute`.
' @return {Dictionary}
''
Public Function CreateNode(ByVal Name As String, Optional ByVal Value As Variant = Null, Optional ChildNodes As Collection, Optional Attributes As Collection) As Dictionary
    Dim web_Node As Dictionary
    Set web_Node = New Dictionary
    
    With web_Node
        .Add "attributes", Attributes ' Can be `Nothing` if no attributes.
        If ChildNodes Is Nothing Then
            .Add "childNodes", New Collection ' Even if there are no child nodes, must be set to an empty collection.
        Else
            .Add "childNodes", ChildNodes
        End If
        On Error Resume Next
            .Add "text", VBA.CStr(Value) ' Attempt to convert `nodeValue` to `text` using VBA. Ignore errors.
        On Error GoTo 0
        .Add "nodeValue", Value
        .Add "nodeName", Name
    End With
    
    Set CreateNode = web_Node
End Function

''
' Helper for use with VBA-XML.
'
' Create an attribute Name-Value pair in a dictionary structure as required by `ConvertToXml`.
' Can be used when building a XML body for `WebRequest`.
'
' @method CreateAttribute
' @param {String} Name
' @param {Value} Value
' @return {Dictionary}
''
Public Function CreateAttribute(ByVal Name As String, Optional ByVal Value As String) As Dictionary
    Dim web_Attribute As Dictionary
    Set web_Attribute = New Dictionary
    
    web_Attribute.Add "name", Name
    web_Attribute.Add "value", Value
    
    Set CreateAttribute = web_Attribute
End Function

''
' Helper for use with VBA-XML.
'
' Return first node with given `nodeName` in `Node.childNodes`.
'
' @method SelectSingleNode
' @param {Dictionary} Node | Node to search within.
' @param {String} nodeName | Xpath search expression.
' @return {Dictionary} Node (if found), else Nothing.
''
Public Function SelectSingleNode(Node As Dictionary, nodeName As String) As Dictionary
    Dim xml_Node As Dictionary
    Dim xml_Item As Variant
    
    If VBA.InStr(nodeName, "/") Then
        ' Recursively search through nodes.
        Set xml_Node = Node
        For Each xml_Item In VBA.Split(nodeName, "/")
            Set xml_Node = SelectSingleNode(xml_Node, VBA.CStr(xml_Item))
            If xml_Node Is Nothing Then Exit Function
        Next xml_Item
        Set SelectSingleNode = xml_Node
    Else
        ' Search childNodes for given nodeName.
        For Each xml_Node In Node.Item("childNodes")
            If xml_Node.Item("nodeName") = nodeName Then
                Set SelectSingleNode = xml_Node
                Exit Function
            End If
        Next xml_Node
    End If
End Function

' --------------------------------------------- '
' Core Methods
' --------------------------------------------- '

''
' Convert XML string to Dictionary or DOMDocument (windows only).
'
' @method ParseXml
' @param {String} XmlString
' @return {DOMDocument|Dictionary}
''
Public Function ParseXml(ByVal XmlString As String) As Object
    Dim xml_String As String
    Dim xml_Index As Long
    xml_Index = 1

    ' Remove vbTab from xml_String
    xml_String = VBA.Replace(XmlString, VBA.vbTab, vbNullString)

    xml_SkipSpaces xml_String, xml_Index
    If Not VBA.Mid$(xml_String, xml_Index, 1) = "<" Then
        ' Error: Invalid XML string
        Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '<'")
    Else
#If Mac Then
        Set ParseXml = New Dictionary
        ParseXml.Add "prolog", xml_ParseProlog(xml_String, xml_Index)
        ParseXml.Add "doctype", xml_ParseDoctype(xml_String, xml_Index)
        ParseXml.Add "nodeName", "#document"
        ParseXml.Add "attributes", Nothing
        ParseXml.Add "childNodes", New Collection
        ParseXml.Item("childNodes").Add xml_ParseNode(xml_String, xml_Index, VBA.IIf(XmlOptions.IncludeNodeMapping, ParseXml, Nothing))
#Else
        If XmlOptions.ForceVbaXml Then
            Set ParseXml = New Dictionary
            ParseXml.Add "prolog", xml_ParseProlog(xml_String, xml_Index)
            ParseXml.Add "doctype", xml_ParseDoctype(xml_String, xml_Index)
            ParseXml.Add "nodeName", "#document"
            ParseXml.Add "attributes", Nothing
            ParseXml.Add "childNodes", New Collection
            ParseXml.Item("childNodes").Add xml_ParseNode(xml_String, xml_Index, VBA.IIf(XmlOptions.IncludeNodeMapping, ParseXml, Nothing))
        Else
            Set ParseXml = CreateObject("MSXML2.DOMDocument")
            ParseXml.Async = False
            ParseXml.LoadXml XmlString
        End If
#End If
    End If
End Function

''
' Convert object (Dictionary/Collection/DOMDocument) to XML string.
'
' @method ConvertToXml
' @param {Variant} XmlValue (Dictionary, Collection, or DOMDocument)
' @param {Integer|String} Whitespace "Pretty" print xml with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ConvertToXml(ByVal XmlValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal xml_CurrentIndentation As Long = 0) As String
    Dim xml_Buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    Dim xml_Indentation As String
    Dim xml_PrettyPrint As Boolean
    Dim xml_Converted As String
    Dim xml_ChildNode As Variant
    Dim xml_Attribute As Variant
    
    xml_PrettyPrint = Not IsMissing(Whitespace)
    
    Select Case VBA.VarType(XmlValue)
    Case VBA.vbNull
        ConvertToXml = vbNullString
    Case VBA.vbDate
        ConvertToXml = ConvertToIso(VBA.CDate(XmlValue))
    Case VBA.vbString
        If Not XmlOptions.UseDoubleForLargeNumbers And xml_StringIsLargeNumber(XmlValue) Then
            ConvertToXml = XmlValue
        Else
            ConvertToXml = xml_Encode(XmlValue)
        End If
    Case VBA.vbBoolean
        ConvertToXml = VBA.IIf(XmlValue, "true", "false")
    Case VBA.vbObject
        If xml_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
            Else
                xml_Indentation = VBA.Space$((xml_CurrentIndentation) * Whitespace)
            End If
        End If
        
        ' Dictionary (Node).
        If VBA.TypeName(XmlValue) = "Dictionary" Then
            ' If root node, parse prolog and child nodes then exit.
            If XmlValue.Item("nodeName") = "#document" Then
                If Not XmlValue.Item("prolog") = vbNullString Then
                    xml_BufferAppend xml_Buffer, XmlValue.Item("prolog"), xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength ' Always put prolog on its own line.
                End If
                xml_Converted = ConvertToXml(XmlValue.Item("childNodes"), Whitespace, xml_CurrentIndentation)
                xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
                Exit Function
            Else
                ' Validate Dictionary structure.
                If Not XmlValue.Exists("nodeName") Or Not XmlValue.Exists("nodeValue") Then
                    Err.Raise 11001, "XMLConverter", "Error parsing XML:" & VBA.vbNewLine & Err.Number & " - " & Err.Description & _
                        "Poorly structured XML Dictionary. Use `ParseXml` with `XmlOptions.ForceVbaXml = True` OR " & _
                        "`CreateNode` and `CreateAttribute` to create a correctly structured XML dictionary object."
                End If
            
                ' Add 'Start Tag'.
                xml_BufferAppend xml_Buffer, xml_Indentation & "<", xml_BufferPosition, xml_BufferLength
                xml_BufferAppend xml_Buffer, XmlValue.Item("nodeName"), xml_BufferPosition, xml_BufferLength
                If XmlValue.Exists("attributes") Then
                    If Not XmlValue.Item("attributes") Is Nothing Then
                        For Each xml_Attribute In XmlValue.Item("attributes")
                            xml_BufferAppend xml_Buffer, " ", xml_BufferPosition, xml_BufferLength
                            xml_BufferAppend xml_Buffer, xml_Attribute.Item("name"), xml_BufferPosition, xml_BufferLength
                            xml_BufferAppend xml_Buffer, "=""", xml_BufferPosition, xml_BufferLength
                            xml_BufferAppend xml_Buffer, xml_Encode(xml_Attribute.Item("value"), """"), xml_BufferPosition, xml_BufferLength
                            xml_BufferAppend xml_Buffer, """", xml_BufferPosition, xml_BufferLength
                        Next xml_Attribute
                    End If
                End If
                
                ' Check for void node.
                If xml_IsVoidNode(XmlValue) Then
                    ' Add 'Empty Element' tag and exit.
                    xml_BufferAppend xml_Buffer, "/>", xml_BufferPosition, xml_BufferLength
                    If xml_PrettyPrint Then
                        xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                        
                        If VBA.VarType(Whitespace) = VBA.vbString Then
                            xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                        Else
                            xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                        End If
                    End If
                    ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
                    Exit Function
                Else
                    ' Finish 'Start Tag' and continue.
                    xml_BufferAppend xml_Buffer, ">", xml_BufferPosition, xml_BufferLength
                End If
                
                ' Add node content.
                If XmlValue.Exists("childNodes") Then
                    If XmlValue.Item("childNodes").Count > 0 Then
                        If xml_PrettyPrint Then
                            xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
            
                            If VBA.VarType(Whitespace) = VBA.vbString Then
                                xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                            Else
                                xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                            End If
                        End If
                    
                        ' Convert childNodes.
                        xml_Converted = ConvertToXml(XmlValue.Item("childNodes"), Whitespace, xml_CurrentIndentation + 1)
                        xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                        xml_BufferAppend xml_Buffer, xml_Indentation, xml_BufferPosition, xml_BufferLength
                    Else
                        ' No child nodes, add text.
                        xml_Converted = ConvertToXml(XmlValue.Item("nodeValue"), Whitespace, xml_CurrentIndentation + 1)
                        xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                    End If
                Else
                    ' No child nodes, add text.
                    xml_Converted = ConvertToXml(XmlValue.Item("nodeValue"), Whitespace, xml_CurrentIndentation + 1)
                    xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                End If
                
                ' Add 'End Tag'.
                xml_BufferAppend xml_Buffer, "</", xml_BufferPosition, xml_BufferLength
                xml_BufferAppend xml_Buffer, XmlValue.Item("nodeName"), xml_BufferPosition, xml_BufferLength
                xml_BufferAppend xml_Buffer, ">", xml_BufferPosition, xml_BufferLength
                
                If xml_PrettyPrint Then
                    xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                    
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                    Else
                        xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                    End If
                End If
            End If
            ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
        
        ' Collection (child nodes)
        ElseIf VBA.TypeName(XmlValue) = "Collection" Then
            For Each xml_ChildNode In XmlValue
                ' Convert node.
                xml_Converted = ConvertToXml(xml_ChildNode, Whitespace, xml_CurrentIndentation)
                If Not xml_Converted = vbNullString Then
                    xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                Else
                    xml_BufferAppend xml_Buffer, "null", xml_BufferPosition, xml_BufferLength
                End If
            Next xml_ChildNode
            
            ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
        
        ' MSXML2.DOMDocument (windows only)
        ElseIf VBA.TypeName(XmlValue) = "DOMDocument" Then
            ' Parse document child nodes.
            ConvertToXml = ConvertToXml(XmlValue.ChildNodes, Whitespace, xml_CurrentIndentation)
        
        ' Prolog (windows only)
        ElseIf VBA.TypeName(XmlValue) = "IXMLDOMProcessingInstruction" Then
            ' Manually parse prolog, as using `XML` property results in lost data (i.e. encoding).
            xml_BufferAppend xml_Buffer, "<?xml ", xml_BufferPosition, xml_BufferLength
            xml_BufferAppend xml_Buffer, XmlValue.Text, xml_BufferPosition, xml_BufferLength
            xml_BufferAppend xml_Buffer, "?>", xml_BufferPosition, xml_BufferLength
            xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength ' Always put prolog on its own line.
            ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
            
        ' Node (windows only)
        ElseIf VBA.TypeName(XmlValue) = "IXMLDOMElement" Then
        
            ' Add 'Start Tag' (incl. attributes).
            xml_BufferAppend xml_Buffer, xml_Indentation & "<", xml_BufferPosition, xml_BufferLength
            xml_BufferAppend xml_Buffer, XmlValue.nodeName, xml_BufferPosition, xml_BufferLength
            If Not XmlValue.Attributes Is Nothing Then
                For Each xml_Attribute In XmlValue.Attributes
                    xml_BufferAppend xml_Buffer, " ", xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, xml_Attribute.Name, xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, "=""", xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, xml_Encode(xml_Attribute.Value, """"), xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, """", xml_BufferPosition, xml_BufferLength
                Next xml_Attribute
            End If
            
            ' Check for void node.
            If xml_IsVoidNode(XmlValue) Then
                ' Add 'Empty Element' tag and exit.
                xml_BufferAppend xml_Buffer, "/>", xml_BufferPosition, xml_BufferLength
                If xml_PrettyPrint Then
                    xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                    
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                    Else
                        xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                    End If
                End If
                ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
                Exit Function
            Else
                ' Finish 'Start Tag' and continue.
                xml_BufferAppend xml_Buffer, ">", xml_BufferPosition, xml_BufferLength
            End If
            
            ' Add node content.
            If XmlValue.ChildNodes.Length > 0 Then
                ' Child node represents the node's text. treat as though it has no child nodes and just add text.
                If XmlValue.ChildNodes.Length = 1 And _
                    (VBA.TypeName(XmlValue.ChildNodes.Item(0)) = "IXMLDOMText" Or VBA.TypeName(XmlValue.ChildNodes.Item(0)) = "IXMLDOMCDATASection") Then
                    Select Case VBA.TypeName(XmlValue.ChildNodes.Item(0))
                    Case "IXMLDOMText"
                        ' Pass value through converter to ensure characters are escaped & converted to text correctly.
                        xml_Converted = ConvertToXml(XmlValue.Text, Whitespace, xml_CurrentIndentation + 1)
                        xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                    Case "IXMLDOMCDATASection"
                        ' CDATA node doesn't pass through converter, as it does not need escaping.
                        xml_BufferAppend xml_Buffer, XmlValue.ChildNodes.Item(0).Xml, xml_BufferPosition, xml_BufferLength
                    End Select
                Else
                    If xml_PrettyPrint Then
                        xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
        
                        If VBA.VarType(Whitespace) = VBA.vbString Then
                            xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                        Else
                            xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                        End If
                    End If
                
                    ' Convert childNodes.
                    xml_Converted = ConvertToXml(XmlValue.ChildNodes, Whitespace, xml_CurrentIndentation + 1)
                    xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, xml_Indentation, xml_BufferPosition, xml_BufferLength
                End If
            Else
                ' No child nodes, add text.
                xml_Converted = ConvertToXml(XmlValue.Text, Whitespace, xml_CurrentIndentation + 1)
                xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
            End If
            
            ' Add 'End Tag'.
            xml_BufferAppend xml_Buffer, "</", xml_BufferPosition, xml_BufferLength
            xml_BufferAppend xml_Buffer, XmlValue.nodeName, xml_BufferPosition, xml_BufferLength
            xml_BufferAppend xml_Buffer, ">", xml_BufferPosition, xml_BufferLength
            
            If xml_PrettyPrint Then
                xml_BufferAppend xml_Buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                
                If VBA.VarType(Whitespace) = VBA.vbString Then
                    xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                Else
                    xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                End If
            End If
            
            ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
        
        ' Child Nodes (windows only)
        ElseIf VBA.TypeName(XmlValue) = "IXMLDOMNodeList" Then
        
            For Each xml_ChildNode In XmlValue
                ' Convert node.
                xml_Converted = ConvertToXml(xml_ChildNode, Whitespace, xml_CurrentIndentation)
                If Not xml_Converted = vbNullString Then
                    xml_BufferAppend xml_Buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                Else
                    xml_BufferAppend xml_Buffer, "null", xml_BufferPosition, xml_BufferLength
                End If
            Next xml_ChildNode
            
            ConvertToXml = xml_BufferToString(xml_Buffer, xml_BufferPosition)
        Else
            Err.Raise 11001, "XMLConverter", "Error parsing XML:" & VBA.vbNewLine & _
                        "`" & VBA.TypeName(XmlValue) & "` is a unrecognised XML object. ConvertToXml method will need " & _
                        "to be updated to correctly convert this XML object."
        End If
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ConvertToXml = VBA.Replace(XmlValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ConvertToXml = XmlValue
        On Error GoTo 0
    End Select
    Exit Function
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function xml_ParseProlog(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_OpeningLevel As Long
    Dim xml_StringLength As Long
    Dim xml_StartIndex As Long
    Dim xml_Chars As String

    xml_SkipSpaces xml_String, xml_Index
    If VBA.Mid$(xml_String, xml_Index, 2) = "<?" Then
        xml_StartIndex = xml_Index
        xml_Index = xml_Index + 2
        xml_StringLength = Len(xml_String)
    
        ' Find matching closing tag, ?>
        Do
            xml_Chars = VBA.Mid$(xml_String, xml_Index, 2)
            
            If xml_Index + 1 > xml_StringLength Then
                Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '?>'")
            ElseIf xml_OpeningLevel = 0 And xml_Chars = "?>" Then
                xml_Index = xml_Index + 2
                Exit Do
            ElseIf xml_Chars = "<?" Then
                xml_OpeningLevel = xml_OpeningLevel + 1
                xml_Index = xml_Index + 2
            ElseIf xml_Chars = "?>" Then
                xml_OpeningLevel = xml_OpeningLevel - 1
                xml_Index = xml_Index + 2
            Else
                xml_Index = xml_Index + 1
            End If
        Loop
        
        xml_ParseProlog = VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
    End If
End Function

Private Function xml_ParseDoctype(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_OpeningLevel As Long
    Dim xml_StringLength As Long
    Dim xml_StartIndex As Long
    Dim xml_Char As String
    
    xml_SkipSpaces xml_String, xml_Index
    If VBA.Mid$(xml_String, xml_Index, 2) = "<!" Then
        xml_StartIndex = xml_Index
        xml_Index = xml_Index + 2
        xml_StringLength = Len(xml_String)
        
        ' Find matching closing tag, >
        Do
            xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
            xml_Index = xml_Index + 1
            
            If xml_Index > xml_StringLength Then
                Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '>'")
            ElseIf xml_OpeningLevel = 0 And xml_Char = ">" Then
                Exit Do
            ElseIf xml_Char = "<" Then
                xml_OpeningLevel = xml_OpeningLevel + 1
            ElseIf xml_Char = ">" Then
                xml_OpeningLevel = xml_OpeningLevel - 1
            End If
        Loop
        
        xml_ParseDoctype = VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
    End If
End Function

''
' Parse Node Attributes.
'
' <title lang="en">Harry Potter</title>
'       ^         ^
'     Start      End
'
' {Dictionary} Attribute
' -> Key: Name  Value: lang
' -> Key: Value Value: en
'
' @method xml_ParseAttributes
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index | Current index position in XML string.
' @return {Collection} Collection of attributes (Dictionary).
''
Private Function xml_ParseAttributes(xml_String As String, ByRef xml_Index As Long) As Collection
    Dim xml_Char As String
    Dim xml_StartIndex As Long
    Dim xml_Quote As String
    Dim xml_Name As String
    
    Set xml_ParseAttributes = New Collection
    xml_SkipSpaces xml_String, xml_Index
    xml_StartIndex = xml_Index
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
        
        Select Case xml_Char
        Case "="
            If xml_Name = vbNullString Then
                ' Found end of attribute name
                ' Extract name, skip '=', find quote char, reset start index
                xml_Name = VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
                xml_Index = xml_Index + 1
                xml_Quote = VBA.Mid$(xml_String, xml_Index, 1)
                xml_Index = xml_Index + 1
                xml_StartIndex = xml_Index
                
                ' Check for valid quote style of attribute value
                If Not xml_Quote = """" And Not xml_Quote = "'" Then
                    ' Invalid Attribute quote.
                    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting ''' or '""'")
                End If
            Else
                ' '=' exists within attribute value. Continue.
                xml_Index = xml_Index + 1
            End If
        Case xml_Quote
            ' Found end of attribute value
            ' Store name, value as new attribute.
            With xml_ParseAttributes
                .Add New Dictionary
                .Item(.Count).Add "name", xml_Name
                .Item(.Count).Add "value", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
            End With
            
            ' Reset variables.
            xml_Name = vbNullString
            xml_Quote = vbNullString
            
            ' Increment.
            xml_Index = xml_Index + 1
            xml_SkipSpaces xml_String, xml_Index
            xml_StartIndex = xml_Index
            
            ' Check for end of tag.
            If VBA.Mid$(xml_String, xml_Index, 1) = ">" Or VBA.Mid$(xml_String, xml_Index, 2) = "/>" Then
                Exit Function ' End of tag, exit.
            End If
        Case Else
            xml_Index = xml_Index + 1
        End Select
    Loop
    
    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '>' or '/>'")
End Function

Private Function xml_ParseNode(xml_String As String, ByRef xml_Index As Long, Optional ByRef xml_Parent As Dictionary) As Dictionary
    Dim xml_StartIndex As Long
    
    xml_SkipSpaces xml_String, xml_Index
    If VBA.Mid$(xml_String, xml_Index, 1) <> "<" Then
        Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '<'")
    Else
        ' Skip opening bracket
        xml_StartIndex = xml_Index
        xml_Index = xml_Index + 1
        
        ' Initialize node
        Set xml_ParseNode = New Dictionary
        If XmlOptions.IncludeNodeMapping Then
            xml_ParseNode.Add "parentNode", xml_Parent
        End If
        xml_ParseNode.Add "attributes", Nothing
        xml_ParseNode.Add "childNodes", New Collection
        xml_ParseNode.Add "text", vbNullString
        If XmlOptions.IncludeNodeMapping Then
            xml_ParseNode.Add "firstChild", Nothing
            xml_ParseNode.Add "lastChild", Nothing
        End If
        xml_ParseNode.Add "nodeValue", Null
        
        ' 1. Parse nodeName
        xml_ParseNode.Add "nodeName", xml_ParseName(xml_String, xml_Index)
        
        ' 2. Parse attributes
        If VBA.Mid$(xml_String, xml_Index, 2) = "/>" Then
            ' '/>' is the 'Empty-element' tag. Nothing more to parse. Skip over closing '/>' and exit
            xml_Index = xml_Index + 2
            xml_ParseNode.Add "xml", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex) ' Add 'xml' text.
            Exit Function
        ElseIf VBA.Mid$(xml_String, xml_Index, 1) = ">" Then
            ' If '>' then end of Start Tag. Skip over closing '>' and continue.
            xml_Index = xml_Index + 1
        Else
            ' If not '/>' or '>' then attributes are present within Start Tag.
            Set xml_ParseNode.Item("attributes") = xml_ParseAttributes(xml_String, xml_Index)
            
            ' Re-do previous checks as index has moved to after attributes.
            If VBA.Mid$(xml_String, xml_Index, 2) = "/>" Then
                ' '/>' is the 'Empty-element' tag. Nothing more to parse. Skip over closing '/>' and exit
                xml_Index = xml_Index + 2
                xml_ParseNode.Add "xml", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex) ' Add 'xml' text.
                Exit Function
            ElseIf VBA.Mid$(xml_String, xml_Index, 1) = ">" Then
                ' If '>' then end of Start Tag. Skip over closing '>' and continue.
                xml_Index = xml_Index + 1
            End If
        End If
    
        ' 3. Parse node content (child nodes, text, value).
        xml_SkipSpaces xml_String, xml_Index
        If Not VBA.Mid$(xml_String, xml_Index, 2) = "</" Then
            If VBA.Mid$(xml_String, xml_Index, 1) = "<" Then
                ' If '<' (but not '</'), then child node exists.
                If XmlOptions.IncludeNodeMapping Then
                    Set xml_ParseNode.Item("childNodes") = xml_ParseChildNodes(xml_String, xml_Index, xml_ParseNode)
                    Set xml_ParseNode.Item("firstChild") = xml_ParseNode.Item("childNodes").Item(1)
                    Set xml_ParseNode.Item("lastChild") = xml_ParseNode.Item("childNodes").Item(xml_ParseNode.Item("childNodes").Count)
                Else
                    Set xml_ParseNode.Item("childNodes") = xml_ParseChildNodes(xml_String, xml_Index)
                End If
                ' Set node 'Text' once child nodes are parsed (node text is space separated text of all child nodes).
                Dim xml_Buffer As String
                Dim xml_BufferPosition As Long
                Dim xml_BufferLength As Long
                Dim xml_ChildNode As Dictionary
                For Each xml_ChildNode In xml_ParseNode.Item("childNodes")
                    xml_BufferAppend xml_Buffer, xml_ChildNode.Item("text"), xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_Buffer, " ", xml_BufferPosition, xml_BufferLength
                Next xml_ChildNode
                xml_ParseNode.Item("text") = xml_BufferToString(xml_Buffer, xml_BufferPosition - 1)
            Else
                ' No child nodes. Set Node Text.
                xml_ParseNode.Item("text") = xml_ParseText(xml_String, xml_Index)
                xml_ParseNode.Item("nodeValue") = xml_ParseValue(xml_ParseNode.Item("text"))
            End If
        End If

        ' Skip over End-Tag '</' + 'nodeName' + '>'.
        xml_Index = xml_Index + 2 + VBA.Len(xml_ParseNode.Item("nodeName")) + 1
        
        ' Add 'xml' text.
        xml_ParseNode.Add "xml", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
    End If
End Function

''
' Call 'xml_ParseNode' to parse each child node.
'
'  <book category="cooking">
'    <title lang="en">Everyday Italian</title>
'    ^
'  Start
'    <author>Giada De Laurentiis</author>
'    <year>2005</year>
'    <price>30.00</price>
'  </book>
'  ^
' End
'
' @method xml_ParseChildNodes
' @param {Dictionary} xml_Node | Parent Node.
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index |  Current index position in XML string.
''
Private Function xml_ParseChildNodes(xml_String As String, ByRef xml_Index As Long, Optional ByRef xml_Parent As Dictionary) As Collection
    Set xml_ParseChildNodes = New Collection
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_SkipSpaces xml_String, xml_Index
        If VBA.Mid$(xml_String, xml_Index, 2) = "</" Then
            Exit Function
        ElseIf VBA.Mid$(xml_String, xml_Index, 1) = "<" Then
            xml_ParseChildNodes.Add xml_ParseNode(xml_String, xml_Index, VBA.IIf(XmlOptions.IncludeNodeMapping, xml_Parent, Nothing))
        Else
            Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '</' or '<'")
        End If
    Loop
End Function

''
' Parse Node Name.
'
' <title lang="en">Harry Potter</title>
'  ^     ^
'Start  End  |   Name --> 'title'
'
' <author>Giada De Laurentiis</author>
'  ^     ^
'Start  End  |   Name --> 'author'
'
' <price />
'  ^     ^
'Start  End  |   Name --> 'price'
'
' @method xml_ParseName
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index | Current index position in XML string.
' @return {String} nodeName
''
Private Function xml_ParseName(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_Char As String
    Dim xml_Buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    
    xml_SkipSpaces xml_String, xml_Index
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
        
        Select Case xml_Char
        Case " ", ">", "/"
            xml_ParseName = xml_BufferToString(xml_Buffer, xml_BufferPosition)
            If xml_Char = " " Then xml_Index = xml_Index + 1 ' Skip space
            Exit Function
        Case Else
            xml_BufferAppend xml_Buffer, xml_Char, xml_BufferPosition, xml_BufferLength
            xml_Index = xml_Index + 1
        End Select
    Loop
            
    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting ' ', '>', or '/>'")
End Function

''
' Parse Node text.
'
' <title lang="en">Harry Potter</title>
'                  ^           ^
'                Start        End
' Text --> 'Harry Potter'
'
' @method xml_ParseText
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index | Current index position in XML string.
' @return {String} Node text
''
Private Function xml_ParseText(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_Char As String
    Dim xml_Buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    Dim xml_StartIndex As Long
    Dim xml_EncodedFound As Boolean
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_Char = VBA.Mid$(xml_String, xml_Index, 1)

        Select Case xml_Char
        Case "<" 'Closing tag.
            xml_ParseText = xml_BufferToString(xml_Buffer, xml_BufferPosition)
            Exit Function
        Case "&"
            ' Remove encoding from XML string. See `xml_Encode` for additional information.
            ' Store start of encoded char and continue.
            xml_StartIndex = xml_Index
            xml_Index = xml_Index + 1
            xml_EncodedFound = False
            ' Find close of encoded char.
            Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
                xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
                Select Case xml_Char
                Case ";"
                    xml_EncodedFound = True
                    Select Case VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex + 1)
                    Case "&quot;"
                        xml_BufferAppend xml_Buffer, """", xml_BufferPosition, xml_BufferLength
                    Case "&amp;"
                        xml_BufferAppend xml_Buffer, "&", xml_BufferPosition, xml_BufferLength
                    Case "&apos;"
                        xml_BufferAppend xml_Buffer, "'", xml_BufferPosition, xml_BufferLength
                    Case "&lt;"
                        xml_BufferAppend xml_Buffer, "<", xml_BufferPosition, xml_BufferLength
                    Case "&gt;"
                        xml_BufferAppend xml_Buffer, ">", xml_BufferPosition, xml_BufferLength
                    Case Else
                        Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '&quot;', '&amp;', '&apos;', '&lt;' or '&gt;'")
                    End Select
                    xml_Index = xml_Index + 1
                    Exit Do
                Case Else
                    xml_Index = xml_Index + 1
                End Select
            Loop
            If Not xml_EncodedFound Then Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting ';'")
        Case Else
            xml_BufferAppend xml_Buffer, xml_Char, xml_BufferPosition, xml_BufferLength
            xml_Index = xml_Index + 1
        End Select
    Loop

    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '<'")
End Function

''
' Parse node 'text' to nodeValue. (i.e., String to Boolean, Double, Date).
'
' @method xml_ParseValue
' @param {String} xml_Text | Text to parse.
' @return {Variant} Node Value
''
Private Function xml_ParseValue(xml_Text As String) As Variant
    If xml_Text = "true" Then
        xml_ParseValue = True
    ElseIf xml_Text = "false" Then
        xml_ParseValue = False
    ElseIf xml_Text = "null" Then
        xml_ParseValue = Null
    ElseIf VBA.IsNumeric(xml_Text) Then
        xml_ParseValue = xml_ParseNumber(xml_Text)
    ElseIf VBA.IsNumeric(VBA.Replace(VBA.Left$(xml_Text, 10), "-", vbNullString)) And VBA.InStr(xml_Text, "T") And VBA.IIf(VBA.InStr(xml_Text, "Z"), VBA.Len(xml_Text) = 20, VBA.Len(xml_Text) = 19) Then
        xml_ParseValue = ParseIso(xml_Text)
    Else
        xml_ParseValue = xml_Text
    End If
End Function

Private Function xml_ParseNumber(xml_Text As String) As Variant
    Dim xml_Index As Long
    Dim xml_Char As String
    Dim xml_Value As String
    Dim xml_IsLargeNumber As Boolean
    Dim xml_IsGUID As Boolean
    Dim xml_IsISODate As Boolean
    
    xml_Index = 1
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_Text) + 1
        xml_Char = VBA.Mid$(xml_Text, xml_Index, 1)

        If VBA.InStr("+-0123456789.eE", xml_Char) And Not xml_Char = vbNullString Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            xml_Value = xml_Value & xml_Char
            xml_Index = xml_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            xml_IsLargeNumber = VBA.IIf(VBA.InStr(xml_Value, "."), VBA.Len(xml_Value) >= 17, VBA.Len(xml_Value) >= 16)
            If Not XmlOptions.UseDoubleForLargeNumbers And xml_IsLargeNumber Then
                xml_ParseNumber = xml_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                xml_ParseNumber = VBA.Val(xml_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function xml_IsVoidNode(xml_Node As Variant) As Boolean
    Select Case VBA.TypeName(xml_Node)
    Case "Dictionary"
        If xml_Node.Exists("childNodes") Then
            xml_IsVoidNode = VBA.IsNull(xml_Node.Item("nodeValue")) And xml_Node.Item("childNodes").Count = 0
        Else
            xml_IsVoidNode = VBA.IsNull(xml_Node.Item("nodeValue"))
        End If
    Case "IXMLDOMElement"
        xml_IsVoidNode = (xml_Node.ChildNodes.Length = 0 And xml_Node.Text = vbNullString)
    End Select
End Function

Private Function xml_Encode(xml_Text As Variant, Optional xml_QuoteChar As String = vbNullString) As String
    ' Variables.
    Dim xml_Index As Long
    Dim xml_Char As String
    Dim xml_AscCode As Long
    Dim xml_Buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    
    For xml_Index = 1 To VBA.Len(xml_Text)
        xml_Char = VBA.Mid$(xml_Text, xml_Index, 1)
        xml_AscCode = VBA.AscW(xml_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If xml_AscCode < 0 Then
            xml_AscCode = xml_AscCode + 65536
        End If

        ' From spec, <, >, &, ", ' characters must be modified.
        Select Case xml_AscCode
        Case 34
            ' " -> 34 -> &quot; | Only encode if attribute quote character is a quotation mark.
            If xml_QuoteChar = VBA.ChrW$(34) Then xml_Char = "&quot;"
        Case 38
            ' & -> 38 -> &amp;
            xml_Char = "&amp;"
        Case 39
            ' ' -> 39 -> &apos; | Only encode if attribute quote character is an apostrophe.
            If xml_QuoteChar = VBA.ChrW$(39) Then xml_Char = "&apos;"
        Case 60
            ' < -> 60 -> &lt;
            xml_Char = "&lt;"
        Case 62
            ' > -> 62 -> &gt;
            xml_Char = "&gt;"
        End Select
        
        xml_BufferAppend xml_Buffer, xml_Char, xml_BufferPosition, xml_BufferLength
    Next xml_Index
    
    xml_Encode = xml_BufferToString(xml_Buffer, xml_BufferPosition)
End Function

Private Sub xml_SkipSpaces(xml_String As String, ByRef xml_Index As Long)
    ' Increment index to skip over spaces
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String) And (VBA.Mid$(xml_String, xml_Index, 1) = " " Or VBA.Mid$(xml_String, xml_Index, 1) = vbCr Or VBA.Mid$(xml_String, xml_Index, 1) = vbLf)
        xml_Index = xml_Index + 1
    Loop
End Sub

Private Function xml_StringIsLargeNumber(xml_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See xml_ParseNumber)
    
    Dim xml_Length As Long
    Dim xml_CharIndex As Long
    xml_Length = VBA.Len(xml_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If xml_Length >= 16 And xml_Length <= 100 Then
        Dim xml_CharCode As String
        
        xml_StringIsLargeNumber = True
        
        For xml_CharIndex = 1 To xml_Length
            xml_CharCode = VBA.Asc(VBA.Mid$(xml_String, xml_CharIndex, 1))
            Select Case xml_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                xml_StringIsLargeNumber = False
                Exit Function
            End Select
        Next xml_CharIndex
    End If
End Function

Private Function xml_ParseErrorMessage(ByVal xml_String As String, ByVal xml_Index As Long, ByVal xml_ErrorMessage As String) As String
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing XML:
    ' <abc>1234</def>
    '          ^
    ' Expecting '</abc>'
    
    Dim xml_StartIndex As Long
    Dim xml_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    xml_StartIndex = xml_Index - 10
    xml_StopIndex = xml_Index + 10
    If xml_StartIndex <= 0 Then
        xml_StartIndex = 1
    End If
    If xml_StopIndex > VBA.Len(xml_String) Then
        xml_StopIndex = VBA.Len(xml_String)
    End If

    xml_ParseErrorMessage = "Error parsing XML:" & VBA.vbNewLine & _
        VBA.Mid$(xml_String, xml_StartIndex, xml_StopIndex - xml_StartIndex + 1) & VBA.vbNewLine & _
        VBA.Space$(xml_Index - xml_StartIndex) & "^" & VBA.vbNewLine & _
        xml_ErrorMessage
End Function

Private Sub xml_BufferAppend(ByRef xml_Buffer As String, _
                             ByRef xml_Append As Variant, _
                             ByRef xml_BufferPosition As Long, _
                             ByRef xml_BufferLength As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim xml_AppendLength As Long
    Dim xml_LengthPlusPosition As Long

    If Not xml_Append = vbNullString Then
        xml_AppendLength = VBA.Len(xml_Append)
        xml_LengthPlusPosition = xml_AppendLength + xml_BufferPosition
    
        If xml_LengthPlusPosition > xml_BufferLength Then
            ' Appending would overflow buffer, add chunk
            ' (double buffer length or append length, whichever is bigger)
            Dim xml_AddedLength As Long
            xml_AddedLength = VBA.IIf(xml_AppendLength > xml_BufferLength, xml_AppendLength, xml_BufferLength)
    
            xml_Buffer = xml_Buffer & VBA.Space$(xml_AddedLength)
            xml_BufferLength = xml_BufferLength + xml_AddedLength
        End If
    
        ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
        ' Function call on left-hand side of assignment must return Variant or Object
        Mid$(xml_Buffer, xml_BufferPosition + 1, xml_AppendLength) = CStr(xml_Append)
        xml_BufferPosition = xml_BufferPosition + xml_AppendLength
    End If
End Sub

Private Function xml_BufferToString(ByRef xml_Buffer As String, ByVal xml_BufferPosition As Long) As String
    If xml_BufferPosition > 0 Then
        xml_BufferToString = VBA.Left$(xml_Buffer, xml_BufferPosition)
    End If
End Function


