Attribute VB_Name = "JsonXmlConverter"
''
' VBA-JsonXml
' (c) RadiusCore Ltd - https://www.radiuscore.co.nz/
'
' Convert between Json and Xml.
'
' @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php
' @depend JsonConverter, XmlConverter, UtcConverter
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

''
' Convert XML (Dictionary/Collection/DOMDocument/String) to JSON object or string.
'
' TODO - Convert document header.
' TODO - Convert XML attributes.
' TODO - Add support for DOMDocument (currently converts DOMDocument to VBA-XML Dictionary to convert).
'
' @method ConvertXmlToJson
' @param {Dictionary|Collection|DOMDocument|String} XmlValue | XML to convert to JSON.
' @param {Integer|String} Whitespace | "Pretty" print JSON with given number of spaces per indentation (Integer) or given string
' @param {VbVarType} ReturnAs | Whether to return a JSON string or object. `vbString` or `vbObject` only.
' @return {Dictionary|Collection|String} JSON object or string.
' @dependency VBA-JSON
' @dependency VBA-XML
''
Public Function ConvertXmlToJson(ByVal XmlValue As Variant, Optional ByVal Whitespace As Variant, Optional ReturnAs As VbVarType = vbString) As Variant
    Dim xml_ChildNode As Dictionary
    Dim xml_JsonObject As Dictionary
    Dim xml_JsonNodeName As String
    Dim xml_ReturnObject As Object
    
    If Not (ReturnAs = vbString Or ReturnAs = vbObject) Then
        Err.Raise 11001, "XMLConverter", "Invalid `ReturnAs` parameter. Must be `vbString` or `vbObject`."
    End If
    
    Select Case VBA.VarType(XmlValue)
    Case VBA.vbString
        Set xml_ReturnObject = ConvertXmlToJson(ParseXml(XmlValue), Whitespace, vbObject)
    Case VBA.vbObject
        ' Dictionary (Node).
        If VBA.TypeName(XmlValue) = "Dictionary" Then
            ' When converting from Json to XML, any `Key` with a space in it has the
            ' space replaced by an underscore. Revert this back to a space now.
            xml_JsonNodeName = VBA.Replace(XmlValue.Item("nodeName"), "_", " ")
            If xml_JsonNodeName = "#document" Then
                If Not XmlValue.Item("prolog") = vbNullString Then
                    ' TODO - Convert document headers.
                End If
                
                ' Set result to `ReturnObject`.
                Set xml_ReturnObject = ConvertXmlToJson(XmlValue.Item("childNodes"), Whitespace, vbObject)
            Else
                ' Validate Dictionary structure.
                If Not XmlValue.Exists("nodeName") Or Not XmlValue.Exists("nodeValue") Then
                    Err.Raise 11001, "XMLConverter", "Error parsing XML:" & VBA.vbNewLine & _
                        "Poorly structured XML Dictionary. Use `ParseXml` with `XmlOptions.ForceVbaXml = True` OR " & _
                        "`CreateNode` and `CreateAttribute` to create a correctly structured XML dictionary object."
                End If
                
                ' Convert XML node to Json object.
                Set xml_JsonObject = New Dictionary
                If XmlValue.Exists("childNodes") Then
                    If XmlValue.Item("childNodes").Count > 0 Then
                        xml_JsonObject.Add xml_JsonNodeName, ConvertXmlToJson(XmlValue.Item("childNodes"), Whitespace, vbObject)
                    Else
                        xml_JsonObject.Add xml_JsonNodeName, XmlValue.Item("nodeValue")
                    End If
                Else
                    xml_JsonObject.Add xml_JsonNodeName, XmlValue.Item("nodeValue")
                End If
                
                ' Set result to `ReturnObject`.
                Set xml_ReturnObject = xml_JsonObject
            End If
            
        ' Collection (child nodes)
        ElseIf VBA.TypeName(XmlValue) = "Collection" Then
            Set xml_JsonObject = New Dictionary
            
            ' Convert child nodes to VBA-JSON Dictionary/Collection structure.
            For Each xml_ChildNode In XmlValue
                ' When converting from Json to XML, any `Key` with a space in it has the
                ' space replaced by an underscore. Revert this back to a space now.
                xml_JsonNodeName = VBA.Replace(xml_ChildNode.Item("nodeName"), "_", " ")
                ' Add each `childNode` to a single dictionary with `Key=nodeName`.
                If xml_ChildNode.Item("childNodes").Count = 0 Then
                    ' No child nodes, add as `Key=nodeName` and `Value=nodeValue`.
                    xml_JsonObject.Add xml_JsonNodeName, xml_ChildNode.Item("nodeValue")
                Else
                    If Not xml_JsonObject.Exists(xml_JsonNodeName) Then
                        ' Add to Dictionary with `Key=nodeName` and `Value=childNodes(converted)`.
                        xml_JsonObject.Add xml_JsonNodeName, ConvertXmlToJson(xml_ChildNode.Item("childNodes"), Whitespace, vbObject)
                    Else
                        ' If `Key` already exists, there are at least two `childNodes` with the same `nodeName`.
                        ' In this case they should be grouped into a Collection(array) within that single
                        ' `Key`. So the `Value`(aka `Dictionary.Item`) in the `Key-Value` pair becomes
                        ' a Collection(array), and subsequent `childNodes` are added to this collection.
                        
                        ' If the `Value` is not a Collection(array), then convert to one and place the exsting
                        ' `Value` into the new Collection(array).
                        If Not TypeOf xml_JsonObject.Item(xml_JsonNodeName) Is Collection Then
                            ' Store existing `Value`.
                            Dim xml_Temp As Variant
                            If TypeOf xml_JsonObject.Item(xml_JsonNodeName) Is Dictionary Then
                                Set xml_Temp = xml_JsonObject.Item(xml_JsonNodeName)
                            Else
                                If VBA.IsNull(xml_JsonObject.Item(xml_JsonNodeName)) Then
                                    Set xml_Temp = New Dictionary
                                Else
                                    xml_Temp = xml_JsonObject.Item(xml_JsonNodeName)
                                End If
                            End If
                            ' Set `Value` to a Collection, add previous `Value` to new Collection.
                            Set xml_JsonObject.Item(xml_JsonNodeName) = New Collection
                            xml_JsonObject.Item(xml_JsonNodeName).Add xml_Temp
                            Set xml_Temp = Nothing
                        End If
                        
                        ' Add new `childNode` to existing Collection(array).
                        Dim xml_AddInPosition As Boolean
                        xml_AddInPosition = False
                        ' If an array doesn't have an approriate key when converting from JSON to XML, nodes are created with
                        ' `nodeName=element` and have the first attribute set as `id` which notates the order of these elements.
                        ' This will check for any nodes that match these requirements (`nodeName=element` & `attribute(1)("Name")=id`)
                        ' and add them back to the JSON Collection(array) in the correct order.
                        If Not xml_ChildNode.Item("attributes") Is Nothing Then
                            If xml_JsonNodeName = "element" And xml_ChildNode.Item("attributes").Item(1).Item("name") = "id" Then
                                If xml_JsonObject.Item(xml_JsonNodeName).Count >= VBA.CLng(xml_ChildNode.Item("attributes").Item(1).Item("value")) Then
                                    xml_AddInPosition = True
                                    xml_JsonObject.Item(xml_JsonNodeName).Add Item:=ConvertXmlToJson(xml_ChildNode.Item("childNodes"), Whitespace, vbObject), Before:=VBA.CLng(xml_ChildNode.Item("attributes").Item(1).Item("value"))
                                End If
                            End If
                        End If
                        If Not xml_AddInPosition Then
                            ' Sequentially add `Value` to JSON array.
                            xml_JsonObject.Item(xml_JsonNodeName).Add ConvertXmlToJson(xml_ChildNode.Item("childNodes"), Whitespace, vbObject)
                        End If
                    End If
                End If
            Next xml_ChildNode
            
            ' Set result to `ReturnObject`.
            If xml_JsonObject.Exists("element") And xml_JsonObject.Count = 1 Then
                ' If an array doesn't have an approriate key when converting from JSON to XML, nodes are created with
                ' `nodeName=element` and have the first attribute set as `id` which notates the order of these elements.
                ' Check whether the JSON Object matches this description (`Key=element`), and if it is the only `Key` in the
                ' JSON Object, then only return the `Value`. This removes the `Key`, correctly reverting the JSON object
                ' back to its original form.
                Set xml_ReturnObject = xml_JsonObject.Item("element")
            Else
                Set xml_ReturnObject = xml_JsonObject
            End If
            
        ' MSXML2.DOMDocument (windows only)
        ElseIf VBA.TypeName(XmlValue) = "DOMDocument" Then
            ' TODO - Add support for DOMDocument.
            ' <--- START TEMP - Convert DOMDocument to VBA-XML Dictionary, then convert to Json.
            Dim xml_ForceVBA As Boolean
            Dim xml_Object As Dictionary
            xml_ForceVBA = XmlConverter.XmlOptions.ForceVbaXml
            XmlConverter.XmlOptions.ForceVbaXml = True
                Set xml_Object = XmlConverter.ParseXml(XmlValue.Xml)
            XmlConverter.XmlOptions.ForceVbaXml = xml_ForceVBA
            Set xml_ReturnObject = ConvertXmlToJson(xml_Object, Whitespace, vbObject)
            ' <--- End TEMP
        End If
    End Select
    
    ' Return JSON object or string as required.
    If ReturnAs = vbObject Then
        Set ConvertXmlToJson = xml_ReturnObject
    Else
        ConvertXmlToJson = ConvertToJson(xml_ReturnObject, Whitespace)
    End If
End Function

''
' Convert JSON (Dictionary/Collection/String) to XML object or string.
'
' TODO - Write function.
'
' @method ConvertJsonToXml
' @param {Dictionary|Collection|String} JsonValue | JSON value to convert to XML.
' @param {Integer|String} Whitespace | "Pretty" print XML with given number of spaces per indentation (Integer) or given string.
' @param {VbVarType} ReturnAs | Whether to return a XML String or Object. `vbString` or `vbObject` only.
' @return {Dictionary|Collection|DOMDocument|String} XML object or string.
' @dependency VBA-JSON
' @dependency VBA-XML
''
Public Function ConvertJsonToXml(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal ReturnAs As VbVarType = vbString) As Variant
    Dim json_ReturnObject As Dictionary
    
    If Not (ReturnAs = vbString Or ReturnAs = vbObject) Then
        Err.Raise 11001, "JSONConverter", "Invalid `ReturnAs` parameter. Must be `vbString` or `vbObject`."
    End If
    
    Select Case VarType(JsonValue)
    Case vbString
        Set json_ReturnObject = ConvertJsonToXml(ParseJson(JsonValue), Whitespace, vbObject)
    Case VBA.vbObject
        If VBA.TypeName(JsonValue) = "Dictionary" Then
            ' TODO - proccess Dictionary.
        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            ' TODO - proccess Collection.
        End If
    End Select
    
    ' Return XML object or string as required.
    If ReturnAs = vbObject Then
        Set ConvertJsonToXml = json_ReturnObject
    Else
        ConvertJsonToXml = ConvertToXml(json_ReturnObject, Whitespace)
    End If
End Function

