﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.0.3705.209
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'This source code was auto-generated by Microsoft.VSDesigner, Version 1.0.3705.209.
'
Namespace localhost
    
    '<remarks/>
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="DataSetServiceSoap", [Namespace]:="http://example.org/dataset-service")>  _
    Public Class DataSetService
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        '<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = "http://localhost/DataSetService/DataSetService.asmx"
        End Sub
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://example.org/dataset-service/GetAuthorsAsXml", RequestNamespace:="http://example.org/dataset-service", ResponseNamespace:="http://example.org/dataset-service", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetAuthorsAsXml() As GetAuthorsAsXmlResponseGetAuthorsAsXmlResult
            Dim results() As Object = Me.Invoke("GetAuthorsAsXml", New Object(-1) {})
            Return CType(results(0),GetAuthorsAsXmlResponseGetAuthorsAsXmlResult)
        End Function
        
        '<remarks/>
        Public Function BeginGetAuthorsAsXml(ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("GetAuthorsAsXml", New Object(-1) {}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndGetAuthorsAsXml(ByVal asyncResult As System.IAsyncResult) As GetAuthorsAsXmlResponseGetAuthorsAsXmlResult
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),GetAuthorsAsXmlResponseGetAuthorsAsXmlResult)
        End Function
    End Class
    
    '<remarks/>
    <System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://example.org/dataset-service")>  _
    Public Class GetAuthorsAsXmlResponseGetAuthorsAsXmlResult
        
        '<remarks/>
        <System.Xml.Serialization.XmlArrayAttribute([Namespace]:="http://example.org/dataset"),  _
         System.Xml.Serialization.XmlArrayItemAttribute("authors", [Namespace]:="http://example.org/dataset", IsNullable:=false)>  _
        Public AuthorSet() As AuthorType
    End Class
    
    '<remarks/>
    <System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://example.org/dataset")>  _
    Public Class AuthorType
        
        '<remarks/>
        Public au_id As String
        
        '<remarks/>
        Public au_lname As String
        
        '<remarks/>
        Public au_fname As String
        
        '<remarks/>
        Public phone As String
        
        '<remarks/>
        Public address As String
        
        '<remarks/>
        Public city As String
        
        '<remarks/>
        Public state As String
        
        '<remarks/>
        Public zip As String
        
        '<remarks/>
        Public contract As Boolean
    End Class
End Namespace
