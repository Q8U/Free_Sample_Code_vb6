﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.0.3705.0
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by wsdl, Version=1.0.3705.0.
// 
using System.Diagnostics;
using System.Xml.Serialization;
using System;
using System.Web.Services.Protocols;
using System.ComponentModel;
using System.Web.Services;


/// <remarks/>
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Web.Services.WebServiceBindingAttribute(Name="EchoClassSoap", Namespace="http://example.org/echo")]
public class EchoClass : System.Web.Services.Protocols.SoapHttpClientProtocol {
    
    /// <remarks/>
    public EchoClass() {
        string urlSetting = System.Configuration.ConfigurationSettings.AppSettings["EchoServiceLocation"];
        if ((urlSetting != null)) {
            this.Url = urlSetting;
        }
        else {
            this.Url = "http://localhost/echoservice/echo.asmx";
        }
    }
    
    /// <remarks/>
    [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://example.org/echo/Echo", RequestNamespace="http://example.org/echo", ResponseNamespace="http://example.org/echo", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
    public string Echo(string input) {
        object[] results = this.Invoke("Echo", new object[] {
                    input});
        return ((string)(results[0]));
    }
    
    /// <remarks/>
    public System.IAsyncResult BeginEcho(string input, System.AsyncCallback callback, object asyncState) {
        return this.BeginInvoke("Echo", new object[] {
                    input}, callback, asyncState);
    }
    
    /// <remarks/>
    public string EndEcho(System.IAsyncResult asyncResult) {
        object[] results = this.EndInvoke(asyncResult);
        return ((string)(results[0]));
    }
}
