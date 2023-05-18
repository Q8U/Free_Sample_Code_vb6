// ActiveDNS.js 1.0
// (c) 1998, Activity Software
//
// This is an example of using ActiveDNS within Windows Scripting
// Host.  It also serves to show how to implement ActiveDNS.DLL in
// JScript.
//
// This example uses 127.0.0.1 (loopback device), which should
// demonstrate the functionality of ActiveDNS in almost every scenario.
// 127.0.0.1 should always resolve to localhost, unless the machine
// has been configured otherwise.
//
// Requires: ActiveDNS.DLL, Windows Scripting Host
//
// To execute: CSCRIPT.EXE ACTIVEDNS.JS

hostname = "localhost";
address  = "127.0.0.1";

DNS=WScript.CreateObject("ActiveDNS.Resolve");

// Check for valid address
isvalid = DNS.IsValidNumericAddress(address);

// Lookup address by hostname
lookup1 = DNS.Lookup("localhost");

// Lookup hostname by address
lookup2 = DNS.ReverseLookup("127.0.0.1");

WScript.Echo("Address ["+address+"] "+ (isvalid ? "is" : "is not") + " valid.");
WScript.Echo("Host ["+hostname+"] resolves to address: "+lookup1);
WScript.Echo("Address ["+address+"] resolves to host: "+lookup2);
