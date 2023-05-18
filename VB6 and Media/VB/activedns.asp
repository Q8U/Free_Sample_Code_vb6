<%@ LANGUAGE="VBSCRIPT" %>

<%

' ActiveDNS.asp 1.0
' (c) 1998, Activity Software 
'
' Example of ActiveDNS using Active Server Pages.  This can be used
' on any server which ActiveDNS.DLL is registered.
' 
' This example is in VBScript.  For an example in JScript, please
' refer to ActiveDNS.js, which is designed for Microsoft Windows
' Scripting Host.
'

Set dns = Server.CreateObject("ActiveDNS.Resolve")

Address = "127.0.0.1"
Hostname = "localhost"

If dns.IsValidNumericAddress(Address) Then
 Response.Write Address & " Is a valid numeric address.<BR>"
Else
 Response.Write Address & " Is not a valid numeric address.<BR>"
End If

Response.Write "Host: " & Hostname & " resolves to address: " & _
                            dns.Lookup(Hostname) & "<BR>"

Response.Write "Address: " & Address & " resolves to host: " & _ 
                            dns.ReverseLookup(Address)  & "<BR>"

%>