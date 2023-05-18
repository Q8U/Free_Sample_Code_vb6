Attribute VB_Name = "modRegistry"
'***********************************************************************
' PURPOSE:     WRAPPER FUNCTIONS FOR CREATING/UPDATING/READING/DELETING
'              REGISTRY ENTRIES
' NOTES:       - YOUR PROJECT MUST HAVE A REFERENCE TO REGTool5.dll
'              (Registry Access Functions) IN ORDER TO USE THIS MODULE
'              - THE STOCK FUNCTIONS WILL NOT ACCEPT A VARIANT OF A STRING
'              SUBTYPE AS THE ValueData ARGUMENT, THUS PASSING IN CStr(MyValue)
'              DOES NOT WORK. THESE WRAPPER FUNCTIONS DO THE CONVERSION FOR
'              YOU SO YOU CAN PASS IN ANY TYPE TO MY WRAPPER FUNCTIONS
'***********************************************************************

Option Explicit

Public Enum RegistryRoot
   HKEY_CLASSES_ROOT = &H80000000      '-2147483648
   HKEY_CURRENT_USER = &H80000001      '-2147483647
   HKEY_LOCAL_MACHINE = &H80000002     '-2147483646
   HKEY_PERFORMANCE_DATA = &H80000004  '-2147483644
   HKEY_USERS = &H80000003             '-2147483645
End Enum


'***********************************************************************
' FUNCTION:    GetRegKey()
' PURPOSE:     WRAPPER FUNCTION FOR READING REGISTRY ENTRIES
' CREATED:     12/10/2001 / JASON BUTERA
' UPDATED:     12/10/2001 / JASON BUTERA
' NOTES:       BE SURE TO USE BACKSLASH (\) WHEN DEFINING THE KEY.
'              THE ValueData ARGUMENT IS PASSED BY REFERENCE AND WILL
'              HOLD THE REGISTRY VALUE AFTER THE FUNCTION IS RUN
' EXAMPLE:     GetRegKey(HKEY_CURRENT_USER, "Software\Example", "MyKey", varValue)
'              RETURNS TRUE IF KEY IS FOUND, FALSE IF NOT
'***********************************************************************
Public Function GetRegKey(KeyRoot As RegistryRoot, KeyName As String, ValueName As String, ByRef ValueData As Variant) As Boolean
   Dim strValueData As String
   strValueData = CStr(ValueData)
   GetRegKey = REGTool5.GetKeyValue(KeyRoot, KeyName, ValueName, strValueData)
   ValueData = strValueData
End Function


'***********************************************************************
' FUNCTION:    SetRegKey()
' PURPOSE:     WRAPPER FUNCTION FOR CREATING/UPDATING REGISTRY ENTRIES
' CREATED:     12/10/2001 / JASON BUTERA
' UPDATED:     12/10/2001 / JASON BUTERA
' NOTES:       BE SURE TO USE BACKSLASH (\) WHEN DEFINING THE KEY.
' EXAMPLE:     SetRegKey(HKEY_CURRENT_USER, "Software\Example", "MyKey", varValue)
'              RETURNS TRUE IF SUCCESSFUL, FALSE IF NOT
'***********************************************************************
Public Function SetRegKey(KeyRoot As RegistryRoot, KeyName As String, ValueName As String, ValueData As Variant) As Boolean
   Dim strValueData As String
   strValueData = CStr(ValueData)
   SetRegKey = REGTool5.UpdateKey(KeyRoot, KeyName, ValueName, strValueData)
   ValueData = strValueData
End Function


'***********************************************************************
' FUNCTION:    DeleteRegKey()
' PURPOSE:     WRAPPER FUNCTION FOR DELETING REGISTRY ENTRIES
' CREATED:     12/10/2001 / JASON BUTERA
' UPDATED:     12/10/2001 / JASON BUTERA
' NOTES:       BE SURE TO USE BACKSLASH (\) WHEN DEFINING THE KEY.
' EXAMPLE:     DeleteRegKey(HKEY_CURRENT_USER, "Software\Example")
'              RETURNS TRUE IF SUCCESSFUL, FALSE IF KEY NOT FOUND
'***********************************************************************
Public Function DeleteRegKey(KeyRoot As RegistryRoot, KeyName As String) As Boolean
   DeleteRegKey = REGTool5.DeleteKey(KeyRoot, KeyName)
End Function

