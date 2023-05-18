Attribute VB_Name = "modReg"
'//This code saves and reads registry keys using the
'//main registry NOT VB & VBA programs
'//Most of the code is mine but some is from other sources
'//Again Sorry it isn't commented but I always copy and
'//paste this module into my programs and I keep forgetting
'//to comment it

Public Const HKEY_CURRENT_USER = &H80000001

Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal Hkey As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias _
"RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As _
String, phkResult As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias _
"RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As _
String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal _
lpValueName As String, ByVal lpReserved As Long, lpType _
As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName _
As String, ByVal Reserved As Long, ByVal dwType As Long, _
lpData As Any, ByVal cbData As Long) As Long

Public Const REG_SZ = 1
Public Const REG_DWORD = 4


Public Function GetString(Hkey As Long, strPath As String, strValue As String)

Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(Hkey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

