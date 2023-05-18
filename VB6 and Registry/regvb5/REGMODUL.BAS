Attribute VB_Name = "REGMODUL"
' --------------------------------------------------------
' This .BAS module contains functions for reading and
'   setting registry values in Windows 95 and Windows NT.
'   The demo is in the form of a functional registry editor,
'   although it is not meant to be a substitute for RegEdit
'   or RegEdt32.  Declares are not provided
'   for some rare registry functions, and not all functions
'   have demo usage implemented.  Only values of type REG_SZ
'   and REG_DWORD may be edited.
'
' Written 10/95 for VB4/32 by Don Bradner, modified 6/97 to
'   deal with VB5 incompatibilities.  Requirements to run include
'   VB5, Windows 95 or NT 3.51 or later, and Comctl32.OCX,
'   which comes with VB5 Pro or EE versions.
'
' This material is placed in the public domain.  No guarantees
'   are made, and no support is provided, but comments/bug
'   reports are welcome to Don Bradner at Compuserve 76130,1007
'   or at dbirdman@arcatapet.com.  Questions/comments are also
'   welcome in the 32-bit section of the Visual Basic Programmer's
'   Journal forum on Compuserve.
'
' Warning:  Editing registry values can seriously impact your
'   computer's operations.  You should only edit values when
'   you know what they should be.  If editing values as a
'   test, make a note of the original value and restore it
'   when you are done.
' --------------------------------------------------------

Option Explicit

' --------------------------------------------------------
' FILETIME type is needed for RegEnumKey and
'   RegQueryInfoKey
' --------------------------------------------------------
Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type

' --------------------------------------------------------
' OsVersionInfo type is needed for GetVersionEx
' --------------------------------------------------------
Type OsVersionInfo
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type

' ---------------------------
' 32-bit registry functions
' ---------------------------
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal dwRes&, ByVal dwType&, lpDataBuff As Any, ByVal nSize&)
Declare Function RegConnectRegistry& Lib "advapi32.dll" (ByVal lpMachineName$, ByVal hKey&, phkResult&)
Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, lpSecurityAttributes&, phkResult&, lpdwDisposition&)
Declare Function RegFlushKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegEnumKeyEx& Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey&, ByVal dwIndex&, ByVal lpname$, lpcbName&, ByVal lpReserved&, ByVal lpClass$, lpcbClass&, lpftLastWriteTime As FILETIME)

Declare Function RegEnumValue& Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpname$, _
    lpcbName&, ByVal lpReserved&, lpdwType&, lpValue As Any, lpcbValue&)

Declare Function RegQueryInfoKey& Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, lpcValues&, lpcbMaxValueNameLen&, lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime As FILETIME)

Declare Function GetVersionEx& Lib "kernel32.dll" Alias "GetVersionExA" (lpStruct As OsVersionInfo)

' --------------------------------------------------------
' Functions used to show hourglass
' --------------------------------------------------------
Declare Function LoadCursor& Lib "User32" Alias "LoadCursorA" (ByVal hInstance&, ByVal lpCursor&)
Declare Function SetCursor& Lib "User32" (ByVal hCursor&)
Public Const IDC_WAIT = 32514&

Public iWaitCursor&

Public lNewKey& 'used to generate unique Node keys

' --------------------------------------------------------
' Return codes from Registration functions.
' --------------------------------------------------------
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
   
Public Const LB_SETHORIZONTALEXTENT = &H400 + 21

Const REG_NONE = 0&                        ' No value type
Public Const REG_SZ = 1&                   ' Unicode nul terminated string
Const REG_EXPAND_SZ = 2&                   ' Unicode nul terminated string
                                           ' (with environment variable references)
Const REG_BINARY = 3&                      ' Free form binary
Public Const REG_DWORD = 4&                ' 32-bit number
Const REG_DWORD_LITTLE_ENDIAN = 4&         ' 32-bit number (same as REG_DWORD)
Const REG_DWORD_BIG_ENDIAN = 5&            ' 32-bit number
Const REG_LINK = 6&                        ' Symbolic Link (unicode)
Const REG_MULTI_SZ = 7&                    ' Multiple Unicode strings
Const REG_RESOURCE_LIST = 8&               ' Resource list in the resource map
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&    ' Resource list in the hardware description
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&


' --------------------------------------------------------
' Read/Write permissions:
' --------------------------------------------------------
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ


' --------------------------------------------------------
' Used with GetVersionEX
' --------------------------------------------------------
Public iWinVers%
Public Const WinNt = 1
Public Const Win32 = 2

Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2


' --------------------------------------------------------
' Public program variables
' --------------------------------------------------------
Public lTempLong&
Public fTempDbl#
Public sTempString$
Public nodetemp As Node


' registry entry important stuff
Public rgeEntry$
Public rgeDataType&
Public rgeValue$
Public rgeMainKey&
Public rgeSubKey$


Public Sub rgeExtractKeys()
    '------------------------------------------------------------------
    'Enter with rgeSubKey containing a full path in the
    'My Computer\HKEY_...\..\ format or HKEY_...\...\ format
    '--------------------------------------------------------------
    If Left$(rgeSubKey, 12) = "My Computer\" Then
        rgeSubKey = Right$(rgeSubKey, Len(rgeSubKey) - 12)
    End If
    
    '----------------------------------------------------------------
    'If there is no main key we have to assume it may already have
    'been extracted.
    '---------------------------------------------------------------
    If Left$(rgeSubKey, 5) <> "HKEY_" Then
        Exit Sub
    ElseIf InStr(rgeSubKey, "\") = 0 Then
        rgeMainKey = GetMainKey(rgeSubKey)
        rgeSubKey = ""
    Else
        rgeMainKey = GetMainKey(Left$(rgeSubKey, InStr(rgeSubKey, "\") - 1))
        rgeSubKey = Right$(rgeSubKey, Len(rgeSubKey) - InStr(rgeSubKey, "\"))
    End If
    ' -----------------------------------------------------
    ' Check rgeMainKey& for validity
    ' -----------------------------------------------------
    If rgeMainKey < &H80000000 Or rgeMainKey > &H80000006 Then
        rgeClear
    End If
        

End Sub
Public Sub rgeClear()
    rgeMainKey = 0
    rgeSubKey = ""
    rgeValue = ""
    rgeDataType = 0
    rgeEntry = ""
End Sub



Sub EditRegValue(ByVal nodX As Node, lRegIndex&)
    ' --------------------------------------------------------
    ' Specific to the RegDemo application.
    ' --------------------------------------------------------
    rgeSubKey = nodX.FullPath
    rgeSubKey = Right$(rgeSubKey, Len(rgeSubKey) - 12)
    If InStr(rgeSubKey, "\") = 0 Then
        rgeMainKey = GetMainKey(rgeSubKey)
        rgeSubKey = ""
    Else
        '-------------------------------------------------
        'This must be a SubKey.
        '-------------------------------------------------
        rgeMainKey = GetMainKey(Left$(rgeSubKey, InStr(rgeSubKey, "\") - 1))
        rgeSubKey = Right$(rgeSubKey, Len(rgeSubKey) - InStr(rgeSubKey, "\"))
    End If

    Dim lRtn&       ' Returned by registry functions, should be 0&
    Dim hKey&       ' Return handle to opened key
    Dim lLenValueName&
    Dim lLenValue&
    Dim lKeyIndx&
   
    ' --------------------------------------------------------
    ' values for QueryInfoKey:
    ' --------------------------------------------------------
    Dim sClassName$
    Dim lClassLen&
    Dim lSubKeys&
    Dim lMaxSubKey&
    Dim lMaxClass&
    Dim lValues&
    Dim lMaxValueName&
    Dim lMaxValueData&
    Dim lSecurityDesc&
    Dim strucLastWriteTime As FILETIME
    
    ' -----------------------------------------------------
    ' Open key
    ' -----------------------------------------------------
    lRtn = RegOpenKeyEx(rgeMainKey, rgeSubKey, 0&, KEY_READ, hKey)
    If lRtn <> ERROR_SUCCESS Then
        MsgBox RtnRegError(lRtn)
      
        ' --------------------------------------------------
        ' No key open, so leave
        ' --------------------------------------------------
        Exit Sub
    End If
    
    ' -----------------------------------------------------
    ' RegQueryInfoKey is used to get the size of the largest
    '   value name and data string.
    ' Other returned values are ignored.
    ' -----------------------------------------------------
    sClassName = Space$(255) 'initialize these because occasional errors otherwise
    lClassLen = CLng(Len(sClassName))
    lRtn = RegQueryInfoKey(hKey, sClassName, lClassLen, 0&, lSubKeys, lMaxSubKey, lMaxClass, lValues, lMaxValueName, lMaxValueData, lSecurityDesc, strucLastWriteTime)

    '-------------------------------------------------------------------
    'If the enumeration fails due to a buffer over-run, we will loop back
    'to this point with larger buffers.
    '-------------------------------------------------------------------
RetryValueHere:
    
    ' --------------------------------------------------
    ' Set variables
    ' --------------------------------------------------
    rgeEntry = Space$(lMaxValueName + 1)
    lLenValueName = CLng(Len(rgeEntry)) '+ 1
    rgeValue = Space$(lMaxValueData + 1)
    lLenValue = CLng(Len(rgeValue))       '+ 1
    
    ' --------------------------------------------------
    ' Call the enumeration function to get the indexed value
    ' --------------------------------------------------
    lRtn = RegEnumValue(hKey, lRegIndex, rgeEntry, lLenValueName, 0&, rgeDataType, ByVal rgeValue, lLenValue)
    
    
    ' --------------------------------------------------
    ' Check for success
    ' --------------------------------------------------
    If lRtn = ERROR_SUCCESS Then
        If rgeDataType <> REG_SZ And rgeDataType <> REG_DWORD Then
            
            '--------------------------------------------------------------
            'Tell us what value types may be edited, along with
            '    the type of value found.
            '--------------------------------------------------------------
            sTempString = "This Demo only supports editing of values with types of REG_SZ and REG_DWORD.  This value is of type "
            Select Case rgeDataType
                Case 2
                    sTempString = sTempString & "REG_EXPAND_SZ."
                Case 3
                    sTempString = sTempString & "REG_BINARY."
                Case 5
                    sTempString = sTempString & "REG_DWORD_BIG_ENDIAN."
                Case 6
                    sTempString = sTempString & "REG_LINK."
                Case 7
                    sTempString = sTempString & "REG_MULTI_SZ."
                Case 8
                    sTempString = sTempString & "REG_RESOURCE_LIST."
                Case 9
                    sTempString = sTempString & "REG_FULL_RESOURCE_DESCRIPTOR."
                Case 10
                    sTempString = sTempString & "REG_RESOURCE_REQUIREMENTS_LIST."
            End Select
            MsgBox sTempString

        Else
            rgeEntry = Mid$(rgeEntry, 1, lLenValueName)
            If lLenValueName = 0 Then
                rgeEntry = "(Default)"
            End If
            rgeValue = Mid$(rgeValue, 1, lLenValue)
            ' --------------------------------------------
            ' Convert DWORD 4 character value to 32-bit
            '   number.
            ' First character is low byte, and so on.
            ' --------------------------------------------
            Form2.Caption = "Edit String Value"
            If rgeDataType = REG_DWORD Then
                fTempDbl = Asc(Mid$(rgeValue, 1, 1)) + &H100& * Asc(Mid$(rgeValue, 2, 1)) + &H10000 * Asc(Mid$(rgeValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(rgeValue, 4, 1)))
                If fTempDbl > &H7FFFFFFF Then
                    rgeValue = Hex$(fTempDbl - 4294967296#)
                Else
                    rgeValue = Hex$(fTempDbl)
                End If
                ' -----------------------------------------
                ' Turn on Hex/Decimal options
                ' -----------------------------------------
                Form2.Frame1.Visible = True
                Form2.Caption = "Edit DWORD Value"
            End If
            
            ' --------------------------------------------
            ' Place the values in the form2 text boxes
            ' --------------------------------------------
            Form2.Text1 = rgeEntry
            Form2.Text2 = rgeValue
        
            ' --------------------------------------------
            ' Pass the key variables to form2 via hidden
            '   text boxes
            ' --------------------------------------------
            Form2.Show 1
        End If

    ElseIf lRtn = ERROR_MORE_DATA Then
        ' -----------------------------------------------
        ' This error means that, despite querying the key
        '   we have not set one of the buffers large
        '   enough. If the buffer is already 20000 we are
        '   not going to be able to edit it.
        ' -----------------------------------------------
        If lMaxValueData >= 20000 Then
            MsgBox ("Value is too large for this editor!")
        Else
        
            ' --------------------------------------------
            ' Increase the buffer sizes and try again
            ' --------------------------------------------
            lMaxValueData = lMaxValueData + 5
            lMaxValueName = lMaxValueName + 5
            GoTo RetryValueHere
        End If
    Else
        
        ' --------------------------------------------------
        ' Key still open, so display the error and fall
        '   thru to the close function below
        ' --------------------------------------------------
        MsgBox RtnRegError(lRtn)
    End If

    ' -----------------------------------------------------
    ' Always close opened keys!
    ' -----------------------------------------------------
    lRtn = RegCloseKey(hKey)

End Sub

Function truncatepath$(TPath$, TForm As Form, TControl As Control)
Dim FirstSlash%
    On Error GoTo TPathError
    TPath = Right$(TPath, Len(TPath) - InStr(TPath, "H") + 1)
    FirstSlash = InStr(3, TPath, "\")
    Do Until TForm.TextWidth(TPath) <= TControl.Width
        lTempLong = InStr(FirstSlash + 5, TPath, "\")
        TPath = Left$(TPath, FirstSlash) & "..." & Mid$(TPath, lTempLong, Len(TPath) - lTempLong + 1)
    Loop
TooLong:
On Error GoTo StillError
If TForm.TextWidth(TPath) > TControl.Width Then
    Do Until TForm.TextWidth(TPath) <= TControl.Width
        TPath = Left$(TPath, Len(TPath) - 1)
    Loop
    TPath = Left$(TPath, Len(TPath) - 3) & "..."
End If
FinishAnyWay:
    truncatepath = TPath
    Exit Function
TPathError:
    Resume TooLong
StillError:
    Resume FinishAnyWay
End Function

Function GetMainKey&(keyname$)
   
   ' -----------------------------------------------------
   ' Used to convert main key strings to their values
   ' -----------------------------------------------------

    Select Case keyname
        Case "HKEY_CLASSES_ROOT"
            GetMainKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            GetMainKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            GetMainKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            GetMainKey = HKEY_USERS
        Case "HKEY_PERFORMANCE_DATA"
            GetMainKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_CURRENT_CONFIG"
            GetMainKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            GetMainKey = HKEY_DYN_DATA
    End Select

End Function


Function RegEnumKeys&(ByVal Node As Node, bFullEnumeration As Boolean)
    lTempLong = SetCursor(iWaitCursor)
    Dim sRoot$, sRoot2$
    rgeSubKey = Node.FullPath
    If rgeSubKey = "My Computer" Then
        Exit Function
    End If
    '---------------------------------------------
    'If we've put in a fake node to set the + image,
    'remove that node to avoid duplication
    '---------------------------------------------
    While Node.Children > 0
        Form1.TreeView1.Nodes.Remove Node.Child.Key
    Wend
    Form1.TreeView1.Enabled = False
    rgeExtractKeys
    sRoot = Node.Key
    ' --------------------------------------------------------
    ' This function will load all subkeys into the TreeView
    ' --------------------------------------------------------
    Dim lRtn&       ' Returned by registry functions, should be 0&
    Dim hKey&       ' Return handle to opened key
    Dim strucLastWriteTime    As FILETIME
    Dim sSubKeyName$
    Dim sClassString$
    Dim lLenSubKey&
    Dim lLenClass&
    Dim lKeyIndx&
    Dim lRet&
    Dim hKey2&
    Dim sSubKey2$
    Dim nodX As Node
    Dim sNewKey$
    
    '---------------------------------------------
    'values for QueryInfoKey:
    '---------------------------------------------
    Dim sClassName$
    Dim lClassLen&
    Dim lSubKeys&
    Dim lMaxSubKey&
    Dim sMaxSubKey$
    Dim lMaxClass&
    Dim sMaxClass$
    Dim lValues&
    Dim lMaxValueName&
    Dim lMaxValueData&
    Dim lSecurityDesc&
    
    ' -----------------------------------------------------
    ' Open key
    ' -----------------------------------------------------
    lRtn = RegOpenKeyEx(rgeMainKey, rgeSubKey, 0&, KEY_READ, hKey)
    If lRtn <> ERROR_SUCCESS Then
        If lRtn = ERROR_ACCESS_DENIED Then
            '---------------------------------------------
            'Grey the key
            'otherwise report error condition
            '---------------------------------------------
            Node.Image = 6
            'Node.Enabled = False 'Doesn't work?
        Else
            MsgBox RtnRegError(lRtn)
        End If
        RegEnumKeys = lRtn
        Form1.TreeView1.Enabled = True
        Exit Function
        
        ' --------------------------------------------------
        ' No key open, so leave
        ' --------------------------------------------------
    End If
    
    ' -----------------------------------------------------
    ' A call to RegQueryInfoKey will tell us the maximum
    '   keyname length
    ' -----------------------------------------------------
    sClassName = Space$(255)
    lClassLen = CLng(Len(sClassName))
    lRet = RegQueryInfoKey(hKey, sClassName, lClassLen, 0&, lSubKeys, lMaxSubKey, lMaxClass, lValues, lMaxValueName, lMaxValueData, lSecurityDesc, strucLastWriteTime)
    sMaxSubKey = Space$(lMaxSubKey + 1)
    sMaxClass = Space$(lMaxClass + 1)

    ' -----------------------------------------------------
    ' Enumerate the keys
    ' -----------------------------------------------------
    lKeyIndx = 0&
    Do While lRtn = ERROR_SUCCESS
        
        ' -----------------------------------------------------
        ' If the enumeration fails due to a buffer over-run,
        '   we will loop back to this point with larger buffers.
        ' -----------------------------------------------------
ReTryKeyEnumeration:
            
        ' --------------------------------------------------
        ' Set variables
        ' --------------------------------------------------
        sSubKeyName = sMaxSubKey
        lLenSubKey = lMaxSubKey
        sClassString = sMaxClass
        lLenClass = lMaxClass
    
        
        ' --------------------------------------------------
        ' Call the enumeration function
        ' --------------------------------------------------
        lRtn = RegEnumKeyEx(hKey, lKeyIndx, sSubKeyName, lLenSubKey, 0&, sClassString, lLenClass, strucLastWriteTime)
        If InStr(sSubKeyName, Chr$(0)) > 1 Then
            sSubKeyName = Left$(sSubKeyName, InStr(sSubKeyName, Chr$(0)) - 1)
        End If
        
        
        ' --------------------------------------------------
        ' Check for success
        ' --------------------------------------------------
        If lRtn = ERROR_SUCCESS Then
            lNewKey = lNewKey + 1
            sNewKey = "A" & Format$(lNewKey, "000000")
            Set nodX = Form1.TreeView1.Nodes.Add(sRoot, tvwChild, sNewKey, sSubKeyName, 1)
            If bFullEnumeration = True Then
                sSubKey2 = sSubKeyName
                If rgeSubKey <> "" Then
                    sSubKey2 = Trim(rgeSubKey) & "\" & sSubKeyName
                 End If
    
                ' -----------------------------------------------
                ' Use RegQueryInfoKey to find out if this key has
                '   subkeys
                ' -----------------------------------------------
                lRet = RegOpenKeyEx(rgeMainKey, sSubKey2, 0&, KEY_READ, hKey2)
                '------------------------------------------------------
                'We are fully enumerating a key, so set images and
                'Recurse a single SubKey to set + indicator if there are
                'subkeys below this one
                '------------------------------------------------------
                If lRet = ERROR_ACCESS_DENIED Then
                    nodX.ExpandedImage = 6
                    nodX.SelectedImage = 6
                ElseIf lRet = ERROR_SUCCESS Then
                    nodX.ExpandedImage = 2
                    nodX.SelectedImage = 2
                    lRet = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, lSubKeys, 0&, 0&, 0&, 0&, 0&, 0&, strucLastWriteTime)
                    ' --------------------------------------------------
                    ' Check for success.  If lSubKeys is greater than zero
                    ' there are subkeys for this key, and we will set a fake
                    ' node under this one to make a + symbol.
                    ' --------------------------------------------------
                    If lRet = ERROR_SUCCESS And lSubKeys > 0 Then
                        sRoot2 = nodX.Key
                        lNewKey = lNewKey + 1
                        sNewKey = "A" & Format$(lNewKey, "000000")
                        Set nodX = Form1.TreeView1.Nodes.Add(sRoot2, tvwChild, sNewKey, "PlaceHolder", 1)
                    End If
                    lRet = RegCloseKey(hKey2)
                End If
            Else
                Exit Do
            End If
            lKeyIndx = lKeyIndx + 1
        ElseIf lRtn = ERROR_MORE_DATA Then
            ' -----------------------------------------------
            ' This error means that, despite querying the key
            '   we have not set one of the buffers large
            '   enough.Increment the buffer sizes and try
            '   again
            ' -----------------------------------------------
            lMaxSubKey = lMaxSubKey + 5
            lMaxClass = lMaxClass + 5
            sMaxSubKey = Space$(lMaxSubKey + 1)
            sMaxClass = Space$(lMaxClass + 1)
            GoTo ReTryKeyEnumeration
        ElseIf lRtn = ERROR_NO_MORE_ITEMS Then
            ' -----------------------------------------------
            ' Not an error, just end of list -- exit the
            '   loop
            ' -----------------------------------------------
            lRtn = ERROR_SUCCESS
            Exit Do
        ElseIf lRtn <> ERROR_SUCCESS Then
            ' --------------------------------------------------
            ' Key still open, so display the error and fall
            '   thru to the close function below
            ' --------------------------------------------------
            MsgBox RtnRegError(lRtn)
            Exit Do
        End If
    Loop

    
    ' -----------------------------------------------------
    ' Set the return to the last error
    ' -----------------------------------------------------
    RegEnumKeys = lRtn

    ' -----------------------------------------------------
    ' Always close opened keys!
    ' -----------------------------------------------------
    lRtn = RegCloseKey(hKey)
    Form1.TreeView1.Enabled = True
End Function
Public Sub RegEnumValues()
    ' --------------------------------------------------------
    'Enter with rgeSubKey containing a full key path, in
    'My Computer\HKEY_..\..\ fashion
    ' --------------------------------------------------------
    'When you don't need to enumerate all values, but just want to
    'retrive a single value, use the provided function RegGetValue
    '---------------------------------------------------------
    Dim lRtn&        ' Returned by registry functions, should be 0&
    Dim hKey&       ' Return handle to opened key
    Dim lLenValueName&
    Dim lLenValue&
    Dim lKeyIndx&
    Dim sBinaryString$
    Dim Item As ListItem
    Dim iTempInt%
    '---------------------------
    'values for QueryInfoKey:
    '---------------------------
    Dim sClassName$
    Dim lClassLen&
    Dim lSubKeys&
    Dim lMaxSubKey&
    Dim lMaxClass&
    Dim lValues&
    Dim lMaxValueName&
    Dim lMaxValueData&
    Dim lSecurityDesc&
    Dim strucLastWriteTime As FILETIME
    Dim pbytValueName() As Byte
    Dim pbytValue() As Byte
    Dim DataType&
    Dim ValueName$
    Dim ValueVal$
    
    Dim iListWidth%  'Used to set listbox scrollbar
    
    lTempLong = SetCursor(iWaitCursor)
    
    rgeExtractKeys
    
    ' -----------------------------------------------------
    ' Open key
    ' -----------------------------------------------------
    lRtn = RegOpenKeyEx(rgeMainKey, rgeSubKey, 0&, KEY_READ, hKey)
    If lRtn <> ERROR_SUCCESS Then
        If lRtn <> ERROR_ACCESS_DENIED Then
            '---------------------------------------------
            'If access is denied don't do anything
            '---------------------------------------------
            MsgBox RtnRegError(lRtn)
        End If
        rgeClear
        
        ' --------------------------------------------------
        ' No key open, so leave
        ' --------------------------------------------------
        Exit Sub
    End If
   
   ' -----------------------------------------------------
   ' Use RegQueryInfoKey to get the maximum value data info.
   ' -----------------------------------------------------
    sClassName = Space$(255)
    lClassLen = CLng(Len(sClassName))
    lRtn = RegQueryInfoKey(hKey, sClassName, lClassLen, 0&, lSubKeys, lMaxSubKey, lMaxClass, lValues, lMaxValueName, lMaxValueData, lSecurityDesc, strucLastWriteTime)
   

    ' -----------------------------------------------------
    ' Enumerate the keys
    ' -----------------------------------------------------
    lKeyIndx = 0&
    Do While lRtn = ERROR_SUCCESS And lKeyIndx < lValues
        sBinaryString = ""

'-------------------------------------------------------------------
'If the enumeration fails due to a buffer over-run, we will loop back
'to this point with larger buffers.
'-------------------------------------------------------------------
ReTryValueEnumeration:
        
        ' --------------------------------------------------
        ' Set variables
        ' --------------------------------------------------
        ValueName = Space$(lMaxValueName + 1)
        lLenValueName = CLng(Len(ValueName)) '+ 1
        ValueVal = Space$(lMaxValueData + 1)
        lLenValue = CLng(Len(ValueVal))       '+ 1
        
        ' --------------------------------------------------
        ' Call the enumeration function
        ' --------------------------------------------------
        'lRtn = RegEnumValue(hKey, lKeyIndx, pbytValueName(0), _
        'lLenValueName, vbNull, rgeDataType, pbytValue(0), _
        'lLenValue)
        lRtn = RegEnumValue(hKey, lKeyIndx, ValueName, _
        lLenValueName, 0&, DataType, ByVal ValueVal, _
        lLenValue)
        
       
        ' --------------------------------------------------
        ' Check for success
        ' --------------------------------------------------
        If lRtn = ERROR_SUCCESS Then
            
            rgeEntry = ValueName
            rgeDataType = DataType
            rgeValue = ValueVal
            'rgeValue = pbytValue()
            '-----------------------------------------------
            'Add an item to the list box
            '-----------------------------------------------
            Set Item = Form1.ListView1.ListItems.Add()
            
            ' -----------------------------------------------
            ' Start building the entry to put in the list box:
            ' -----------------------------------------------
            rgeEntry = Mid$(rgeEntry, 1, lLenValueName)
        
            ' -----------------------------------------------
            ' Default values don't have a name.
            ' -----------------------------------------------
            If lLenValueName = 0 Then
                rgeEntry = "(Default)"
            End If
            
            rgeValue = Mid$(rgeValue, 1, lLenValue)
            Select Case rgeDataType
                Case REG_MULTI_SZ
                    ' --------------------------------------
                    ' REG_MULTI_SZ strings are a series of
                    '   zero terminated strings. If we don't
                    '   strip out the zeros, only the first
                    '   one will display.
                    ' We will replace them with spaces.
                    ' --------------------------------------
                    Item.SmallIcon = 4
                    Do While InStr(rgeValue, Chr$(0))
                        rgeValue = Left$(rgeValue, InStr(rgeValue, Chr$(0)) - 1) & " " & Right$(rgeValue, Len(rgeValue) - InStr(rgeValue, Chr$(0)))
                    Loop
                Case REG_SZ
                    ' --------------------------------------
                    ' REG_SZ values are zero-terminated
                    '   strings, and are the most common
                    '   values.
                    ' --------------------------------------
                    Item.SmallIcon = 4
                    
                    '---------------------------------------
                    'Put quotes around the string
                    '---------------------------------------
                    rgeValue = """" & Left$(rgeValue, lLenValue - 1) & """"
                
                Case REG_EXPAND_SZ
                    '---------------------------------------
                    'Environmental variables that are binary
                    'but evaluate as strings.  Not edited by
                    'this program.
                    '---------------------------------------
                    Item.SmallIcon = 5
                Case REG_FULL_RESOURCE_DESCRIPTOR
                    ' --------------------------------------
                    ' Resource Descriptors require a special
                    '   editor to properly be displayed or
                    '   edited.
                    ' --------------------------------------
                    Item.SmallIcon = 5
                    rgeValue = "REG_FULL_RESOURCE_DESCRIPTOR"

                Case REG_DWORD
                    ' --------------------------------------
                    ' REG_DWORD values are 32-bit unsigned
                    '   integers
                    ' Tortuous manipulation to make values
                    '   above 7FFFFFFF appear as positive
                    '   values.
                    ' VB Longs would display them as
                    '   negative numbers.
                    ' --------------------------------------
                    Item.SmallIcon = 5
                    fTempDbl = Asc(Mid$(rgeValue, 1, 1)) + &H100& * Asc(Mid$(rgeValue, 2, 1)) + &H10000 * Asc(Mid$(rgeValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(rgeValue, 4, 1)))
                    If fTempDbl > &H7FFFFFFF Then
                        rgeValue = "&H" & Hex$(fTempDbl - 4294967296#)
                    Else
                        rgeValue = "&H" & Hex$(fTempDbl)
                    End If
                    rgeValue = rgeValue & " (" & Format$(fTempDbl) & ")"

                Case REG_BINARY
              
                    ' --------------------------------------
                    ' Binary values may be of any length,
                    '   and may represent text or other data.
                    ' They require a special editor to
                    '   modify them.
                    ' --------------------------------------
                    Item.SmallIcon = 5
                    For iTempInt = 1 To Len(rgeValue)
                        sBinaryString = sBinaryString & Format$(Hex(Asc(Mid$(rgeValue, iTempInt, 1))), "00") & " "
                    Next iTempInt
                    rgeValue = sBinaryString
            End Select

            If Len(rgeValue) = 0 Then
                rgeValue = "(value not set)"
            End If

         
            ' -----------------------------------------------
            ' Enter the value into the list box
            ' -----------------------------------------------
            Item.Text = rgeEntry
            Item.SubItems(1) = rgeValue
            Item.Tag = CStr(lKeyIndx)
            ' -----------------------------------------------
            ' Increment the key and do it again.
            ' -----------------------------------------------
            lKeyIndx = lKeyIndx + 1

        ElseIf lRtn = ERROR_MORE_DATA Then
            ' -----------------------------------------------
            ' This error means that, despite querying the key,
            '   we have not set one of the buffers large
            '   enough.  Increment the buffer sizes and try
            '   again
            ' -----------------------------------------------
            lMaxValueData = lMaxValueData + 5
            lMaxValueName = lMaxValueName + 5
            GoTo ReTryValueEnumeration

        ElseIf lRtn = ERROR_NO_MORE_ITEMS Then
            ' -----------------------------------------------
            ' Not an error, just end of list -- exit the
            '   loop
            ' -----------------------------------------------
            lRtn = ERROR_SUCCESS
            Exit Do

        Else
            ' --------------------------------------------------
            ' Key still open, so display the error and fall
            '   thru to the close function below
            ' --------------------------------------------------
            MsgBox RtnRegError(lRtn)
            Exit Do
        End If
    Loop
    
    ' -----------------------------------------------------
    ' Always close opened keys!
    ' -----------------------------------------------------
    lRtn = RegCloseKey(hKey)

End Sub

Function RegGetValue$()
   ' This function is not used in this demo, but is a sample for
   ' retrieving a single value
   ' rgeMainKey must be one of the Publicly declared HKEY constants.
   Dim sKeyType&       'returns the key type.  This function expects REG_SZ or REG_DWORD
   Dim ret&            'returned by registry functions, should be 0&
   Dim lpHKey&         'return handle to opened key
   Dim lpcbData&       'length of data in returned string
   Dim ReturnedString$ 'returned string rgeValue
   Dim fTempDbl!
   If rgeMainKey >= &H80000000 And rgeMainKey <= &H80000006 Then
      ' Open key
      ret = RegOpenKeyEx(rgeMainKey, rgeSubKey, 0&, KEY_READ, lpHKey)
      If ret <> ERROR_SUCCESS Then
         RegGetValue = ""
         Exit Function     'No key open, so leave
      End If
      
      ' Set up buffer for data to be returned in.
      ' Adjust next rgeValue for larger buffers.
      lpcbData = 255
      ReturnedString = Space$(lpcbData)

      ' Read key
      ret& = RegQueryValueEx(lpHKey, rgeValue, ByVal 0&, sKeyType, ReturnedString, lpcbData)
      If ret <> ERROR_SUCCESS Then
         RegGetValue = ""   'Key still open, so finish up
      Else
        If sKeyType = REG_DWORD Then
            fTempDbl = Asc(Mid$(ReturnedString, 1, 1)) + &H100& * Asc(Mid$(ReturnedString, 2, 1)) + &H10000 * Asc(Mid$(ReturnedString, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(ReturnedString, 4, 1)))
            ReturnedString = Format$(fTempDbl, "000")
        End If
         RegGetValue = Left$(ReturnedString, lpcbData - 1)
    End If
      ' Always close opened keys.
      ret = RegCloseKey(lpHKey)
   End If
End Function

Sub RegSetValue()
    ' --------------------------------------------------------
    ' DWORD Values must be in Hex form for this function to
    '   work.
    ' --------------------------------------------------------
    Dim lRtn&            'returned by registry functions, should be 0&
    Dim hKey&         'return handle to opened key
    Dim iFirstChar%
    Dim iSecondChar%
    Dim iThirdChar%
    Dim iFourthChar%
    Dim lpDisp&
    
    If rgeDataType <> REG_SZ And rgeDataType <> REG_DWORD Then
        MsgBox "This demo only supports writing keys of the types REG_SZ and REG_DWORD.  This key uses a different type."
        Exit Sub
    End If
    
    ' -----------------------------------------------------
    ' Check rgeMainKey for validity
    ' -----------------------------------------------------
    If rgeMainKey >= &H80000000 And rgeMainKey <= &H80000006 Then

        ' -----------------------------------------------------
        ' Open key
        ' -----------------------------------------------------
        lRtn = RegCreateKeyEx(rgeMainKey, rgeSubKey, 0&, "", 0&, KEY_WRITE, 0&, hKey, lpDisp)
        If lRtn <> ERROR_SUCCESS Then
            MsgBox RtnRegError(lRtn)
            rgeClear
            Exit Sub       'No key open, so leave
        End If
      
        ' -----------------------------------------------------
        ' Write new rgeValue to key
        ' -----------------------------------------------------
        If rgeDataType = REG_DWORD Then
            rgeValue = Left(Trim(rgeValue), 8)
            If Left$(rgeValue, 2) <> "&H" Then
            rgeValue = "&H" & Left(Trim(rgeValue), 8)
            End If
            If Len(rgeValue) <= 6 Then
            rgeValue = rgeValue & "&"
            End If
        
            ' -----------------------------------------------------
                'Convert number string to 32-bit DWORD and save:
            ' -----------------------------------------------------
            lRtn = RegSetValueEx(hKey, rgeEntry, 0&, rgeDataType, CLng(Val(rgeValue)), 4&)
        Else

            ' -----------------------------------------------------
                'Save type REG_SZ (strings)
            ' -----------------------------------------------------
            lRtn = RegSetValueEx(hKey, rgeEntry, 0&, rgeDataType, ByVal rgeValue, CLng(Len(rgeValue)))
        End If
        If lRtn <> ERROR_SUCCESS Then
            MsgBox RtnRegError(lRtn) 'Key still open, so finish up
        End If
        ' -----------------------------------------------------
        ' Always close opened keys!
        ' -----------------------------------------------------
        lRtn = RegCloseKey(hKey)
    End If
End Sub

Private Function RtnRegError$(errorcode&)
    Select Case errorcode
        Case 1009, 1015
            ' -----------------------------------------------------
                'We're in trouble now
            ' -----------------------------------------------------
            RtnRegError = "The Registry Database is corrupt!"
        Case 2, 1010
            RtnRegError = "Bad Key Name!"
        Case 1011
            RtnRegError = "Can't Open Key"
        Case 4, 1012
            RtnRegError = "Can't Read Key"
        Case 5
            RtnRegError = "Access to this key is denied."
        Case 1013
            RtnRegError = "Can't Write Key"
        Case 8, 14
            RtnRegError = "Out of memory"
        Case 87
            RtnRegError = "Invalid Parameter"
        Case 234
            RtnRegError = "Error - There is more data than the buffer can handle!"
        Case Else
            RtnRegError = "Undefined Key Error Code" & Str$(errorcode) & "!"
    End Select
End Function

Function WordLo(lLongIn&) As Integer
    If (lLongIn And &HFFFF&) > &H7FFF Then
        WordLo = (lLongIn And &HFFFF&) - &H10000
    Else
        WordLo = lLongIn And &HFFFF&
    End If
End Function

