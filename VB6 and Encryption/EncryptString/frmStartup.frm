VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Just Another Joe Production"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRatio 
      Height          =   765
      Left            =   6975
      TabIndex        =   15
      ToolTipText     =   "Ratio of Length of Encrypted to Length of Original"
      Top             =   3750
      Width           =   1965
   End
   Begin VB.TextBox txtTimeDecrypt 
      Height          =   690
      Left            =   3975
      TabIndex        =   12
      ToolTipText     =   "Time to Decrypt (in sec)"
      Top             =   4650
      Width           =   1215
   End
   Begin VB.TextBox txtTimeEncrypt 
      Height          =   690
      Left            =   1275
      TabIndex        =   11
      Top             =   4650
      Width           =   1215
   End
   Begin VB.TextBox txtLenEncrypted 
      Height          =   690
      Left            =   3975
      TabIndex        =   8
      ToolTipText     =   "Length of Encrypted String"
      Top             =   3825
      Width           =   1215
   End
   Begin VB.TextBox txtLenOrginal 
      Height          =   690
      Left            =   1275
      TabIndex        =   7
      Top             =   3825
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "&Click here to Encrypt/Decrypt"
      Height          =   3465
      Left            =   7350
      TabIndex        =   6
      Top             =   225
      Width           =   1590
   End
   Begin VB.TextBox txtDecrypted 
      Height          =   915
      Left            =   1425
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Decrypted String"
      Top             =   2775
      Width           =   5790
   End
   Begin VB.TextBox txtEncrypted 
      Height          =   1365
      Left            =   1425
      MultiLine       =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Encrypted String"
      Top             =   1275
      Width           =   5790
   End
   Begin VB.TextBox txtOriginal 
      Height          =   915
      Left            =   1425
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "String To Encrypt"
      Top             =   225
      Width           =   5790
   End
   Begin VB.Label lblRatio 
      Caption         =   "Ratio of Length of Encrypted to Length of Original"
      Height          =   615
      Left            =   5400
      TabIndex        =   16
      ToolTipText     =   "Ratio of Length of Encrypted to Length of Original"
      Top             =   3825
      Width           =   1515
   End
   Begin VB.Label lblTimeDecrypt 
      Caption         =   "Time to Decrypt (in sec)"
      Height          =   540
      Left            =   2700
      TabIndex        =   14
      ToolTipText     =   "Time to Decrypt (in sec)"
      Top             =   4725
      Width           =   1215
   End
   Begin VB.Label lblTimeEncrypt 
      Caption         =   "Time to Encrypt (in sec)"
      Height          =   540
      Left            =   0
      TabIndex        =   13
      Top             =   4725
      Width           =   1215
   End
   Begin VB.Label lblLenEncrypted 
      Caption         =   "Length of Encrypted String"
      Height          =   540
      Left            =   2700
      TabIndex        =   10
      ToolTipText     =   "Length of Encrypted String"
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label lblLenOrginal 
      Caption         =   "Length of Orgininal String"
      Height          =   540
      Left            =   0
      TabIndex        =   9
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label lblDecrypted 
      Caption         =   "Decrypted String"
      Height          =   765
      Left            =   75
      TabIndex        =   5
      ToolTipText     =   "Decrypted String"
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Label lblEncrypted 
      Caption         =   "Encrypted String"
      Height          =   540
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Encrypted String"
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label lblOriginal 
      Caption         =   "String To Encrypt"
      Height          =   690
      Left            =   75
      TabIndex        =   1
      ToolTipText     =   "String To Encrypt"
      Top             =   225
      Width           =   1215
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   いいいいいいい                                                 いいいいいいい
'   いいいいいいい  All Functions and Subroutines are the Complete いいいいいいい
'   いいいいいいい    and Expressed Property of Joseph Sullivan.   いいいいいいい
'   いいいいいいい                                                 いいいいいいい
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   いいいいいいい                                                 いいいいいいい
'   いいいいいいい  If you have any questions or comments, please  いいいいいいい
'   いいいいいいい     contact Mr. Sullivan at bhJoeS@aol.com.     いいいいいいい
'   いいいいいいい                                                 いいいいいいい
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   いいいいいいい                                                 いいいいいいい
'   いいいいいいい          Visual Basic 5.0 Specific Code         いいいいいいい
'   いいいいいいい                                                 いいいいいいい
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�

'   Module Name
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   frmStartup

'   Last Updated
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   Tuesday, August 01, 2000

'   Dependants
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   modSecurity

'   Private Dimensions
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   [None]

'   Private Constants
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   [None]

'   Public Subroutines
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   [None]

'   Private Subroutines
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   btnStart_Click
'   Form_Load

'   Public Functions
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   [None]

'   Private Functions
'   いいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいいい�
'   [None]

Option Explicit

Private Sub btnStart_Click()

Remarks:
    '   This subroutine fills in the various textboxes on the form 'frmStartup'
    '   using the functions 'Encrypt' and 'Decrypt' found in the module
    '   'modSecurity'.  The time it takes to complete these tasks, the length of
    '   the results, and a character per second count is also found and place
    '   into the appropriate textboxes.  Upon any error, the error is ignored and
    '   execution of the subroutine continues.

OnError:
    On Error GoTo ErrHandler

Dimensions:
    Dim intMousePointer As Integer
    Dim strStartTime As Date

Constants:
    '   [None]

MainCode:
    '   Store the current value of the mouse pointer
    Let intMousePointer = Screen.MousePointer
    '   Change the mousepointer to an hourglass.
    Let Screen.MousePointer = vbHourglass
    '   Store the length of the textbox 'txtOriginal' into the textbox
    '   'txtLenOriginal'
    Let Me.txtLenOrginal = Len(Me.txtOriginal)
    '   Store the current time and date into the variable 'strStartTime'
    Let strStartTime = Now
    '   Using the function 'Encrypt' found in the 'modSecurity' module, encrypt
    '   the value found in the textbox 'txtOriginal' and place it into the
    '   textbox 'txtEncrypted'
    Let Me.txtEncrypted = Encrypt(Me.txtOriginal)
    '   Store the difference (in seconds) between the current time and the value
    '   found in the variable 'strStartTime' into the textbox 'txtTimeEncrypt'
    Let Me.txtTimeEncrypt = Abs(DateDiff("s", Now, strStartTime))
    '   Store the length of the textbox 'txtEncrypted' into the textbox
    '   'txtLenEncrypted'
    Let Me.txtLenEncrypted = Len(Me.txtEncrypted)
    '   Store the current time and date into the variable 'strStartTime'
    Let strStartTime = Now
    '   Using the function 'Decrypt' found in the 'modSecurity' module, decrypt
    '   the value found in the textbox 'txtEncrypted' and place it into the
    '   textbox 'txtDecrypted'
    Let Me.txtDecrypted = Decrypt(Me.txtEncrypted)
    '   Store the difference (in seconds) between the current time and the value
    '   found in the variable 'strStartTime' into the textbox 'txtTimeDecrypt'
    Let Me.txtTimeDecrypt = Abs(DateDiff("s", Now, strStartTime))
    '   Divide the textbox 'txtLenOriginal' into the textbox 'txtLenEncrypted',
    '   ensuring that two decimal places are fixed, and place it into the textbox
    '   'txtRatio'
    Let Me.txtRatio = Int(Me.txtLenEncrypted * 100 / Me.txtLenOrginal) / 100
    '   Return the mousepointer to the value that it was before the function
    '   started
    Let Screen.MousePointer = intMousePointer
    '   Stop execution of the subroutine
    Exit Sub

ErrHandler:
    '   Begin selecting occurences of an error number when an error has occured
    Select Case Err.Number
        '   For all occurences of an error number, do what follows
        Case Else
            '   Erase the error
            Err.Clear
            '   Go to the line of code that follows the error
            Resume Next
        '   Stop selecting occurences of an error number
    End Select

End Sub

Private Sub Form_Load()

Remarks:
    '   This subroutine fills in the textbox 'txtOriginal' with the string
    '   'Just Another Joe Production'.  Upon any error, the error is ignored and
    '   execution of the subroutine continues.

OnError:
    On Error GoTo ErrHandler

Dimensions:
    Dim intMousePointer As Integer

Constants:
    '   [None]

MainCode:
    '   Store the current value of the mouse pointer
    Let intMousePointer = Screen.MousePointer
    '   Change the mousepointer to an hourglass.
    Let Screen.MousePointer = vbHourglass
    '   As a default, place the string 'Just Another Joe Production' into the
    '   textbox 'txtOriginal'
    Let Me.txtOriginal = "Just Another Joe Production"
    '   Return the mousepointer to the value that it was before the function
    '   started
    Let Screen.MousePointer = intMousePointer
    '   Stop execution of the subroutine
    Exit Sub

ErrHandler:
    '   Begin selecting occurences of an error number when an error has occured
    Select Case Err.Number
        '   For all occurences of an error number, do what follows
        Case Else
            '   Erase the error
            Err.Clear
            '   Go to the line of code that follows the error
            Resume Next
        '   Stop selecting occurences of an error number
    End Select

End Sub
