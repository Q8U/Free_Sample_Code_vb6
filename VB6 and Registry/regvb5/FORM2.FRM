VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   Caption         =   "Form2"
   ClientHeight    =   2784
   ClientLeft      =   2760
   ClientTop       =   3432
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2784
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Base"
      Height          =   912
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   2232
      Begin VB.OptionButton Option1 
         Caption         =   "Decimal"
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   0
         Top             =   540
         Width           =   1512
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hexadecimal"
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1512
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   312
      Index           =   1
      Left            =   4140
      TabIndex        =   6
      Top             =   1920
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   312
      Index           =   0
      Left            =   2940
      TabIndex        =   5
      Top             =   1920
      Width           =   912
   End
   Begin VB.TextBox Text2 
      Height          =   312
      Left            =   300
      TabIndex        =   4
      Top             =   1260
      Width           =   4812
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   312
      Left            =   300
      TabIndex        =   2
      Top             =   420
      Width           =   4812
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Value Data:"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1632
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Value Name:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   360
      TabIndex        =   1
      Top             =   60
      Width           =   2172
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bPainted As Boolean

Private Sub Command1_Click(iIndex%)
    Dim nRtn&
    Dim varTempVar
    Dim Item As ListItem
    Select Case iIndex
        Case 0
            
            '-------------------------------------------------
            'OK was clicked
            'Values were loaded into visible and
            'hidden text boxes by the EditRegValue function.
            '-------------------------------------------------
            rgeValue = Text2
            rgeEntry = Text1
            If rgeEntry = "(Default)" Then
                
                '---------------------------------------------
                'We are visually displaying (Default) when the
                'value name is null.
                '---------------------------------------------
                rgeEntry = ""
            End If

            '-------------------------------------------------
            'We handle DWORDs differently than strings
            '-------------------------------------------------
            If rgeDataType = REG_DWORD Then
                
                '---------------------------------------------
                'We do some tortuous maneuvers to make sure
                'that all values are hex longs, including decimal
                'values greater than 7FFFFFFF
                '---------------------------------------------
                varTempVar = Text2
                If Option1(0) Then
                    varTempVar = "&H" & varTempVar
                End If
                If IsNumeric(varTempVar) Then
                    If Option1(0) Then
                        rgeValue = Text2
                    Else
                        fTempDbl = Val(Text2 & "#")
                        If fTempDbl > &H7FFFFFFF Then
                            rgeValue = Hex$(fTempDbl - 4294967296#)
                        Else
                            rgeValue = Hex$(fTempDbl)
                        End If
                    End If
                Else
                    '-----------------------------------------
                    'We have a DWORD value, but the text box
                    'does not contain an number.
                    '-----------------------------------------
                    Exit Sub
                End If
            End If
            
            '-------------------------------------------------
            'Write new value to the registry
            '-------------------------------------------------
            RegSetValue
            
            '-------------------------------------------------
            'If the value was successfully changed, update the list box
            '-------------------------------------------------
            If nRtn = 0 Then
                rgeEntry = Text1 'Get "Default" back
                If rgeDataType = REG_SZ Then
                    rgeValue = """" & rgeValue & """"
                End If
                If rgeDataType = REG_DWORD Then
                    
                    '-----------------------------------------
                    'Tortuous calculations again to display positive
                    'integer values greater than a VB long may contain.
                    '-----------------------------------------
                    fTempDbl = Val("&H" & rgeValue)
                    If fTempDbl < 0 Then
                        fTempDbl = fTempDbl + 4294967296#
                    End If
                    If Right$(rgeValue, 1) = "&" Then
                        rgeValue = Left$(rgeValue, Len(rgeValue) - 1)
                    End If
                    rgeValue = UCase(rgeValue) & " (" & Format$(fTempDbl) & ")"
                End If
                
                Set Item = Form1.ListView1.ListItems.Item(Form1.ListView1.SelectedItem.Index)
                Item.SubItems(1) = rgeValue
            End If
            Unload Form2
    
        Case 1
            
            '-------------------------------------------------
            'Cancel was clicked
            '-------------------------------------------------
            Unload Form2
    End Select
End Sub

Private Sub Form_Paint()
    If Not bPainted Then
        bPainted = True
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2) + 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
bPainted = False
End Sub

Private Sub Option1_Click(iIndex%)
Select Case iIndex
    Case 0
        
        '------------------------------------------
        'Convert from decimal to Hex
        '------------------------------------------
        sTempString = Text2
        If Val(sTempString) < 0 Then
            
            '--------------------------------------
            'Force values that otherwise would be
            'given negative integer numbers to type
            'long.
            '--------------------------------------
            sTempString = sTempString & "&"
        End If
        fTempDbl = Val(Text2 & "#")
            
            '--------------------------------------
            'Hex function will fail on values outside
            'the normal VB long range.  Convert such
            'to a negative value within the range
            '--------------------------------------
            If fTempDbl > &H7FFFFFFF Then
                Text2 = Hex$(fTempDbl - 4294967296#)
            Else
                Text2 = Hex$(fTempDbl)
            End If
    Case 1
        
        '------------------------------------------
        'Convert from Hex to decimal
        '------------------------------------------
        sTempString = Text2
        If Len(sTempString) < 5 Then
            
            '--------------------------------------
            'Force values that otherwise would be
            'given negative integer numbers to type
            'long.
            '--------------------------------------
            sTempString = sTempString & "&"
        End If
        fTempDbl = Val("&H" & sTempString)
        If fTempDbl < 0 Then
            
            '--------------------------------------
            'Val function will return a negative long
            'for hex numbers greater than 7FFFFFFF.
            'Convert these to positive values for
            'display purposes.
            '--------------------------------------
            fTempDbl = fTempDbl + 4294967296#
        End If
        Text2 = Format$(fTempDbl)
End Select
End Sub

