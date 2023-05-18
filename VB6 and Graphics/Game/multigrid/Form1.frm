VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Bookstore"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Copy selected titles to the list window"
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   6000
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\DevStudio\VB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   5520
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   6375
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0010
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    List1.Clear         ' Clears the listbox
End Sub

Private Sub Command2_Click()
    Unload Form1
    End                 ' Ends application
End Sub

Private Sub Command3_Click()
Dim SelRecord As Integer

For SelRecord = 0 To DBGrid1.SelBookmarks.Count - 1
    Data1.Recordset.Bookmark = DBGrid1.SelBookmarks(SelRecord)
    List1.AddItem Data1.Recordset(0)
Next SelRecord

End Sub

Private Sub Data1_Reposition()
    ' Assign the Title to the caption property
    Data1.Caption = Data1.Recordset(0)
End Sub

Private Sub DBGrid1_DblClick()
    ' Add the title to the list
    List1.AddItem Data1.Recordset(0)
End Sub



Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call GridMultiSelect(Button, Shift, X, Y, DBGrid1, Data1.Recordset)
End Sub

Private Sub Form_Load()
Form1.Width = 9500
End Sub

Private Sub Form_Resize()

Const Spacer As Integer = 125
DBGrid1.Width = (Form1.Width - (Spacer * 2)) - 100
DBGrid1.Left = Spacer
List1.Width = (Form1.Width - (Spacer * 2)) - 100
List1.Left = Spacer

End Sub
