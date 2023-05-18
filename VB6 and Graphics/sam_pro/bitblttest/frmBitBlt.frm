VERSION 5.00
Begin VB.Form frmBitBlt 
   Caption         =   "BitBlt Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   197
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBitBlt 
      Caption         =   "Draw It"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox picBitBlt 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   975
      Left            =   1680
      Picture         =   "frmBitBlt.frx":0000
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "frmBitBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'BitBlt Sample Test Project
'

Option Explicit


'Local BitBlt declaration
Private Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long

Private Sub cmdBitBlt_Click()

Me.Cls
BitBlt Me.hDC, 0, 0, picBitBlt.ScaleWidth, picBitBlt.ScaleHeight, picBitBlt.hDC, 0, 0, vbSrcCopy

End Sub


