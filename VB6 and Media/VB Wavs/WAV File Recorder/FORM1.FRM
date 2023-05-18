VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   2295
   ClientTop       =   2025
   ClientWidth     =   5430
   Height          =   2025
   Left            =   2235
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   5430
   Top             =   1680
   Width           =   5550
   Begin CommandButton btnPlay 
      Caption         =   "Play"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin CommandButton btnRecord 
      Caption         =   "Record"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Sub btnPlay_Click ()
  PlayRecSound
End Sub

Sub btnRecord_Click ()
  RecordSound
End Sub

Sub Form_Unload (Cancel As Integer)
  CloseSound
End Sub

