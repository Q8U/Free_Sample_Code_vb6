VERSION 5.00
Begin VB.Form frmMemoryDC 
   AutoRedraw      =   -1  'True
   Caption         =   "Using Memory DCs"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw the sprite"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Bitmap"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmMemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Memory Device Context
'

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
'****************************************

Dim DCMask As Long
Dim DCSprite As Long

Const SpriteWidth As Long = 64
Const SpriteHeight As Long = 64




Private Sub cmdDraw_Click()

'draw the transparent sprite
BitBlt Me.hdc, 0, 0, SpriteWidth, SpriteHeight, DCMask, 0, 0, vbSrcAnd
BitBlt Me.hdc, 0, 0, SpriteWidth, SpriteHeight, DCSprite, 0, 0, vbSrcPaint

Me.Refresh

End Sub

Private Sub cmdExit_Click()

DeleteGeneratedDC DCMask
DeleteGeneratedDC DCSprite

Unload Me
Set frmMemoryDC = Nothing

End Sub

Private Sub cmdLoad_Click()
'Generate the two DCs, one for the mask and one for the sprite
DCMask = GenerateDC(App.Path & "\mask.bmp")

If DCMask <= 0 Then
    MsgBox "Failure in creating mask dc"
    Exit Sub
End If

DCSprite = GenerateDC(App.Path & "\sprite.bmp")

If DCSprite <= 0 Then
    MsgBox "Failure in creating sprite dc"
    'Clean up after the created mask
    DeleteGeneratedDC DCMask
End If


End Sub

'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function
