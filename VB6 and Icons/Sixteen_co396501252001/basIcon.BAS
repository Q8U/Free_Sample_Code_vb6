Attribute VB_Name = "Basic_Code"
Global BitPos(7) As Integer
Global BoxEndX As Integer
Global BoxEndY As Integer
Global BoxFlag As Boolean
Global BoxHeight As Integer
Global BoxStartX As Integer
Global BoxStartY As Integer
Global BoxWidth As Integer
Global CircEndX As Integer
Global CircEndY As Integer
Global CheckText As String
Global CircFlag As Boolean
Global CircStartFlag As Boolean
Global CircStartX As Integer
Global CircStartY As Integer
Global ClickedPixelX As Integer
Global ClickedPixelY As Integer
Global ExitAfterSaveFlag As Boolean
Global FileChangedFlag As Boolean
Global FileFormat As Byte
Global FileData As String * 745
Global FileHeader As String * 103
Global FileFoundFlag As Boolean
Global FillFlag As Boolean
Global FormOpenLoadedFlag As Boolean
Global Hdr As String
Global LineEndX As Integer
Global LineEndY As Integer
Global LineStartFlag As Boolean
Global LineStartX As Integer
Global LineStartY As Integer
Global LineFlag As Boolean
Global LoadAfterSaveFlag As Boolean
Global LoadName As String
Global LoadPath As String
Global NewAfterSaveFlag As Boolean
Global NumberOfIcons As Integer
Global OpenPal(15) As Long
Global Pal(15) As Long
Global PathText As String
Global PenFlag As Boolean
Global PixArray(31, 31) As Long
Global PrevButton As Byte
Global PrevCircEndX As Integer
Global PrevCircEndY As Integer
Global PlotAdjust As Integer
Global RadiusSquared As Single
Global SaveName As String
Global SavePath As String
Global SelectedColor As Integer
Global SizeOfFile As Long
Global SkipFlag As Boolean
Global SqWidth As Integer
Global SqHeight As Integer
Global StartOfData As Integer
Global TempColor As Long
Global TempDrive As String
Global TempDirPath As String
Global TempFileName As String
Global TempPattern As String
Global TextStatus As Byte
Global TransparentColor As Integer
Global UnusedColors(15) As Integer
Global XFactor As Single
Global XFocus As Single
Global XPlot As Integer
Global X As Integer
Global Y As Integer
Global YFactor As Single
Global YFocus As Single
Global YPlot As Integer

Sub CheckForBlack()
TransparentColor = -1
For n% = 0 To 15
If Pal(n%) = 0 Then
  TransparentColor = n%
  TransparentColor = n%
  If (n% And 1) = 0 Then
    frmMain.imgRtArrow.Top = frmMain.lblPalNum(n%).Top + 4
    frmMain.imgRtArrow.Left = frmMain.lblPalNum(n%).Left - 30
    frmMain.imgRtArrow.Visible = True
    frmMain.imgLtArrow.Visible = False
    Else
    frmMain.imgLtArrow.Top = frmMain.lblPalNum(n%).Top + 4
    frmMain.imgLtArrow.Left = frmMain.lblPalNum(n%).Left + 30
    frmMain.imgLtArrow.Visible = True
    frmMain.imgRtArrow.Visible = False
  End If
  Exit Sub
End If
Next n%
TransparentColor = -1
frmMain.imgRtArrow.Visible = False
frmMain.imgLtArrow.Visible = False
End Sub

Sub DrawGrid()
Co& = &HC0C0C0
For X = 0 To frmMain.picMain.Width - 1 Step SqWidth + 1
frmMain.picMain.Line (X, 0)-(X, frmMain.picMain.Height - 1), Co&
Next X
For Y = 0 To frmMain.picMain.Height - 1 Step SqHeight + 1
frmMain.picMain.Line (0, Y)-(frmMain.picMain.Width - 1, Y), Co&
Next Y
End Sub

Sub DrawIcon()
frmMain.MousePointer = 11
For Y = 0 To 31
For X = 0 To 31
frmMain.picMain.Line (X * (SqWidth + 1) + 1, Y * (SqHeight + 1) + 1)-((X + 1) * (SqWidth + 1) - 1, (Y + 1) * (SqHeight + 1) - 1), Pal(PixArray(X, Y)), BF
frmMain.picIcon1.PSet (X, Y), Pal(PixArray(X, Y))
If PixArray(X, Y) = TransparentColor Then
  frmMain.picMain.Line (Int((X + 0.5) * (SqWidth + 1)), Int((Y + 0.5) * (SqHeight + 1)))-(Int((X + 0.5) * (SqWidth + 1)) + 1, Int((Y + 0.5) * (SqHeight + 1)) + 1), 16777215, BF
  frmMain.picIcon1.PSet (X, Y), 12632256 'RGB(192, 192, 192)
End If
Next X
Next Y
frmMain.MousePointer = 1
End Sub

Sub LoadIconFile()
On Error GoTo Oops
If LoadAfterSaveFlag = True Then GoTo SkipCheck
If FileChangedFlag = True Then
  A% = MsgBox("Save Current Icon File?", 51, "File NOT Saved")
  If A% = 6 Then
    'Yes
    LoadAfterSaveFlag = True
    Load frmSaveAs
    Exit Sub
  End If
  'If a% = 7 Then Stop 'No
  If A% = 2 Then
    LoadAfterSaveFlag = False
    Unload frmSaveAs
    frmMain.Enabled = True
    frmMain.SetFocus
    Exit Sub 'Cancel
  End If
End If
SkipCheck:
frmMain.MousePointer = 11
For A% = 0 To 15
Pal(A%) = OpenPal(A%)
UnusedColors(A%) = 1
Next A%

For Y = 0 To 31
For X = 0 To 30 Step 2
Byt% = X \ 2 + Y * 16 + 106
PixArray(X, 31 - Y) = (Asc(Mid$(FileData, Byt%, 1)) And 240) / 16
PixArray(X + 1, 31 - Y) = Asc(Mid$(FileData, Byt%, 1)) And 15
Next X
Next Y

'Check for unused color
For Y = 0 To 31
For X = 0 To 31
UnusedColors(PixArray(X, Y)) = 0
Next X, Y
TransparentColor = -1
For A% = 0 To 15
If UnusedColors(A%) = 1 Then
  TransparentColor = A%
  Exit For
End If
Next A%
'Find pure black palette that is used in icon
TransparentColor = -1
For A% = 0 To 15
If (Pal(A%) = 0) And (UnusedColors(A%) = 0) Then
  TransparentColor = A%
  Exit For
End If
Next A%
If TransparentColor < 0 Then GoTo Skip1
If TransparentColor < 0 Then GoTo Skip1

'Check to see if black has mixed transparent/opaque pixels
Transparent% = 0
Opaque% = 0
Byt% = 618
For Y = 31 To 0 Step -1
For X = 0 To 24 Step 8
For X2% = 0 To 7
If PixArray(X + X2%, Y) = TransparentColor Then
  If (Asc(Mid$(FileData, Byt%, 1)) And BitPos(X2%)) > 0 Then
    Transparent% = 1
    Else
    Opaque% = 1
  End If
  If Opaque% = 1 And Transparent% = 1 Then GoTo MixedTrue
End If
Next X2%
Byt% = Byt% + 1
Next X
Next Y
GoTo Skip1

MixedTrue:
'Ok to Change transparency format
Byt% = 618
For Y = 31 To 0 Step -1
For X = 0 To 24 Step 8
For X2% = 0 To 7
If PixArray(X + X2%, Y) = TransparentColor Then
  If (Asc(Mid$(FileData, Byt%, 1)) And BitPos(X2%)) = 0 Then
    PixArray(X + X2%, Y) = Transparent%
    'frmMain.picMain.Line ((X + X2%) * (SqWidth + 1) + 1, Y * (SqHeight + 1) + 1)-((X + X2% + 1) * (SqWidth + 1) - 1, (Y + 1) * (SqHeight + 1) - 1), 16777215 - Pal(PixArray(X + X2%, Y))
  End If
End If
Next X2%
Byt% = Byt% + 1
Next X
Next Y
Pal(TransparentColor) = 0
Pal(Transparent%) = 65793

Skip1:
frmMain.MousePointer = 1
Call DrawIcon
Call ShowPalette
FileChangedFlag = False: frmMain.cmdRefresh.Enabled = False
LoadAfterSaveFlag = False
SaveName = LoadName
SavePath = LoadPath
If Len(CurDir$) > 3 Then
  frmMain.Text1.Text = CurDir$ + "\" + LoadName
  Else: frmMain.Text1.Text = CurDir$ + LoadName
End If
frmMain.Enabled = True
frmMain.SetFocus
Exit Sub

Oops:
frmMain.MousePointer = 1
End Sub
Sub NewFile()
NewAfterSaveFlag = False
FileFoundFlag = False
FileChangedFlag = False
frmMain.cmdRefresh.Enabled = False
SaveName = "UNTITLED.ico"
frmMain.Text1.Text = SaveName
Pal(0) = 0
Pal(1) = RGB(1, 1, 1) '65793
Pal(2) = RGB(0, 0, 128) '8388608
Pal(3) = RGB(0, 128, 0) '32768&
Pal(4) = RGB(0, 128, 128) '8421376
Pal(5) = 127
Pal(6) = RGB(128, 0, 128) '8388736
'Pal(6) = RGB(128, 128, 0) '32639&
Pal(7) = RGB(128, 128, 128) '8421504
Pal(8) = RGB(192, 192, 192) '12632256
Pal(9) = RGB(0, 0, 255)     '16711680
Pal(10) = RGB(0, 255, 0) '65280
Pal(11) = RGB(0, 255, 255) '16776960
Pal(12) = 255
Pal(13) = RGB(255, 0, 255) '16711935
Pal(14) = RGB(255, 255, 0)   '65535
Pal(15) = RGB(255, 255, 255) '16777215
frmMain.lblBlueSB.Caption = frmMain.vsbBlue.Value
frmMain.lblGreenSB.Caption = frmMain.vsbGreen.Value
frmMain.lblRedSB.Caption = frmMain.vsbRed.Value
SelectedColor = 0
Call ShowPalette
Erase PixArray
Call DrawIcon
Call DrawGrid
End Sub

Sub SaveIconFile()
On Error GoTo Oops
If InStr(SaveName, ".") = 0 Then SaveName = SaveName + ".ico"
frmMain.MousePointer = 11

Open SaveName For Binary As #1
Get #1, 1, FileHeader
NumberOfIcons = Asc(Mid$(FileHeader, 5, 1))
If NumberOfIcons > 1 Then
  Close
  A% = MsgBox("Destination file contains multiple icons." & Chr$(13) & "Please choose a different name.", 48, "Save Error")
  Load frmSaveAs
  Exit Sub
End If
Unload frmSaveAs
For X = 1 To 62
d$ = Chr$(Val("&H" + Mid$(Hdr, X * 2 - 1, 2)))
Put #1, X, d$
Next X
For A% = 0 To 15
B% = Int(Pal(A%) / 65536) And 255
g% = Int(Pal(A%) / 256) And 255
r% = Pal(A%) And 255
p& = r% * 65536 + g% * 256& + B%
Put #1, A% * 4 + 63, p&
Next A%
For Y = 0 To 31
For X = 0 To 30 Step 2
d$ = Chr$(PixArray(X, 31 - Y) * 16 Or PixArray(X + 1, 31 - Y))
Put #1, Y * 16 + 127 + X / 2, d$
Next X
Next Y
'Transparent black
For Y = 0 To 31
For X = 0 To 24 Step 8
Sum% = 0
Byt% = 639 + Int((X + (31 - Y) * 32) / 8)
For X2% = 0 To 7
Bit% = 0
If PixArray(X + X2%, Y) = TransparentColor Then Bit% = 1
Sum% = Sum% Or (BitPos(X2%) * Bit%)
Next X2%
d$ = Chr$(Sum%)
Put #1, Byt%, d$
Next X
Next Y
Close
frmMain.MousePointer = 1
If ExitAfterSaveFlag = True Then
  Unload frmSaveAs
  Unload frmMain
  Exit Sub
End If
Call DrawIcon
FileChangedFlag = False
frmMain.cmdRefresh.Enabled = False
'SavePath = CurDir$
If Len(CurDir$) > 3 Then
  frmMain.Text1.Text = SavePath + "\" + SaveName
  Else: frmMain.Text1.Text = SavePath + SaveName
End If
frmMain.Enabled = True
frmMain.SetFocus
If LoadAfterSaveFlag = True Then
  Call LoadIconFile
  Exit Sub
End If
If NewAfterSaveFlag = True Then Call NewFile
LoadPath = SavePath
Exit Sub

Oops:
e% = Err
Close
A% = MsgBox("Cannot Save" + Chr$(10) & Err.Description, 48, "Save Error")
frmMain.Enabled = True
frmMain.SetFocus
LoadAfterSaveFlag = False
frmMain.MousePointer = 1
ExitAfterSaveFlag = False
End Sub
Sub ShowPalette()
For A% = 0 To 15
frmMain.lblPalNum(A%).BackColor = Pal(A%)
Next A%
Call SetPalColors
Call OutlineBox
End Sub
Sub ButtonSelect(A As Integer)
For n% = 0 To 4
If A = n% Then
  frmMain.imgButtonOn(n%).Visible = True
  frmMain.imgButtonOff(n%).Visible = False
  frmMain.lblButtonLabel(n%).Enabled = True
  Else
  frmMain.imgButtonOff(n%).Visible = True
  frmMain.imgButtonOn(n%).Visible = False
  frmMain.lblButtonLabel(n%).Enabled = False
End If
Next n%
PenFlag = (A = 0)
FillFlag = (A = 1)
LineFlag = (A = 2)
CircFlag = (A = 3)
BoxFlag = (A = 4)
frmMain.chkSolidCirc.Enabled = (A = 3)
frmMain.chkSolidBox.Enabled = (A = 4)
End Sub

Sub CheckFileFormat(ByVal CheckText)
FileFormat = 0
Open CheckText For Binary As #1
Get #1, 1, FileHeader
Close
NumberOfIcons = Asc(Mid$(FileHeader, 5, 1))
StartOfData = 0

If (Left$(FileHeader, 4) = "BMv" + Chr$(2)) And (FileLen(CheckText) = 630) And LCase$(Right$(CheckText, 4)) = ".bmp" Then
  StartOfData = 14
  GoTo FoundOkIcon
End If

For A% = 0 To NumberOfIcons - 1
If Mid$(FileHeader, 7 + A% * 16, 3) = (Chr$(32) + Chr$(32) + Chr$(16)) Then
  StartOfData = Asc(Mid$(FileHeader, 19 + A% * 16, 1)) + Asc(Mid$(FileHeader, 20 + A% * 16, 1)) * 256
  GoTo FoundOkIcon
End If
Next A%

FileFormat = 0
If FormOpenLoadedFlag = False Then Exit Sub
frmOpenFile.cmdOK.Enabled = False
frmOpenFile.lblFormatStatus.ForeColor = 255
frmOpenFile.lblFormatStatus.Caption = "File Format NOT OK"
frmOpenFile.picIcon2.Visible = False
Exit Sub

FoundOkIcon:
FileFormat = 1
Call PreviewOpen(CheckText)
If FormOpenLoadedFlag = False Then Exit Sub
frmOpenFile.cmdOK.Enabled = True
frmOpenFile.lblFormatStatus.ForeColor = 38400
frmOpenFile.lblFormatStatus.Caption = "File Format OK"
frmOpenFile.lblFormatStatus.Refresh
End Sub

Sub GetFileTextInfo()
On Error GoTo Oops

TempFileName = ""
TextStatus = 0

'Root + path and/or file name?
If (Left$(PathText, 1) = "\") Then
  TempDirPath = ""
  PathText = Mid$(PathText, 2)
End If

'Create temp dir path
If InStr(PathText, "\") > 0 Then
  Do
  n% = InStr(PathText, "\")
  TempDirPath = TempDirPath + Left$(PathText, n%)
  PathText = Mid$(PathText, n% + 1)
  Loop While InStr(PathText, "\") > 0
End If

'Is remaining text a directory?
e% = 0
ChDir TempDrive + TempDirPath + PathText
If e% = 0 Then
  'Yes, it's a directory
  TempDirPath = TempDirPath + PathText
  TextStatus = 1 'Valid Dir only, no filename.
  Exit Sub
End If
'No, its't not a directory,
' get file name.
e% = 0
Tmp$ = Dir$(TempDrive + TempDirPath + PathText, 6)
If e% = 0 Then
  'Valid & assume filename doesn't exist
  TextStatus = 2
  TempFileName = PathText
End If
If Tmp$ > "" Then
  'Valid & filename exists
  TextStatus = 3
  TempFileName = PathText
End If
'If TextStatus = 0, invalid entry.
Exit Sub
Oops:
e% = Err
Resume Next
End Sub
Sub OutlineBox()
'Outline palette box
frmMain.shpOutlineBox.Left = frmMain.lblPalNum(SelectedColor).Left - 4
frmMain.shpOutlineBox.Top = frmMain.lblPalNum(SelectedColor).Top - 4
B% = Int(Pal(SelectedColor) / 65536) And 255
g% = Int(Pal(SelectedColor) / 256) And 255
r% = Pal(SelectedColor) And 255
SkipFlag = True 'Prevents FileChangeFlag and CmdRefresh from being enabled
frmMain.vsbBlue.Value = B%: frmMain.lblBlueSB.Caption = B%
frmMain.vsbGreen.Value = g%: frmMain.lblGreenSB.Caption = g%
frmMain.vsbRed.Value = r%: frmMain.lblRedSB.Caption = r%
SkipFlag = False 'Allow FileChangeFlag and CmdRefresh to being enabled
End Sub
Sub SetPalColors()
For n% = 0 To 15
frmMain.lblPalNum(n%).BackColor = Pal(n%)
r% = &HFF& And Pal(n%)
g% = (&HFF00& And Pal(n%)) \ 256
B% = (&HFF0000 And Pal(n%)) \ 65536
Tmp% = Int(r% * 0.3 + g% * 0.59 + B% * 0.11 + 0.01)
If Tmp% > 127 Then
  frmMain.lblPalNum(n%).ForeColor = 0
  Else
  frmMain.lblPalNum(n%).ForeColor = 16777215
End If
Next n%
Call CheckForBlack
End Sub
Sub PreviewOpen(ByVal Name As String)
On Error GoTo ErrHandler
Open Name For Binary As #1
Get #1, StartOfData, FileData
Close
'Next line can be substituted for line below it if you want to check extension also
'If (Left$(FileData, 4) = "BMv" + Chr$(2)) And (SizeOfFile& = 630) And LCase$(Right$(frmOpenFile.Text1.Text, 4)) = ".bmp" Then
For A% = 0 To 15
Byt% = A% * 4 + 42
OpenPal(A%) = RGB(Asc(Mid$(FileData, Byt% + 2, 1)), Asc(Mid$(FileData, Byt% + 1, 1)), Asc(Mid$(FileData, Byt%, 1)))
Next A%
If FormOpenLoadedFlag = False Then Exit Sub
For Y = 0 To 31
For X = 0 To 30 Step 2
Byt% = X \ 2 + Y * 16 + 106
frmOpenFile.picIcon2.PSet (X, 31 - Y), OpenPal((Asc(Mid$(FileData, Byt%, 1)) And 240) / 16)
frmOpenFile.picIcon2.PSet (X + 1, 31 - Y), OpenPal(Asc(Mid$(FileData, Byt%, 1)) And 15)
Next X
Next Y
frmOpenFile.picIcon2.Visible = True
ErrHandler:
e% = Err
End Sub
Sub DrawCircle(ByVal CircX1 As Single, ByVal CircY1 As Single, ByVal CircX2 As Single, ByVal CircY2 As Single)

If (CircX1 = CircX2) Or (CircY1 = CircY2) Then Exit Sub
If CircX1 > CircX2 Then
  swp% = CircX1
  CircX1 = CircX2
  CircX2 = swp%
End If
If CircY1 > CircY2 Then
  swp% = CircY1
  CircY1 = CircY2
  CircY2 = swp%
End If

XFocus = (CircX1 + CircX2) / 2
YFocus = (CircY1 + CircY2) / 2
RadiusSquared = (CircX2 - XFocus) * (CircX2 - XFocus)
XFactor = 1
YFactor = (CircY2 - YFocus) / (CircX2 - XFocus)
Tmp! = (CircY2 - YFocus) * (CircY2 - YFocus)
If Tmp! > RadiusSquared Then
  RadiusSquared = Tmp!
  XFactor = (CircX2 - XFocus) / (CircY2 - YFocus)
  YFactor = 1
End If

prevyplot% = Int(Sqr(RadiusSquared) * YFactor + YFocus + 0.5)
For XPlot = Int(XFocus + 0.6) To CircX2
  YPlot = Int(Sqr(RadiusSquared - ((XPlot - XFocus) / XFactor) ^ 2) * YFactor + YFocus + 0.5)
  Call StorePixels
  If XPlot < CircX2 Then
    Tempy% = Int(Sqr(RadiusSquared - ((XPlot - XFocus + 0.999) / XFactor) ^ 2) * YFactor + YFocus + 0.5)
    If (YPlot - Tempy%) > 1.5 Then GoTo LoopY
  End If
  prevyplot% = YPlot
SkipPlot:
Next XPlot

LoopY:
For YPlot = prevyplot% To Int(YFocus + 0.6) Step -1
  XPlot = Int(Sqr(RadiusSquared - ((YPlot - YFocus) / YFactor) ^ 2) * XFactor + XFocus + 0.5)
  Call StorePixels
Next YPlot
End Sub
Sub Fill()
  frmMain.MousePointer = 11
  ClickedColor% = PixArray(ClickedPixelX, ClickedPixelY)
  PixArray(ClickedPixelX, ClickedPixelY) = 16
SearchPixels:
  PixelsChanged% = 0
  For X = 0 To 31
  For Y = 0 To 31
  If PixArray(X, Y) = 16 Then
    If (X > 0) Then
      If PixArray(X - 1, Y) = ClickedColor% Then PixArray(X - 1, Y) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
    If (Y > 0) Then
      If PixArray(X, Y - 1) = ClickedColor% Then PixArray(X, Y - 1) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
    If (X < 31) Then
      If PixArray(X + 1, Y) = ClickedColor% Then PixArray(X + 1, Y) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
    If (Y < 31) Then
      If PixArray(X, Y + 1) = ClickedColor% Then PixArray(X, Y + 1) = 16: PixelsChanged% = PixelsChanged% + 1
    End If
  End If
  Next Y
  Next X
  If PixelsChanged% > 0 Then GoTo SearchPixels
  For X = 0 To 31
  For Y = 0 To 31
  If PixArray(X, Y) = 16 Then PixArray(X, Y) = SelectedColor
  Next Y
  Next X
  frmMain.MousePointer = 1
End Sub
Sub StorePixels()
If frmMain.chkSolidCirc.Value = 0 Then
  'frmMain.picMain.Line (XPlot * (SqWidth + 1) + 1, YPlot * (SqHeight + 1) + 1)-((XPlot + 1) * (SqWidth + 1) - 1, (YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  'frmMain.picMain.Line ((XFocus * 2 - XPlot) * (SqWidth + 1) + 1, YPlot * (SqHeight + 1) + 1)-((XFocus * 2 - XPlot + 1) * (SqWidth + 1) - 1, (YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  'frmMain.picMain.Line (XPlot * (SqWidth + 1) + 1, (YFocus * 2 - YPlot) * (SqHeight + 1) + 1)-((XPlot + 1) * (SqWidth + 1) - 1, (YFocus * 2 - YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  'frmMain.picMain.Line ((XFocus * 2 - XPlot) * (SqWidth + 1) + 1, (YFocus * 2 - YPlot) * (SqHeight + 1) + 1)-((XFocus * 2 - XPlot + 1) * (SqWidth + 1) - 1, (YFocus * 2 - YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  PixArray(XPlot, YPlot) = SelectedColor
  PixArray(Int(XFocus * 2 - XPlot + 0.5), YPlot) = SelectedColor
  PixArray(XPlot, YFocus * 2 - YPlot) = SelectedColor
  PixArray(Int(XFocus * 2 - XPlot + 0.5), Int(YFocus * 2 - YPlot + 0.5)) = SelectedColor
  Else
  For A% = Int(XFocus * 2 - XPlot + 0.5) To XPlot
  PixArray(A%, YPlot) = SelectedColor
  PixArray(A%, YFocus * 2 - YPlot) = SelectedColor
  'frmMain.picMain.Line (A% * (SqWidth + 1) + 1, YPlot * (SqHeight + 1) + 1)-((A% + 1) * (SqWidth + 1) - 1, (YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  'frmMain.picMain.Line (XPlot * (SqWidth + 1) + 1, (YFocus * 2 - YPlot) * (SqHeight + 1) + 1)-((XPlot + 1) * (SqWidth + 1) - 1, (YFocus * 2 - YPlot + 1) * (SqHeight + 1) - 1), Pal(SelectedColor), BF
  Next A%
End If
End Sub
