VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tag Replace"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Version"
      Height          =   855
      Left            =   30
      TabIndex        =   25
      Top             =   1110
      Width           =   1260
      Begin VB.CheckBox Ver 
         Caption         =   "Ver 2"
         Height          =   270
         Index           =   2
         Left            =   75
         TabIndex        =   27
         Top             =   480
         Width           =   1035
      End
      Begin VB.CheckBox Ver 
         Caption         =   "Ver 1"
         Height          =   285
         Index           =   1
         Left            =   75
         TabIndex        =   26
         Top             =   225
         Width           =   1065
      End
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   374
      Left            =   0
      TabIndex        =   24
      Top             =   1980
      Width           =   4741
      _ExtentX        =   8361
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8308
            Text            =   "Progress"
            TextSave        =   "Progress"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Found"
      Height          =   3315
      Left            =   80
      TabIndex        =   9
      Top             =   2040
      Width           =   4410
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   253
         Left            =   121
         Picture         =   "frmReplace.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         ToolTipText     =   "Drag to Copy or Add to Playlists. (Good for Adding to CD Burning Software)"
         Top             =   242
         Width           =   253
      End
      Begin VB.TextBox Comment 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   960
         TabIndex        =   23
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Title 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   960
         TabIndex        =   22
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Artist 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   960
         TabIndex        =   21
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Album 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   960
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox Genre 
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Text            =   "Genre"
         Top             =   1800
         Width           =   3225
      End
      Begin VB.TextBox Year 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   960
         MaxLength       =   4
         TabIndex        =   16
         Top             =   2160
         Width           =   810
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Change"
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Filename 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Lablle 
         Caption         =   "Genre"
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label Label6 
         Caption         =   "Year"
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Comment"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Title"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Artist"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Album"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox tReplace 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Find 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Comment"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Title"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Artist"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Album"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   253
      Left            =   1452
      OleObjectBlob   =   "frmReplace.frx":030A
      TabIndex        =   28
      Top             =   242
      Width           =   1100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   253
      Left            =   1452
      OleObjectBlob   =   "frmReplace.frx":0370
      TabIndex        =   29
      Top             =   968
      Width           =   1100
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MinHeight As Long 'When Smaller
Dim MaxHeight As Long 'When Taller
Const CompleteMsg = "Finished Searching Through Song Tags" 'Completed Message shown in a Message Box when Searching is all done
Dim Current As Integer
Dim ID3Tag As cIDV3

Private Sub Command1_Click()
    'Close this Window
    Unloading = True
    DoEvents
    Unload Me
End Sub

Private Sub Command2_Click()
    'Find the Tag
    '(Pressing Enter Activates this Button)
    Dim Found As Boolean
    Dim I As Integer
    Unloading = True
    DoEvents 'Do everything else with the Unloading = true so it will Stop everything
    Unloading = False
    With frmTagger 'Everything with .Something will access just like frmTagger.Something
        'Are there files
        If .ListView1.ListItems.Count = 0 Then MsgBox CompleteMsg: Exit Sub: Me.Height = MinHight: Frame1.Visible = False
        'there are Files
        If .ListView1.MultiSelect Then 'Code for Multiselect
            For I = Current To .ListView1.ListItems.Count 'Look through them all
                'Unfortunately I have found no better way for Multi then to add a For Loop and Check them all.
                
                If .ListView1.ListItems(I).Selected Then 'Check if its selected
                    Filename = .ListView1.ListItems(I).SubItems(1) & .ListView1.ListItems(I).Text
                    StatBar.Panels(1).Text = "Searching " & .ListView1.ListItems(I).Text
                    ID3Tag.Filename = Filename 'Activate the ID3Tag Module
                    If Ver(1).Value = vbChecked Then 'Check if the Version 1 Check is checked
                        ID3Tag.ID3v2 = False 'Version is 1 therefore ID3v2 = False
                        ID3Tag.ReadTag 'Read the tag
                        
                        'If they are checked then Look in them and see if you can find the tag being looked for
                        If Check1(0).Value = vbChecked Then If InStr(1, ID3Tag.Album, Find, vbTextCompare) > 0 Then Found = True
                        If Check1(1).Value = vbChecked And Found = False Then If InStr(1, ID3Tag.Artist, Find, vbTextCompare) > 0 Then Found = True
                        If Check1(2).Value = vbChecked And Found = False Then If InStr(1, ID3Tag.Title, Find, vbTextCompare) > 0 Then Found = True
                        If Check1(3).Value = vbChecked And Found = False Then If InStr(1, ID3Tag.Comments, Find, vbTextCompare) > 0 Then Found = True
                    End If
                    If Ver(2).Value = vbChecked And Found = False Then 'See if the Version 2 Checkbox is checked and Nothing has been found yet
                        ID3Tag.ID3v2 = True 'Version is 2
                        ID3Tag.ReadTag 'Read the tag
                        
                        'If they are checked then Look in them and see if you can find the tag being looked for
                        If Check1(0).Value = vbChecked Then If InStr(1, ID3Tag.Album, Find, vbTextCompare) > 0 Then Found = True
                        If Check1(1).Value = vbChecked And Found = False Then If InStr(1, ID3Tag.Artist, Find, vbTextCompare) > 0 Then Found = True
                        If Check1(2).Value = vbChecked And Found = False Then If InStr(1, ID3Tag.Title, Find, vbTextCompare) > 0 Then Found = True
                        If Check1(3).Value = vbChecked And Found = False Then If InStr(1, ID3Tag.Comments, Find, vbTextCompare) > 0 Then Found = True
                    End If
                    If Found Then 'If you found something in this file
                        ID3Tag.ReadTag 'Read the Tag
                        'And Display the information in the Text Boxes
                        Album = ID3Tag.Album
                        Artist = ID3Tag.Artist
                        Title = ID3Tag.Title
                        If MP3Tag.ID3v2 Then
                            Genre.Text = ID3Tag.GenreV2
                        Else
                            Genre.ListIndex = ID3Tag.GenreV1
                        End If
                        
                        Year = ID3Tag.Year
                        Comment = ID3Tag.Comments
                        Current = I + 1
                        Me.Height = MaxHeight
                        Frame1.Visible = True
                        Exit Sub
                    End If
                End If
                If Unloading Then Exit Sub 'Unload from the Unloading Variable
                DoEvents
            Next I
            MsgBox CompleteMsg 'Display the Complete Message
            Me.Height = MinHeight
            Frame1.Visible = False
            Current = 1
        Else
            With frmTagger 'Set the Current Data on the frmTagger to update Display
                If Check1(0).Value = vbChecked Then .Album = Replace(.Album, Find, tReplace, , , vbTextCompare)
                If Check1(1).Value = vbChecked Then .Artist = Replace(.Artist, Find, tReplace, , , vbTextCompare)
                If Check1(2).Value = vbChecked Then .Title = Replace(.Title, Find, tReplace, , , vbTextCompare)
                If Check1(3).Value = vbChecked Then .Comment = Replace(.Comment, Find, tReplace, , , vbTextCompare)
                .Command3.Value = True
            End With
        End If
    End With
End Sub

Private Sub Command3_Click()
'
If LCase(Right(Filename, 4)) = ".mp3" Then 'check the extention
    If Dir(Filename) <> "" Then 'See if the file exists
        If Check1(0).Value = vbChecked Then
            'Replace the String if it exists
            Album = Replace(Album, Find, tReplace, , , vbTextCompare)
        End If
        If Check1(1).Value = vbChecked Then
            'Replace the String if it exists
            Artist = Replace(Artist, Find, tReplace, , , vbTextCompare)
        End If
        If Check1(2).Value = vbChecked Then
            'Replace the String if it exists
            Title = Replace(Title, Find, tReplace, , , vbTextCompare)
        End If
        If Check1(3).Value = vbChecked Then
            'Replace the String if it exists
            Comment = Replace(Comment, Find, tReplace, , , vbTextCompare)
        End If
    End If
End If
    Command4_Click
    Command2_Click
    
End Sub

Private Sub Command4_Click()
    With ID3Tag 'With ID3Tag so every .Something would be ID3Tag.Something
        If Ver(2).Value = vbChecked Then 'If Version 2 is Included
            .ID3v2 = True 'Its Version 2 Therefore True
            .Filename = Filename    'Set the File
            .ReadTag                'Read the Tag
                                    'Set all the Variables
            .Artist = Artist
            .Album = Album
            .Title = Title
            .GenreV2 = Genre.Text
            .Year = Year
            .Comments = Comment
            .WriteTag               'Write the tag
        End If
        
        If Ver(1).Value = vbChecked Then
            .ID3v2 = False 'Version is not 2.. Therefore false
            
                                    'Same as last time
            .Filename = Filename
            .ReadTag
            .Artist = Artist
            .Album = Album
            .Title = Title
            .GenreV1 = Genre.ListIndex
            .Year = Year
            .Comments = Comment
            .WriteTag
        End If
    End With
End Sub

Private Sub Find_Change()
Current = 1
'Set the Buttons enabled or disabled for the Matching criteria
If Len(Find) = 0 Or tReplace = Find Then
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Command2.Enabled = True
    Command3.Enabled = True
End If
End Sub

Private Sub Form_Load()
Current = 1
Unloading = False
Set ID3Tag = New cIDV3 'Set the Class Module to Activate it
    With frmTagger
        Ver(1).Value = .cVer(1).Value
        Ver(2).Value = .cVer(2).Value
        If .ListView1.MultiSelect Then 'Select the same as the Multiselect Checks
            Check1(0).Value = .Check2(0).Value
            Check1(1).Value = .Check2(1).Value
            Check1(2).Value = .Check2(2).Value
            Check1(3).Value = .Check2(5).Value
        Else
            Command3.Visible = False ' No need for just Find so Replace is the only thing there
            Command2.Caption = "Replace" 'Reset the Button to Replace
        End If
    End With
    'Activate the Genre Combo
    Genre.Clear
    Genre.AddItem "[Blues]"
    Genre.AddItem "[Classic Rock]"
    Genre.AddItem "[Country]"
    Genre.AddItem "[Dance]"
    Genre.AddItem "[Disco]"
    Genre.AddItem "[Funk]"
    Genre.AddItem "[Grunge]"
    Genre.AddItem "[Hip-Hop]"
    Genre.AddItem "[Jazz]"
    Genre.AddItem "[Metal]"
    Genre.AddItem "[New Age]"
    Genre.AddItem "[Oldies]"
    Genre.AddItem "[Other]"
    Genre.AddItem "[Pop]"
    Genre.AddItem "[R&B]"
    Genre.AddItem "[Rap]"
    Genre.AddItem "[Reggae]"
    Genre.AddItem "[Rock]"
    Genre.AddItem "[Techno]"
    Genre.AddItem "[Industrial]"
    Genre.AddItem "[Alternative]"
    Genre.AddItem "[Ska]"
    Genre.AddItem "[Death Metal]"
    Genre.AddItem "[Pranks]"
    Genre.AddItem "[Soundtrack]"
    Genre.AddItem "[Euro-Techno]"
    Genre.AddItem "[Ambient]"
    Genre.AddItem "[Trip-Hop]"
    Genre.AddItem "[Vocal]"
    Genre.AddItem "[Jazz+Funk]"
    Genre.AddItem "[Fusion]"
    Genre.AddItem "[Trance]"
    Genre.AddItem "[Classical]"
    Genre.AddItem "[Instrumental]"
    Genre.AddItem "[Acid]"
    Genre.AddItem "[House]"
    Genre.AddItem "[Game]"
    Genre.AddItem "[Sound Clip]"
    Genre.AddItem "[Gospel]"
    Genre.AddItem "[Noise]"
    Genre.AddItem "[Alt. Rock]"
    Genre.AddItem "[Bass]"
    Genre.AddItem "[Soul]"
    Genre.AddItem "[Punk]"
    Genre.AddItem "[Space]"
    Genre.AddItem "[Meditative]"
    Genre.AddItem "[Instrumental Pop]"
    Genre.AddItem "[Instrumental Rock]"
    Genre.AddItem "[Ethnic]"
    Genre.AddItem "[Gothic]"
    Genre.AddItem "[Darkwave]"
    Genre.AddItem "[Techno-Industrial]"
    Genre.AddItem "[Electronic]"
    Genre.AddItem "[Pop-Folk]"
    Genre.AddItem "[Eurodance]"
    Genre.AddItem "[Dream]"
    Genre.AddItem "[Southern Rock]"
    Genre.AddItem "[Comedy]"
    Genre.AddItem "[Cult]"
    Genre.AddItem "[Gangsta Rap]"
    Genre.AddItem "[Top 40]"
    Genre.AddItem "[Christian Rap]"
    Genre.AddItem "[Pop/Funk]"
    Genre.AddItem "[Jungle]"
    Genre.AddItem "[Native American]"
    Genre.AddItem "[Cabaret]"
    Genre.AddItem "[New Wave]"
    Genre.AddItem "[Phychedelic]"
    Genre.AddItem "[Rave]"
    Genre.AddItem "[Showtunes]"
    Genre.AddItem "[Trailer]"
    Genre.AddItem "[Lo-Fi]"
    Genre.AddItem "[Tribal]"
    Genre.AddItem "[Acid Punk]"
    Genre.AddItem "[Acid Jazz]"
    Genre.AddItem "[Polka]"
    Genre.AddItem "[Retro]"
    Genre.AddItem "[Musical]"
    Genre.AddItem "[Rock & Roll]"
    Genre.AddItem "[Hard Rock]"
    Genre.AddItem "[Folk]"
    Genre.AddItem "[Folk/Rock]"
    Genre.AddItem "[National Folk]"
    Genre.AddItem "[Swing]"
    Genre.AddItem "[Fast-Fusion]"
    Genre.AddItem "[Bebob]"
    Genre.AddItem "[Latin]"
    Genre.AddItem "[Revival]"
    Genre.AddItem "[Celtic]"
    Genre.AddItem "[Bluegrass]"
    Genre.AddItem "[Avantegarde]"
    Genre.AddItem "[Gothic Rock]"
    Genre.AddItem "[Progressive Rock]"
    Genre.AddItem "[Psychedelic Rock]"
    Genre.AddItem "[Symphonic Rock]"
    Genre.AddItem "[Slow Rock]"
    Genre.AddItem "[Big Band]"
    Genre.AddItem "[Chorus]"
    Genre.AddItem "[Easy Listening]"
    Genre.AddItem "[Acoustic]"
    Genre.AddItem "[Humour]"
    Genre.AddItem "[Speech]"
    Genre.AddItem "[Chanson]"
    Genre.AddItem "[Opera]"
    Genre.AddItem "[Chamber Music]"
    Genre.AddItem "[Sonata]"
    Genre.AddItem "[Symphony]"
    Genre.AddItem "[Booty Bass]"
    Genre.AddItem "[Primus]"
    Genre.AddItem "[Porn Groove]"
    Genre.AddItem "[Satire]"
    Genre.AddItem "[Slow Jam]"
    Genre.AddItem "[Club]"
    Genre.AddItem "[Tango]"
    Genre.AddItem "[Samba]"
    Genre.AddItem "[Folklore]"
    Genre.AddItem "[Ballad]"
    Genre.AddItem "[Power Ballad]"
    Genre.AddItem "[Rhythmic Soul]"
    Genre.AddItem "[Freestyle]"
    Genre.AddItem "[Duet]"
    Genre.AddItem "[Punk Rock]"
    Genre.AddItem "[Drum Solo]"
    Genre.AddItem "[A Capella]"
    Genre.AddItem "[Euro-House]"
    Genre.AddItem "[Dance Hall]"
    Genre.AddItem "[Goa]"
    Genre.AddItem "[Drum & Bass]"
    Genre.AddItem "[Club-House]"
    Genre.AddItem "[Hardcore]"
    Genre.AddItem "[Terror]"
    Genre.AddItem "[Indie]"
    Genre.AddItem "[BritPop]"
    Genre.AddItem "[Negerpunk]"
    Genre.AddItem "[Polsk Punk]"
    Genre.AddItem "[Beat]"
    Genre.AddItem "[Christian Gangsta Rap]"
    Genre.AddItem "[Heavy Metal]"
    Genre.AddItem "[Black Metal]"
    Genre.AddItem "[Crossover]"
    Genre.AddItem "[Contemporary Christian]"
    Genre.AddItem "[Christian Rock]"
    Genre.AddItem "[Merengue]"
    Genre.AddItem "[Salsa]"
    Genre.AddItem "[Trash Metal]"
    Genre.AddItem "[Anime]"
    Genre.AddItem "[JPop]"
    Genre.AddItem "[Synthpop]"
frmTagger.Skin1.ApplySkin Me.hwnd 'Skin the Form

'Auto Size the Components
MinHeight = Me.Height + (100 + Command2.Top + Command2.Height + StatBar.Height - Me.ScaleHeight)
MaxHeight = Me.Height + (60 + Frame1.Top + Frame1.Height + StatBar.Height - Me.ScaleHeight)
Me.Height = MinHeight
Me.Width = Me.Width + (100 + Command2.Left + Command2.Width - Me.ScaleWidth)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unloading = True 'Unloading = True to try to stop Everything thats Running in the back
    DoEvents 'Doevents to Allow it to find the Unloading = True
    DoEvents 'Once more just for good measure
    Unload Me 'Unload
End Sub

Private Sub Picture1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    'Draging the Notes will allow you to Copy it to a folder
    'Usefull for adding to CD Burning Software, Playlists or whatever
    AllowedEffects = vbDropEffectCopy
    Data.SetData , vbCFFiles
    Data.Files.Add Filename.Text
End Sub


Private Sub tReplace_Change()
'When replace changes check to see if the Buttons should be disabled
If Len(Find) = 0 Or tReplace = Find Then
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Command2.Enabled = True
    Command3.Enabled = True
End If
End Sub
