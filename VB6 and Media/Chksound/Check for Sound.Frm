VERSION 4.00
Begin VB.Form Sound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check for Sound"
   ClientHeight    =   735
   ClientLeft      =   4170
   ClientTop       =   2685
   ClientWidth     =   3375
   Height          =   1140
   Icon            =   "Check for Sound.frx":0000
   Left            =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   3375
   Top             =   2340
   Width           =   3495
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Check 
      Caption         =   "Check for &Sound Card..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Sound"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()

Unload Me 'exit the program

End Sub


Private Sub Check_Click()

Dim rtn As Integer 'declare the needed variables

rtn = waveOutGetNumDevs() 'check for a sound card

If rtn = 1 Then 'if returned is greater than 1 then a sound card exists
   MsgBox "Your system supports a sound card."
Else 'otherwise no sound card found
   MsgBox "Your system cannot play Sound Files."
End If
    
End Sub


Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

'This project was downloaded from
'
'http://www.brianharper.demon.co.uk/
'
'Please use this project and all of its source code however you want.
'
'UNZIPPING
'To unzip the project files you will need a 32Bit unzipper program that
'can handle long file names. If you have a latest copy of Winzip installed
'on your system then you may use that. If you however dont have a copy,
'then visit my web site, go into the files section and from there you can
'click on the Winzip link to goto their site and download a copy of the
'program. By doing this you will now beable to unzip the project files
'retaining their proper long file names.
'Once upzipped, load up your copy of Visual Basic and goto
'File/Open Project. Locate the project files to where ever you unzipped
'them, then click Open. The project files will be loaded and are now ready
'for use.
'
'THE PROJECT
'If your making a program to use Windows sound functions, then first of all
'your going to need to know whether or not the user using your application
'has a sound card installed. To do this it is quite simple. A function to
'the API can tell you whether the system has a sound card installed or not.
'It requires little code to do this, so makes it very simple.
'
'NOTES
'I have only provided the necessary project files with the zip. This keeps
'the size of the zip files down to a minimum and enables me to upload more
'prjects files to my site.
'
'I hope you find the project usful in what ever you are programming. I
'have tried to write out a small explanation of what each line of code
'does in the project, although most of it is pretty simple to understand.
'
'If you find any bugs in the code then please dont hesitate to Email me and
'I will get back to you as soon as possible. If you however need help on a
'different matter concerning Visual Basic then please please Email me as
'I like to here from people and here what they are programming.
'
'My Email address is:
'Brian@brianharper.demon.co.uk
'
'My web site is:
'http://www.brianharper.demon.co.uk/
'
'Please visit my web site and find many other useful projects like this.
'

End Sub


