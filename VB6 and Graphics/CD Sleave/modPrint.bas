Attribute VB_Name = "modPrint"
'####################################################
'#      ALL code was created by Andy McCurtin       #
'#  This code took me a long time to perfect (and a #
'#  rainforests worth of paper)                     #
'#  It was not created to be used for any illegal   #
'#  purposes, however what you do with it is your   #
'#  business and I take no responsibility for your  #
'#  actions.                                        #
'#  If you improve this code please send me a copy  #
'#    (How often do people actually do this ????)   #
'#                                                  #
'#  Please do not plagiarise my code, I like many   #
'#  other people spend time on our programs so      #
'#  others can learn from them and when brainless   #
'#  Wa**ers steal our work and put their names on   #
'#  it and then have the audacity to take credit    #
'#  for it being their own, it it makes you think   #
'#  twice before giving any more code away incase   #
'#  these idiots steal that as well.                #
'#  (As you may have gathered I have had my work    #
'#  plagiarised in the past and I had to think long #
'#  and hard before uploading this code)            #
'#                                                  #
'#  Anyway enough preaching (I've got WAY TOO much  #
'#  tequila and vodka in my system)                 #
'#  Any problems with this code please e-mail me    #
'#  Any questions please e-mail, I'll try to help   #
'#                                                  #
'#              andy_mccurtin@yahoo.com             #
'#                                                  #
'#  It's 02:02am GMT I've been drinking way too     #
'#  long, and can hardly see straight so I leave    #
'#  you with these words of wisdom...               #
'#  U're my bestest mate in d wurld!!!!! and I lurv #
'#  ya.......... DRROOOOOLLLL... ZZZZZZZZZZZZZZZ    #
'#                                                  #
'#  P.S  I really should stop drinking when I       #
'#  program                                         #
'####################################################


'//When printing an image the format I have used is as
'//follows :-
'//Printer.PaintPicture Image1.Picture, X1, Y1, Width1, _
'//     Height1
'//
'// Image1 stretch proprty should be set to True
'// X1 is the Left margin (how big a space from the left
'// edge of the paper)
'// Y1 is the Top margin (how big a space from the top
'// of the page)
'// Width1 is the width of the image being printed
'// Height1 is the height of the image being printed
'//
'// When you see frmMain!img###### this allows you to use
'// a control on Form frmMain (you can call any control)
'//
'// 1440 Twips = (Approx) 1 inch
'//
'// CD cover measurments are as follows :-
'// Front Cover = W6900, H6900
'// Inside Cover = W6900, H6900
'// Front & Inside Cover as one = W13800 , H6900
'// Back Cover = W8530, H6700

Option Explicit
Public prnCopies As Integer

'//Print ONLY front cover
Public Sub PrintFront()
'Variables
Dim TopMargin As Integer
Dim LeftMargin As Integer
Dim Width As Integer
Dim Height As Integer
        
    'assigns values (in twips) to variables
    TopMargin = 1440
    LeftMargin = 1440
    Width = 6900
    Height = 6900
          
    Printer.Orientation = 1 'Prints in protrait
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
   
    Printer.PaintPicture frmMain!imgFront.Picture, _
        LeftMargin, TopMargin, Width, Height
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub

'//Print ONLY back cover
Public Sub PrintBack()
'Variables
Dim bTopMargin As Integer
Dim bLeftMargin As Integer
Dim bWidth As Integer
Dim bHeight As Integer
    
    'assigns values (in twips) to variables
    bTopMargin = 1440
    bLeftMargin = 1440
    bWidth = 8530
    bHeight = 6700
    
    Printer.Orientation = 1 'Prints in Portrait
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgBack.Picture, bLeftMargin, _
        bTopMargin, bWidth, bHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print ONLY front and inside (separate) covers
Public Sub PrintFrontAndInside()
'Variables
Dim FrontTopMargin As Integer
Dim FrontLeftMargin As Integer
Dim FrontWidth As Integer
Dim FrontHeight As Integer
Dim InsideTopMargin As Integer
Dim InsideLeftMargin As Integer
Dim InsideWidth As Integer
Dim InsideHeight As Integer
    
    'assigns values (in twips) to variables
    FrontTopMargin = 1000
    FrontLeftMargin = 6900 + 1010
    FrontWidth = 6900
    FrontHeight = 6900
    
    'assigns values (in twips) to variables
    InsideTopMargin = 1000
    InsideLeftMargin = 1000
    InsideWidth = 6900
    InsideHeight = 6900
    
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFrontS.Picture, FrontLeftMargin, _
    FrontTopMargin, FrontWidth, FrontHeight
    
    Printer.PaintPicture frmMain!imgInsideS.Picture, InsideLeftMargin, _
        InsideTopMargin, InsideWidth, InsideHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print Front and Inside (Whole) cover
Public Sub PrintWhole()
'Variables
Dim WholeTopMargin As Integer
Dim WholeLeftMargin As Integer
Dim WholeWidth As Integer
Dim WholeHeight As Integer
    
    'assigns values (in twips) to variables
    WholeTopMargin = 1440
    WholeLeftMargin = 1440
    WholeWidth = 13800
    WholeHeight = 6900
    
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFront_Inside_W.Picture, WholeLeftMargin, _
    WholeTopMargin, WholeWidth, WholeHeight

    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print ONLY Front and Back covers
Public Sub PrintFrontAndBack()
'Variables
Dim FrontTopMargin As Integer
Dim FrontLeftMargin As Integer
Dim FrontWidth As Integer
Dim FrontHeight As Integer
Dim BackTopMargin As Integer
Dim BackLeftMargin As Integer
Dim BackWidth As Integer
Dim BackHeight As Integer

    'assigns values (in twips) to variables
    FrontTopMargin = 1000
    FrontLeftMargin = 1440 + 720
    FrontWidth = 6900
    FrontHeight = 6900
    
    'assigns values (in twips) to variables
    BackTopMargin = 6900 + 1440
    BackLeftMargin = 1440
    BackWidth = 8530
    BackHeight = 6700
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFront.Picture, FrontLeftMargin, _
    FrontTopMargin, FrontWidth, FrontHeight
    
    Printer.PaintPicture frmMain!imgBack.Picture, BackLeftMargin, _
        BackTopMargin, BackWidth, BackHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
End Sub


'//Print Front and Inside (Separate) and Back covers
Public Sub PrintFrontsAndInsideSAndBack()
'Fron and Inside Variables
Dim FrontTopMargin As Integer
Dim FrontLeftMargin As Integer
Dim FrontWidth As Integer
Dim FrontHeight As Integer
Dim InsideTopMargin As Integer
Dim InsideLeftMargin As Integer
Dim InsideWidth As Integer
Dim InsideHeight As Integer
'Back Variables
Dim bTopMargin As Integer
Dim bLeftMargin As Integer
Dim bWidth As Integer
Dim bHeight As Integer
 
    
    'assigns values (in twips) to variables
    FrontTopMargin = 1000
    FrontLeftMargin = 6900 + 1010
    FrontWidth = 6900
    FrontHeight = 6900
    
    'assigns values (in twips) to variables
    InsideTopMargin = 1000
    InsideLeftMargin = 1000
    InsideWidth = 6900
    InsideHeight = 6900
    
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFrontS.Picture, FrontLeftMargin, _
    FrontTopMargin, FrontWidth, FrontHeight
    
    Printer.PaintPicture frmMain!imgInsideS.Picture, InsideLeftMargin, _
        InsideTopMargin, InsideWidth, InsideHeight
    
    Printer.NewPage 'Begins new page for back cover
        
    'assigns values (in twips) to variables
    bTopMargin = 1440
    bLeftMargin = 1440
    bWidth = 8530
    bHeight = 6700
    
    Printer.Orientation = 1 'Prints in Portrait
    
    Printer.PaintPicture frmMain!imgBack.Picture, bLeftMargin, _
        bTopMargin, bWidth, bHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
    
End Sub


'//Print Front and Inside (whole) and Back covers
Public Sub PrintWholeAndBack()
'Whole Variables
Dim WholeTopMargin As Integer
Dim WholeLeftMargin As Integer
Dim WholeWidth As Integer
Dim WholeHeight As Integer
    
'Back Variables
Dim bTopMargin As Integer
Dim bLeftMargin As Integer
Dim bWidth As Integer
Dim bHeight As Integer
    
    
    'assigns values (in twips) to variables
    WholeTopMargin = 1440
    WholeLeftMargin = 1440
    WholeWidth = 13800
    WholeHeight = 6900
        
    Printer.Orientation = 2 'Prints in Landscape
    
    If prnCopies = 0 Then
        Printer.copies = 1
    Else
        Printer.copies = prnCopies 'no of copies to print
    End If
    
    Printer.PaintPicture frmMain!imgFront_Inside_W.Picture, WholeLeftMargin, _
    WholeTopMargin, WholeWidth, WholeHeight
    
    Printer.NewPage 'Begins new page for back cover
    
    'assigns values (in twips) to variables
    bTopMargin = 1440
    bLeftMargin = 1440
    bWidth = 8530
    bHeight = 6700
    
    Printer.Orientation = 1 'Prints in Portrait
    
    Printer.PaintPicture frmMain!imgBack.Picture, bLeftMargin, _
        bTopMargin, bWidth, bHeight
    
    Printer.EndDoc 'without this nothing will print until the program closes
    
End Sub

