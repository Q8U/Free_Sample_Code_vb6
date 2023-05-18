The VBUMidiPlayer control is Freeware.  You may 
distribute it freely with your applications.  You are free 
to distribute copies of the control by itself as long as 
you also distribute the files, unmodified, that came with
this control.  These files are listed below.

The VBUMidiPlayer is a light-weight, easy to use control.
It will let you play a midi file just by setting the
filename property and calling the MidiPlay method. You
also can set the PlayFrom and PlayTo properties if you only
want to play a portion of the file.  Also this control has
a Completed event to let you know when the file is done
playing, and an MidiError Event when an error occurs
while trying to play the Midi file.
The VBUMidiPlayer also contains a Repeat property which, 
when set to true, will cause the file to repeat as soon
as it is finished.

NOTE:
If you exit your application by pressing the stop button
on the VB Toolbar while the music is playing, the music
will continue to play.  The same is true if you use the 
End command while developing in the VB environment.  This
is because the terminate event in the Midi control never
gets called when you exit that way.  This is not a serious
problem. All you will need to do to stop the music is to
run your program again and call the MidiStop method.

The VBUMidiPlayer control was developed with 
Visual Basic 5.0  This means that you will need the VB 5 
runtime files. I recommend using this control with VB 5
(Mainly because I haven't tested it under any other
development language.)

To register this control just run the Install.exe file.
This install will let you copy the VBUMidiPlayer to your
system directory and register it for you.

If the install does not work then try this:
open up VB and go to the screen where you pick your 
components (In VB5 it is under the Projects|Components Menu).
Click on the browse button. Now find the VBUMidiPlayer.OCX
and click OK.  Thats it.  Just remember to check the X box 
next to the component name to add it to your project.

These are the files that you should have received with this
control.

     VBUMidiPlayer.ocx
     ReadMe.txt  (What your reading right now.)
     Install.exe


DISCLAIMER:
-----------
THIS CONTROL IS PROVIDED AS IS AND WITHOUT ANY WARRANTIES 
WHATSOEVER.  THE USER IS ADVISED TO THOROUGHLY TEST AND DEBUG
THEIR PROGRAM.  THE USER ASSUMES THE ENTIRE RISK OF USING
THE CONTROL.  THE AUTHOR OF THIS CONTROL IS UNDER NO LEGAL
RESPONSIBILTY.

If you have any questions or comments then please send me an
EMail to: 
larryallen@geocities.com   or

Visit My Website at:
Larry Allen's Visual Basic Universe
www.geocities.com/SiliconValley/Horizon/8806

Thank You
