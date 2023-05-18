CDNotification ActiveX Control v1.0

Copyright (c) 1999 Hai Li, Zeal SoftStudio.
Release Date: August 10, 1999
Homepage: http://members.tripod.com/~zealsoft/
          http://www.nease.net/~zealsoft/indexc.html
EMail: haili@public.bta.net.cn

This is an English Read Me file. Chinese people please 
read ReadMeC.txt.

CDNotification ActiveX Control is a CARDWARE. If you 
are using this program, please send me a postcard, 
not e-mail. My address is
	Hai Li
	No. 1607 Unit 133
	Beijing Institute of Technology
	Beijing 100081
	PR China

What's CDNotification ActiveX Control
--------------------------------------
CDNotification ActiveX Control v1.0 is an control that
allows your program to know when a CD is inserted in or
ejected out of the CD-Rom drive.

One day, I got the ComonentSource Demo CDs which can 
change its content when I changed the CD. I'm impressed
by it and want to add the feature into my applications. 
If you like it, please send me a postcard and tell me 
why you use it and your suggestion.

If you want to get the source of this control, you can
refer to Buy Source Code section in this file.

Newsletter
--------------
If you want get a notification when we release a new 
version or new free control, you can visit 
http://members.tripod.com/~zealsoft/cdnotify
and subscribe Free Control newsletter.

Install/Uninstall
--------------------
You can unzip all files to your hard disk. 

CDNotify.ocx and all other files in VB5 subfolder are
used with Visual Basic 5.0 Service Pack 3(SP3).

CDNotify6.ocx and all other files in VB6 subfolder are
used with Visual Basic 6.0.

If you want VB 5.0 (SP3) or VB 6 Runtime DLLs, you can 
download them from http://members.tripod.com/~zealsoft/cdnotify.

To uninstall the files, you can simply delete all files
which you copy to your disk.

How to use 
------------

1) Properties

   Enabled property
      Data type: Boolean
      Remarks: When this property is set to True(default), the 
               control will fire corresponding event when a 
               CD-Rom is inserted or ejected out.

2) Events

   Arrival Event 
      Syntax: Private Sub obj_Arrival(ByVal Drive As String)
      Remarks: Occurs when a new CD arrived. The parameter Drive 
               indicates the CD-Rom drive.

   RemoveComplete Event 
      Syntax: Private Sub obj_RemoveComplete(ByVal Drive As String)
      Remarks: Occurs when CD is removed. The parameter Drive indicates
               the CD-Rom drive.


Samples
---------
Visual Basic 5.0 and 6.0 samples are included, which are 
located in Vb5 and Vb6 directory. 

Buy Source Code
-----------------
If you need to know the secret why the contorl work, 
you can visit 
http://members.tripod.com/~zealsoft/cdnotify
to buy the source code of the control(US$10). Both VB 5.0
and VB 6.0 source code are included.

History
---------
1.0	Initial release