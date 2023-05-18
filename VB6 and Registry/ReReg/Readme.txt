############################################################
#                                                          #
# Date: 7-6-2000                                           #
#                                                          #
# ReReg was developed by:                                  #
#                                                          #
# Name:     Bas van de Ree                                 #
# Company:  N.V. Interpolis                                #
# Place:    Tilburg, The Netherlands                       #
# Email:    b.vd.ree@interpolis.nl                         #
#                                                          #
############################################################

* Zipped-files:
frmReReg.frm	- Form that demonstrates the usage of ReReg
frmReReg.frx	- Form's binary file
Readme.txt	- This file
ReReg.vbw	- Demonstration Project workspace
ReReg.vbp	- Demonstration Project
ReReg.exe	- The compiled demonstration
ReReg.ctx	- Binary file for ReReg.ctl
ReReg.ctl	- The ReReg user control

* Purpose
Reload the Windows Registry to initialize changes

* Method
Kill the Explorer process and reinitialize it through API
calls

-= ReReg notes =-

If any errors occur during the reload:
try increasing the Timer interval of the ReReg control

I've tested ReReg on the following systems with success:

Pentium 90    : 80MB RAM ; Windows 98
Pentium II 266; 128MB RAM; Windows 98

If you have any comments, ideas or enhancements in mind,
please let me know. You can reach me by e-mail.

