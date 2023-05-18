
basRegistry.bas

Updating the Windows registry.

Updated the routine to return a TRUE or FALSE for a key delete.

Perform the four basic functions on the Windows registry using
string or numeric data.
           
           Add
           Change
           Delete
           Query

Important:     All Key data strings should be treated as case
               sensitive.  Always backup your registry 
               (System.dat and User.dat) before performing any
               type of updates.

Software developers vary on where they want to update the
registry with their particular information.  The most common
are in HKEY_lOCAL_MACHINE or HKEY_CURRENT_USER.

This has been tested on Windows 95/98 and Windows NT 4.0. 

This was NT tested by Brett Gerhardi (Brett.Gerhardi@trinite.co.uk)
He found that Windows NT requires that you delete each major
key separately.  Windows 95/98 can delete the top level key
and all the sub level keys in one command.


-----------------------------------------------------------------
Written by Kenneth Ives                    kenaso@home.com

All of my routines have been compiled with VB6 Service Pack 3.
There are several locations on the web to obtain these
modules.

Whenever I use someone else's code, I will give them credit.  
This is my way of saying thank you for your efforts.  I would
appreciate the same consideration.

Read all of the documentation within this program.  It is very
informative.  Also, if you learn to document properly now, you
will not be scratching your head next year trying to figure out
exactly what you were programming today.  Been there, done that.

This software is FREEWARE. You may use it as you see fit for 
your own projects but you may not re-sell the original or the 
source code. If you redistribute it you must include this 
disclaimer and all original copyright notices. 

No warranty express or implied, is given as to the use of this
program. Use at your own risk.

If you have any suggestions or questions, I'd be happy to
hear from you.
-----------------------------------------------------------------
 