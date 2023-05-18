
Advanced ZIP Password Recovery 2.44
==========================================
(c) 1999 Elcom Ltd (V.Katalov, A.Malyshev)


Contents
--------

  Introduction
  Requirements
  Known bugs and limitations
  Where to get the latest version


Introduction
------------

This program (Advanced ZIP Password Recovery, or simply AZPR)
can be used to recover the lost password for a ZIP archive. At
the moment, there is no known method to extract the password
from the compressed file, so the only available methods are
"brute force" and dictionary-based attacks.

Well, there are a lot of programs like this around, but
all of them have their own "pros" and "cons". Here is a brief
list of AZPR's advantages:

- The program has a convenient user interface
- The program is very fast: more than one million passwords per
  second (on Pentium II)
- The program can work with archives containing only one 
  encrypted file
- Self-extracting archives are supported
- The program is customizable: you can set the password length
  (or length range), the character set to be used to generate
  the passwords, and a couple of other options
- You can select the custom character set for brute-force attack
  (non-English characters are supported)
- Dictionary-based attack is available
- The "brute-force with mask" attack is available
- The maximum password length is not limited (in registered
  version)
- No special virtual memory requirements
- You can interrupt the program at any time, and start from the
  same point later
- The program can work in the background, using the CPU only 
  when it is in idle state

The next versions will have much more useful features, of
course.


Requirements
------------

- Windows 95 (any version), Windows 98, Windows NT 4.0 or
  Windows 2000 running on Pentium CPU
- 4 megabytes of RAM
- about 1 megabyte of hard disk space


Known bugs and limitations
--------------------------

- When the files in archive are "stored" (without compression,
  just with encryption) -- the performance might be lower than
  expected (especially on large files), because decrypting the
  whole file is required.
- If the archive contains two or more encrypted files, the
  program assumes that all of them are encrypted with the same
  password.


Where to get the latest version
-------------------------------

The latest version of AZPR is always available from our web page
at http://www.elcomsoft.com/azpr.html. Other password recovery
products (for ARJ archives, Microsoft Access 95/97 databases,
Microsoft Word/Excel (all versions) and Windows NT are available
from our server at:

http://www.elcomsoft.com/prs.html

If you'd like to receive notifications about new releases of AZPR,
please subscribe to our mailing list; mailing list archive and
subscription information is available at:

http://azpr.listbot.com
