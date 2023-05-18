----------------------------------------------
| README FILE FOR AiE's VB MP3 PLAYER EXAMPLE |
----------------------------------------------
+++++++++++++++++++++++++++++++++++++++++++++++
-----------------------------------------------

PART-I : SETTING UP YOUR SYSTEM TO PLAY MP3S VIA' ACTIVE-MOVIE 2.X 

PART-II : EXAMPLE OF RW'S SYSTEM.INI

PART-III : CONFIGURING THE PROJECT TO RUN ON YOUR SYSTEM

PART-IV : CONGRADULATIONS

PART-V : AIE'S RANTS ABOUT MAKING THIS AVAILABLE TO YOU. :)

-----------------------------------------------

--------
PREFACE: 
--------

First off I have always been an excellent reader, but a horrible
speller as far as grammer. Now that that part is out of the way we
can officially get down to business.

I First figured out how to do this a long, long time ago in a far
far away coffee house where angels presented me with gifts of Yuban joy
and A really great view of their uhm....

Anyhow.. Since we first released this code a long time ago
"Note it wasn't really promoted" some people have tried to get rich off
of it.. We didn't like seeing such people taking advantage of our kind
gesture to the VB community and immediately halted distribution.

At any rate.. Every now and then I notice a few posts in the VB Usenet groups
asking for pirated copies of Maquisyetms.. "or whatever that 2 man show calls itself"
ActiveX crontols to play MP3's through VB..

anyhow.. ladies and gentelmen.. I have no opinion of such an act
"Especially where these people are concerned" however it still goes
against everything we have worked for as programmers.

At any rate as I said earlier:
"some people have tried to get rich off of it" when it
came to our origional code... is anybody getting this yet?

they are the "some people" I'm referring to "Maquisitem".
And now I have "over a long period of time" considered re-releasing 
this code again.. and due to the major outpour of VB Programmers support
for us to do so.. I Decided that I'd let it back out to the worlds hands.

What many people may, or may not know is that AiE Prides itself on 
being the most innovative group of VB Multimedia Developers in the World.
many of our past accomplishments are recognized by the VB'ers out there
while others like "Maquisystem" will try to take credit for our years
of Experience, and Multimedia Mastery.

We have always been the kind of people to give back to those among us
in the VB community.. and some of you might know about these efforts
we have made to keep that promise.

Basically.. This code is ours, but free for your use.
you may use it Royalty free as long as you don't try to turn around
and re-package it for sell at $300+ bucks like Maquisystem did.

Also this time the Code isn't in ActiveX form.
We have provided her as she origionally was.
However you will need to take certain steps to ensure your
system has been optomized the code and ActiveMovie control.

This example takes advantage of an Exploit that we Designed
that will become clear in the next area.

Before I go.. one last thing.. now that you have the power
and knowledge to do what you've never been able to do before
"play mp3's in VB" please do us a favor, and join us in our
campaign to boycott Maquisystem Midia, and their "so called controls"
that were made from code we provided.

they are the trash of the developer community, and should be
put out of business for trying to take advantage of our kindess
and willingness to give back to the VB community, and for that alone
as far as myself, or my colleagues go Maquisystem is officially blacklisted
from seeing as much as one line of AiE code again.

now.. on to the stuff we've all been waiting for!!!

--------------------------------------------------------------------------
--------------------------------------------------------------------------
PART-I : SETTING UP YOUR SYSTEM TO PLAY MP3S VIA' ACTIVE-MOVIE 2.X 
--------------------------------------------------------------------------
--------------------------------------------------------------------------

It may not come as a surprise to some of you that this is an exploit
for the ActiveMovie 2.x control.. however if it does, do a test and
put one on a form and go to the "FileName" property.. no listing for MP3's eh?

Don't sweat it.. We've got the answer. :-)

you will need to optimize your system by placing a command line
into your system.ini file for Windows95/98 to use the hidden MP3
decoding capabilites of the ActiveMovie control.

before I give you that one like.. be aware that all your users will
need to have their sysetms optimized as-well. but then again we all
know what batchfiles are, right? and there *are* setup tools to edit
files during installations.. so that shouldn't be to hard. :)

here she goes, this is that one line.. 
followed by why it needs to be there and what it means. :) 
---------------------------------

ActiveMovie=mciqtz.drv

---------------------------------

yup.. pretty small line huh?
like I said. it's an exploit. :)

as for the explanations:

1- it calls the "until now" rarely known about compiled franhaufer
codex for decoding mp3 files in the actual control.

2- why do you need to do it? duhh! your un-locking a part of ActiveMovie
that Microsoft didn't want to tell you about.. the powerful part.
and when you add this line into your system.ini file you will be
un-crippeling it everytime you boot up the puter. :)

Why Would MS do such a thing?
well.. uhmnn.. geeze guys.. this is MS we're talking about here
I'm sure you can all come up with 100's of different explanations as to
why MS likes slipping you a plate of bull, and keeping the good stuff quiet.

like I said.. this is MS we're talking about.. enough said. :)

This now brings us to our next section.. where she goes into your system.ini
before you continue you can take a peek at "RW.INI" which is a copy of my system.ini
so you can have a *real* example of what this small line looks like when implemented. 

else you can just scroll down below, and take a look at the following snippet. :)

--------------------------------------------------------------------------
--------------------------------------------------------------------------
PART-II : EXAMPLE OF RW'S SYSTEM.INI
--------------------------------------------------------------------------
--------------------------------------------------------------------------


[iccvid.drv]

[mciseq.drv]

[mci]               <-- when you get here
cdaudio=mcicda.drv
sequencer=mciseq.drv
waveaudio=mciwave.drv
avivideo=mciavi.drv
MPEGVideo=mciatim1.drv
videodisc=mcipionr.drv
vcr=mcivisca.drv
VIDEOCD=mciatim1.drv
ActiveMovie=mciqtz.drv  <-- Paste the line over here

[drivers32]
VIDC.IV41=ir41_32.dll
MSACM.imaadpcm=imaadp32.acm

THE [MCI] SECTION IS WHERE THAT LINE NEEDS TO BE PUT, RIGHT AT THE END. :)


--------------------------------------------------------------------------
--------------------------------------------------------------------------
PART-III : CONFIGURING THE PROJECT TO RUN ON YOUR SYSTEM
--------------------------------------------------------------------------
--------------------------------------------------------------------------

ok.. now that you have the code implemented, you might as well try this out
right? nope.. there's just 2 more steps..

STEP#1 = Reboot the computer.

STEP#2 = Change the path for my Example Project.

ie= I'm sure you've experienced this before. you get a nice VB Snippet of code
which doesn't work.. this often happens because the project was created in a directory
other than the one you'll unzip the files into.

you can test to see if this will happen by opening the project and seeing if
it says "file not found continue loading project?" errors.

I'm sure that you all know what you'll need to do here. :)

secondly.. because I left the examples un-compiled there might be
a few of you who get errors when opening the file in ActiveMovie.

there's a very easy solution to this. :)

all you have to do is go into the filename property for the ActiveMovie1
control and open the file "AiE.mp3" which is a song I'm sure most of you have
heard by now. :)

you can see it in the browse dialog by selecting "ALL FILES" from the drop-down menu
and double clicking on "AiE.mp3".

Remember boys, and girls.. I said this was a hidden exploit for ActiveMovie
and do you really think after all this configuring that MS will provide you
with an "Open an MP3" file option in the open dialog? nope.. like I said thats MS! :)

--------------------------------------------------------------------------
--------------------------------------------------------------------------
PART-IV : CONGRADULATIONS
--------------------------------------------------------------------------
--------------------------------------------------------------------------

You did it! :)
You now know how to implement MP3 files into your VB projects.
you may have gotten confused along the way, but if you paid attention to my directions
you've got yourself a completely cool do-it-yourself solution to using MP3s through VB5+ :)

and the best part is that it didn't cost you 1 red cent.. i know I know.. 
your thanks are very appreciated, and I can imagine your in the mood to thank us
for helping the VB community do this now right? ok.. well the next section is
for those of you who wish to thank us.. nope we don't want donations. :)

just read the next section. :)


--------------------------------------------------------------------------
--------------------------------------------------------------------------
PART-V : AIE'S RANTS ABOUT MAKING THIS AVAILABLE TO YOU. :)
--------------------------------------------------------------------------
--------------------------------------------------------------------------

For the record, we are thankful that you are thakful.

right now AiE is going though a Huge promotional venture, to boost
recognition for both our company, and upcoming game called TALUS.

p.s.- TALUS implements a lot of our technology, and more! :)

anyhow.. all you need to do is tell your friends about us
thats all we want.. also We'd like to get a really great
number of multimedia related threads going on our discussion boards.

you can visit both our website, or WebBoard by clicking on the hyperlinks in
the Example projects, or by cutting and pasting the following addresses seen
below into the location textbox in your web browser. :)

Also note: anytime we do anything, or release projects/code/etc..
you can be assured it will always be announced there first. :)


http://members.tripod.com/~AceInterActive/     <--  WEBSITE

http://disc.server.com/discussion.cgi?id=17429  <-- WEB FORUM


thats all for now, Keep Kewl.. and We'll be releasing more stuff soon so
always check the addresses above for the Latest news about AiE, and it's Projects.

As for me.. I've gotta run.. the Coffee's Ready and I'm feeling tired. :)



-- 
 Regards,

       Rw Bivins, AiE(tm.) Founder.
         http://members.tripod.com/~AceInterActive/
           "We Don't Make Games, We make KickAss Games."
 