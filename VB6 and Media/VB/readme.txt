EQPro 1.5 build 110
-------------------

This is what's new in this build:

	- Enhanced compatibility for the SBLive! sound card.

EQPro 1.5 build 106
-------------------

This is what's new in this build:

	- Finally, the documentation is now available in HTML format.
	- The FindLines function has been enhaced with a Hold parameter.
		Read the documentation for more information on how this new parameter
		works.
	- Support for the Microsoft Digital Sound System 80 speakers and similar
		devices, has been added.
	- Fixed a bug which was making EQPro loose the first line on every available 				component.
	- The source code of the Tester application has been revised and updated accordingly 		to the new features of this build.

	IMPORTANT:
	This new build is not binary compatible with the previous ones, so you should re-	compile your application before redistributing them with this new build.
	For Visual basic users, you should enable the Upgrade ActiveX Components check box
	in roder to let Visual Basic automaticly update your project files with this new 	binary.

EQPro 1.5 build 103
-------------------

This is what's new in this build:

	- Recompiled version of the binary OCX using VB6 Service Pack 3
		It looks like Microsoft has finally released a Service Pack
		that actually works.
		You'll probably notice that EQPro runs a little bit faster!
		And I must say that is MUCH more stable.
		You'll need to have previously installed SP3 in order to correctly
		register this new release.
	- Added a new property: Panning
		You can use this property to read/write the panning value of the
		selected line.
		The property  will return a value from -100 to 100. Where -100 is
		when the panning is full to the left, 100 when is full to the right
		and 0 when is at the default center position (no panning).
		As always, check the tester's application source code to know
		how it works.


Notes on updating your current projects:

	If you get an error when opening a project using a previous
	version of EQPro, do the following:
		- Open the project and answer Yes to all the prompts.
		- Without opening any file (or form) from the project
		  include the new EQPro into the project using CTRL+T
		- Save the project
		- Re-open it and it should work fine...


Distributing your projects using EQPro

	This setup includes both the Dependency information plus the license file
	which enables you to use the EQPro control in your projects.
	Both files will be used internally by the Visual Basic editor and the Compiler.
	You should NEVER distribute those files with your project.


Well, that's all for now...
If you need some further information, please let me know!

Xavier Flix
xavier@xfx.net
http://software.xfx
