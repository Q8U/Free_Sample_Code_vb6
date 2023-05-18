SIMPLE - SImple Mpeg audio PLayer and Encoder (Release 1.3)
http://www.geocities.com/simplelance
simplelance@hotmail.com

Files:
    smp3p.exe - Simple MP3 Player version 0.8c
    smp3m.exe - Simple MP3 Maker version 1.0g (Encoder version 1.1)

sMP3p (player) Keyboard Shortcuts:
    Previous       Z
    Play           X
    Pause          C
    Stop           V
    Next           B
    Play List      L
    Shuffle        S
    Volume Up      +
    Volume Down    -

To create your own player display designs:

    1. Create an image named main.bmp. This should be the player
       display. This file is composed of three equally sized bitmaps.
       The first one is the main image. The second one is the "mask"
       which defines which portions of the main image will be displayed
       (to create a floating window). Black pixels will be displayed
       and white pixels will not. The third one defines the image for
       buttons when selected.

    2. Create an image named eq.bmp. This should be the equalizer
       display. This file is composed of three equally sized bitmaps.
       The first one is the main image. The second one is the "mask"
       which defines which portions of the main image will be displayed
       (to create a floating  window). Black pixels will be displayed
       and white pixels will not. The third one defines the image for
       buttons when selected and the image of the slider buttons.

    3. Create a file named display.ini to specify the portions of the
       display that act as buttons and text displays.

       The buttons have the format:

           btn_??? = x y w h
               where x and y are the left and top co-ordinates in the
               bitmap for the button, and w and h are the width and
               height

           example:
               btn_stop 4 10 15 12
               - the stop button is some rectangle at 4, 10 with width
                 15 and height 12

       The text display has the format:

           disp_??? x y w h font_size font_color font_name
               where you have the usual co-ordinates and width and
               height followed by the size of the font for display, the
               color of the font (in hex) and the name of the font.

           example:
               disp_file 10 30 90 15 12 00ffffff Arial
               - the name of the file being played will have the font
                 size 12, white color and Arial type

           Note that the color is in the RGB format, so, pure red is
           00ff0000, pure green is 0000ff00 and pure blue is 000000ff.

       The equalizer tracks has the format:

           eq_interval = n
               where n is the interval between equalizer tracks

           eq_track = x y w h
               defines the coordinates and size of the track image

           eq_slider = x y w h
               defines the coordinates and size of the slider button

sMP3m (MP3 Maker) command-line options
    WAV filename
    -mono (convert to mono)
    -js (joint-stereo)
    -NNN (bit-rate, e.g., -128 , -192)
    -hsq (highest quality VBR)
    -hq (high quality VBR)
    -nq (normal quality VBR)
    -lq (low quality VBR)
    -lrq (lower quality VBR)
    -lsq (lowest quality VBR)
    -delete (delete input WAV file after successful encoding)
    Examples: smp3m track01.wav -js -160
              smp3m track04.wav track05.wav -hsq -delete

