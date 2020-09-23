Javamorse is a project based on the java morse coder on the web from
Steven Phillips (http://morsecode.scphillips.com/). His sourcecode is available in a 
modified version at http://www.omnicron.com/~ford/java/NMorse.html.

I know that there is already a whole bunch of morse code apps out there, but 
none of them were exactly what I wanted.

For Javamorse I have used directsound, for which I am grateful to Petr Supina,
who gave me the idea with his Oscillator project. (see his link here http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39362&lngWId=1).

Javamorse converts text into morse strings, which are used to generate a soundstream
in memory, which can then be played using the direct sound engine (requires DX7). 
Speed can be selected between 4 and 40 wpm. The speed is only a rough indication, if
you measure the actual speed at the various settings you may notice that it is off
a bit here and there.

The trainer also has a random character generator which you can set to generate letters,
numbers, or a combination of both.

When playing the morse strings, the text is highlighted as playing progresses, this is handy
for novice morse operators.

Javamorse has only been tested in win xp, I have no idea whether early windows versions 
support this.

To compile the project, a reference to dx7 is required.

No need to vote. Please report however any bugs that you may encounter. I have tried to debug
as much as possible, but I'm sure that there are still bugs around.
Where I thought this was handy, I've commented the code. It is not awfully complicated so even
without commenting too heavily this code is easy enough to follow I think.

Have fun!

