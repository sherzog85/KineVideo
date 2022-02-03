# KineVideo
## What is this project?
A long time ago (in 2010/2011) I was a trainee teacher for physics and math. In those days I held a teaching series how to analyse a flight curve
(i.e. the ball during a penalty shot in basketball) with the help of video material. 
In this program one can mark the position of the flying ball frame by frame and export the result in a file (tab-separated text file or Excel-XML-format).
The following picture shows an example with a locomotive (okay,okay it is not a ball ;-) ):
<img src="https://github.com/sherzog85/KineVideo/blob/main/media/kinevideo_example.png"/>

The program was written in C++ based on the Windows API and compiled with [Borland Compiler 5.5](https://www.kompf.de/cplus/artikel/bcc32.html). For playback of the video (in AVI-format) the 
media control interface ([MCI](https://docs.microsoft.com/en-us/windows/win32/multimedia/mci)) was used.

## History
In those days I had contact via my supervisor with a school book publisher, who considered to publish this piece of the software along with a new physics teaching book.
However, I didn't work out. I think they lost interest. I always had in mind to share this software, but was  too lazy to do it. But now, I'm taking the time for it.


## Status
In Windows XP the program ran smoothly with all different kind of videos. However, in Windows 10 I succeed in loading uncompressed video files only.
With a VfW (Video-for-Windows) codec an encoded video file (i.e. with H264-codec) is loaded, but the picture is not shown properly.
Since a video showing the flying curve of a ball shouldn't be much longer than some seconds, I don't think that this a drawback.
Up to now only an implementation in German exists and the manual and instruction video (in subfolder manual) are also in German only.


So have fun with this piece of software.

