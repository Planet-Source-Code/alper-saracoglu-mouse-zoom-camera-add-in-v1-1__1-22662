AUTHOR:		Alper Saracoglu
DATE:		April 23, 2001
NAME:		Mouse Zoom Camera Add-In
TYPE:		Visual Basic Add-In for VB (full commented source code)
RELEASE:	Version 1.1

OVERVIEW:
This is the mouse-zoom camera add-in that I use when I am designing busy
forms. I wanted to share it with you. The code is fully commented.

UPDATES (since v1.0):
1.The mouse-cam pauses now when any code-window gets focus, and resumes for any other window.
2.The last zoom level is saved in registry, so that add-in starts with the last zoom level.
3.The link to Office DLL is removed. MenuBar icon placement is now without creating a
  menubar object. The add-in uses therefore much less resources now. 
4.MenuBar icon is changed. It looks more professional now.
5.Button tool-tips added.

INCLUDED:
There is a userdocument called docMouseCam that is the visible and dockable
part of the add-in. It consists of a PictureBox, for bitblt'ing, and some
command buttons for controlling the zoom level, snapshot interval etc.
The global declarations are on modMain, and the CommandBar bitmap and
DLL icon is on the resource file. The Connect Designer module, named 
MouseCam.Dsr is from the Addin template of VB6. I have not changed this file
(Thus, the comments are in German, from my German VB6. 
The English version of VB has AddIn Template in English. You 
can start a new Add-in Project to read the comments.)

HOW IT WORKS:
Compile the project to an ActiveX.DLL, (or just use the compiled DLL) and drop
the ActiveX.DLL in the C:\Program Files\Microsoft Visual Studio\Common\MSDev98\AddIns 
folder. From Add-Ins Menu, start Add-in Manager, and load MouseCam Add-in.
An icon will be added to your toolbar. Clicking this icon will make the MouseCam
tool-window visible. Initially, mousecam will be off. Start it with the on/off button.
The initial snapshot interval is 100 ms. You can change this by clicking the 
interval button. The initial zoom is 100%. You can go up to 9900% by clicking the
zoom-in button. When the zoom is more than 100%, the 1:1 button will be enabled.
Clicking this button sets the zoom to 100%, and makes the ZoomBack button enabled.
You can go back to the last zoom level, by clicking the ZoomBack button.

KNOWN BUGS:
Code appears to be rock-solid. I have been using this Add-in for a long time. 
If you notice anything non-functional, please contact me.

PLANNED UPDATES:
1.I am trying to change the caption of the add-in in run-time, so that the caption shows
  the current zoom level and status of camera. The userdocument object does not have a hwnd,
  and the caption of the toolbox itself is read only. The setwindowtext api does not work.
2.I do not know if it would ever be needed, but I also plan to implement zoom levels below
  100% (ie. 50, 25, 10 etc), and a hotkey to dump the display in a given directory as .bmp

DISCLAIMER:
1. BECAUSE THE PROGRAM IS LICENSED FREE OF CHARGE, THERE IS NO WARRANTY FOR THE PROGRAM, 
TO THE EXTENT PERMITTED BY APPLICABLE LAW. EXCEPT WHEN OTHERWISE STATED IN WRITING THE 
COPYRIGHT HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY OF 
ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK 
AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU. SHOULD THE PROGRAM PROVE 
DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING, REPAIR OR CORRECTION. 
2. IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING WILL ANY 
COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MAY MODIFY AND/OR REDISTRIBUTE THE PROGRAM 
AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY GENERAL, SPECIAL, 
INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE USE OR INABILITY TO USE THE 
PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF DATA OR DATA BEING RENDERED INACCURATE 
OR LOSSES SUSTAINED BY YOU OR THIRD PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE 
WITH ANY OTHER PROGRAMS), EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE 
POSSIBILITY OF SUCH DAMAGES. 


LICENSE:
You are not given any restrictions for the usage of this code except
for the inclusion of it in commercial projects, where permission from
me would be required (or atleast the inclusion be brought to my notice).
If you appreciate the work put into the creation of this Add-in,
please include my name and if possible a link to my homepage or email
address in the credits section. This DLL may be used in all kinds of 
projects which are intended for FREE NON-PROFIT distribution.
A phrase such as 'Uses code from LoGo Systemhaus www.logo-systemhaus.de'
is also sufficent.

yours sincerely,
Alper Saracoglu
saracoglu@estetiksoft.de
saracoglu@logo-systemhaus.de
ICQ UIN: 4099829
