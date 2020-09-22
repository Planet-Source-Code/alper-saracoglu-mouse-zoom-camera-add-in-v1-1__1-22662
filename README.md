<div align="center">

## Mouse\-zoom\-camera Add\-In v1\.1

<img src="PIC20014241149105545.gif">
</div>

### Description

This is the latest version of my mouse-zoom camera add-in. It is especially useful when designing skinned forms. You can zoom in the form so that you can place your controls at the right places. The source code is heavily commented.

UPDATES (since v1.0):

1.The mouse-cam pauses now when any code-window gets focus, and resumes for any other window.

2.The last zoom level is saved in registry, so that add-in starts with the last zoom level.

3.The link to Office DLL is removed. MenuBar icon placement is now without creating a menubar object. The add-in uses therefore much less resources now.

4.MenuBar icon is changed. It looks more professional now.

5.Button tool-tips added.
 
### More Info
 
HOW IT WORKS:

Compile the project to an ActiveX.DLL and drop the ActiveX.DLL in the C:\Program Files\Microsoft Visual Studio\Common\MSDev98\AddIns folder. From Add-Ins Menu, start Add-in Manager, and load MouseCam Add-in. An icon will be added to your toolbar. Clicking this icon will make the MouseCam

tool-window visible. Initially, mousecam will be off. Start it with the on/off button. The initial snapshot interval is 100 ms. You can change this by clicking the interval button. The initial zoom is 100%. You can go up to 9900% by clicking the

zoom-in button. When the zoom is more than 100%, the 1:1 button will be enabled. Clicking this button sets the zoom to 100%, and makes the ZoomBack button enabled. You can go back to the last zoom level, by clicking the ZoomBack button.


<span>             |<span>
---                |---
**Submitted On**   |2001-04-24 17:42:30
**By**             |[Alper Saracoglu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alper-saracoglu.md)
**Level**          |Advanced
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Mouse\-zoom187834242001\.zip](https://github.com/Planet-Source-Code/alper-saracoglu-mouse-zoom-camera-add-in-v1-1__1-22662/archive/master.zip)

### API Declarations

uses StretchBlt, GetDesktopWindows, GetWindowDC, GetCursorPos API's





