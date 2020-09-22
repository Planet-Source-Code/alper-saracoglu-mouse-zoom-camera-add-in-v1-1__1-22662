Attribute VB_Name = "modMain"
Option Explicit

Global gVBInstance  As VBIDE.VBE       'Instanz der VB-IDE
Global gwinWindow   As VBIDE.Window    'Zum Absichern, daß nur eine
                                       'Instanz ausgeführt wird
Global gdocMouseCam As docMouseCam     'userdokument-Object

Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"

'Enum to control Camera States:
Enum mcCameraState
    mcStop = 0
    mcRun = 1
    mcPause = 2
End Enum

Global CameraState As mcCameraState 'current State of Camera

Function InRunMode(VBInst As VBIDE.VBE) As Boolean
  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function

