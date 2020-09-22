VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9630
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10395
   _ExtentX        =   18336
   _ExtentY        =   16986
   _Version        =   393216
   Description     =   "Mouse Zoom-Camera Add-In"
   DisplayName     =   "MouseCam"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const guidMYTOOL$ = "_M_O_U_S_E__C_A_M_"

Public WithEvents CmpHandler  As VBComponentsEvents        'Komponenten-Ereignisbehandlungsroutine
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents          'Befehlsleisten-Ereignisbehandlungsroutine
Attribute MenuHandler.VB_VarHelpID = -1

Sub Show()
  On Error GoTo ShowErr
  gwinWindow.Visible = True

  Exit Sub
ShowErr:
  MsgBox Err.Description, , "Sub_Show"
  Err.Clear
  
End Sub

Public Property Get NonModalApp() As Boolean
  NonModalApp = True  'Verwendet von der Add-In-Symbolleiste
End Property

'------------------------------------------------------
'Diese Methode fügt das Add-In der VB-Symbolleiste
' hinzu. Sie wird vom VB-AddIn-Manager aufgerufen
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
  On Error GoTo AddinInstance_OnConnectionErr
  
  Dim aiTmp As AddIn
  
  'Speichern der VB-Instanz
  Set gVBInstance = Application

  If Not gwinWindow Is Nothing Then
    'Wird schon ausgeführt, daher nur anzeigen
    Show
    If ConnectMode = ext_cm_AfterStartup Then
      'Gestartet vom Add-In-Manager
      AddToCommandBar
    End If
    Exit Sub
  End If
  
  'Erstellen des Symbolfensters
  If ConnectMode = ext_cm_External Then
    'Überprüfen, ob es schon ausgeführt wird.
    On Error Resume Next
    Set aiTmp = gVBInstance.Addins("MouseCam.Connect")
    On Error GoTo AddinInstance_OnConnectionErr
    If aiTmp Is Nothing Then
      'Anwendung nicht in der VBADDIN.INI-Datei,
      'daher nicht in der Auflistung.
      'Daher Versuch, erstes Add-In in der Auflistung zu verwenden,
      'nur damit die vorliegende Anwendung ausgeführt wird. Falls
      'kein Add-In vorhanden ist, tritt ein Fehler auf, und diese
      'Anwendung wird nicht ausgeführt.
      Set gwinWindow = gVBInstance.Windows.CreateToolWindow(gVBInstance.Addins(1), "MouseCam.docMouseCam", "Mouse Camera", guidMYTOOL$, gdocMouseCam)
    Else
      If aiTmp.Connect = False Then
        Set gwinWindow = gVBInstance.Windows.CreateToolWindow(aiTmp, "MouseCam.docMouseCam", "Mouse Camera", guidMYTOOL$, gdocMouseCam)
      End If
    End If
  Else
    'Muß vom Add-In-Manager aufgerufen worden sein
    Set gwinWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "MouseCam.docMouseCam", "Mouse Camera", guidMYTOOL$, gdocMouseCam)
  End If

  'Zuweisen der Ereignisbehandlungsroutinen für Projekt, Komponenten
  'und Steuerelemente
  Set Me.CmpHandler = gVBInstance.Events.VBComponentsEvents(Nothing)
  
  If ConnectMode = vbext_cm_External Then
    'Gestartet von der Add-In-Symbolleiste aus
    Show
  ElseIf ConnectMode = vbext_cm_AfterStartup Then
    'Gestartet vom Add-In-Manager aus
    AddToCommandBar
  End If

  Exit Sub
  
AddinInstance_OnConnectionErr:
  MsgBox Err.Description, , "AddinInstance_OnConnection"
  Err.Clear
  
End Sub

'------------------------------------------------------
'Dieses Ereignis entfernt das Befehlsleisten-Menü
'Es wird vom VB Add-In-Manager aufgerufen
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
  On Error GoTo IDTExtensibility_OnDisconnectionErr
  'Löschen des Befehlsleisteneintrags
  gVBInstance.CommandBars(2).Controls("Mouse Camera window").Delete

  'Speichern des Formularzustands für den nächsten Aufruf von VB
  If gwinWindow.Visible Then
    SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "1"
  Else
    SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "0"
  End If
  
  Set gwinWindow = Nothing
  
IDTExtensibility_OnDisconnectionErr:
  Err.Clear
  Resume Next
  
End Sub

'Dieses Ereignis wird ausgelöst, wenn die
'IDE (integrierte Entwicklungsumgebung) vollständig geladen ist.
Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
  AddToCommandBar
End Sub

'Dieses Ereignis wird ausgelöst, wenn auf das
'Befehlsleisten-Steuerelement in der IDE geklickt wird.
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Show
End Sub


'Dieses Ereignis wird ausgelöst, wenn ein Formular in der IDE
'aktiviert wird.
Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
On Error GoTo CmpHandler_ItemActivatedErr


Exit Sub
CmpHandler_ItemActivatedErr:

End Sub

'Dieses Ereignis wird ausgelöst, wenn ein Formular im Projektfenster
'ausgewählt wird.
Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
  CmpHandler_ItemActivated VBComponent
End Sub

Sub AddToCommandBar()
  On Error GoTo AddToCommandBarErr
  
  'Sicherstellen, daß Standard-Symbolleiste eingeblendet ist
  gVBInstance.CommandBars(2).Visible = True
  'Hinzufügen zur Befehlsleiste
  'Mit folgender Zeile wird der MouseCam-Manager der Standard-
  'Symbolleiste rechts vom Werkzeugsammlungssymbol hinzugefügt.
  gVBInstance.CommandBars(2).Controls.Add 1, , , gVBInstance.CommandBars(2).Controls.Count
  'Festlegen der Titelleiste
  gVBInstance.CommandBars(2).Controls(gVBInstance.CommandBars(2).Controls.Count - 1).Caption = "Mouse Camera window"
  'Kopieren des Symbols in die Zwischenablage.
  Clipboard.SetData LoadResPicture(1000, 0)
  'Festlegen des Symbols für die Schaltfläche.
  gVBInstance.CommandBars(2).Controls(gVBInstance.CommandBars(2).Controls.Count - 1).PasteFace
  'Zuweisen des Ereignisses.
  Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(gVBInstance.CommandBars(2).Controls(gVBInstance.CommandBars(2).Controls.Count - 1))
  
  'Wiederherstellen des letzten Zustands.
  If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") = "1" Then
    'Dadurch wird das Formular beim Erstellen der Verbindung angezeigt.
    Me.Show
  End If
  
  Exit Sub
    
AddToCommandBarErr:
  MsgBox Err.Description, , "AddToCommandBar"
  Err.Clear
  
End Sub

