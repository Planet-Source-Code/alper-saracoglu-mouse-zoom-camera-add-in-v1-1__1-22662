VERSION 5.00
Begin VB.Form frmMouseCam 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Mouse Cam"
   ClientHeight    =   1530
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   2040
   Icon            =   "frmMouseCam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   945
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   1845
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1035
      TabIndex        =   1
      Top             =   1140
      Width           =   900
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   1140
      Width           =   900
   End
End
Attribute VB_Name = "frmMouseCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub OKButton_Click()
    MsgBox "AddIn operation on: " & VBInstance.FullName
End Sub
