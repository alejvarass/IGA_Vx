VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Pantalla de inicio.frx":0000
   ScaleHeight     =   2119.676
   ScaleMode       =   0  'User
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   2610
      Top             =   2955
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    CentrarFormulario Me
    Timer.Enabled = True
End Sub



Private Sub Timer_Timer()
  Unload Me
End Sub
