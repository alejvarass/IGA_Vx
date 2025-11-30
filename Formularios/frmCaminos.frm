VERSION 5.00
Begin VB.Form frmCaminos 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rutas"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmCaminos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6705
   Begin VB.Frame fraCmdSalir 
      Height          =   615
      Left            =   45
      TabIndex        =   1
      Top             =   510
      Width           =   1110
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00E8E8E8&
         Height          =   600
         Left            =   0
         Picture         =   "frmCaminos.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1110
      End
   End
   Begin VB.Label lblArchivos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gcaminosarchivos"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   1275
   End
End
Attribute VB_Name = "frmCaminos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim max

lblArchivos.caption = "Archivo de resultados: " & gCaminoArchivos


End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmCaminos = Nothing
End Sub
