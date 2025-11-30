VERSION 5.00
Begin VB.Form frmSpanCrono 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setear Span crono"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3195
   Icon            =   "frmSpanCrono.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCmdAceptar 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E8E8E8&
         Default         =   -1  'True
         Height          =   810
         Left            =   -360
         Picture         =   "frmSpanCrono.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.TextBox txtSpanCrono 
      Height          =   285
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame fraCmdCancelar 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1455
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00E8E8E8&
         Cancel          =   -1  'True
         Height          =   810
         Left            =   -240
         Picture         =   "frmSpanCrono.frx":0B36
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Label lblSpanCrono 
      BackStyle       =   0  'Transparent
      Caption         =   "Span crono:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmSpanCrono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    gSpanYCrono = CLng(txtSpanCrono.Text)
    FrmAnalisis.iPlotMasterLog.YAxis(0).Span = gSpanYCrono
    SaveSetting "Iga", "Config", "SpanCrono", gSpanYCrono
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    txtSpanCrono.Text = gSpanYCrono
End Sub

Private Sub txtSpanCrono_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub
