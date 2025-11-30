VERSION 5.00
Begin VB.Form frmConfigurarZonaAdic 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definir Zona Adicional"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigurarZonaAdic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCmdCancelar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00E8E8E8&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -240
         Picture         =   "frmConfigurarZonaAdic.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame fraCmdAceptar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E8E8E8&
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -360
         Picture         =   "frmConfigurarZonaAdic.frx":0B85
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zona Adicional:                        min."
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   270
      Width           =   2535
   End
End
Attribute VB_Name = "frmConfigurarZonaAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    MdiPrincipal.SetearMenuZonaMuerta CLng("0" & Text1)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = gZonaMuerta
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

