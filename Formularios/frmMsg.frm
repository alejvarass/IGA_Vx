VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5370
   Begin VB.Frame fraCmdAceptar 
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -360
         Picture         =   "frmMsg.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5355
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Public Sub MostrarMsg(msg As String, caption As String, Optional ByVal Formulario As Variant)
If FormularioCargado("MdiPrincipal") Then
    If Not FormularioCargado("frmMsg") Then

    lblMsg.caption = msg
    frmMsg.caption = caption
    
    'si el texto tiene menos de 30 caracteres seteo el ancho como el
    'len del texto, por una cuestión de estética
    If (Len(msg) < 30) Then
        cmdAceptar.Width = 855
        lblMsg.Width = Len(msg) + cmdAceptar.Width + 800
        frmMsg.Width = lblMsg.Width
        fraCmdAceptar.Left = frmMsg.Width - cmdAceptar.Width - 100
    Else
        'defino la altura del lbl, según el largo del texto
        lblMsg.Height = ((Len(msg) / 33)) * 210
        
        'ubico el botón al final del lbl, dependiendo de su alto
        fraCmdAceptar.Top = lblMsg.Height + 100
        
        'seteo el alto del form
        frmMsg.Height = cmdAceptar.Top + cmdAceptar.Height + 1200
    End If
    If Not IsMissing(Formulario) Then
        Me.Show , Formulario
        CentrarFormulario Me
    Else
        Me.Show
        CentrarFormulario Me
    End If
    
    cmdAceptar.SetFocus
    
    End If
    Else
        MsgBox msg, vbOKOnly + vbInformation, caption
        Unload Me
    End If
    
End Sub
