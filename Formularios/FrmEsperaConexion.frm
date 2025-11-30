VERSION 5.00
Begin VB.Form FrmEsperaConexion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexión"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Abortar y salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ESPERANDO CONECTAR CON DATA CENTER ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "FrmEsperaConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload MdiPrincipal
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
    CentrarFormulario Me
   
End Sub
