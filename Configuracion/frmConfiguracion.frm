VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IGA - Configuración del sistema"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   45
      TabIndex        =   4
      Top             =   -30
      Width           =   7110
      Begin VB.TextBox txtCaminoArchivos 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   4560
      End
      Begin VB.CommandButton cmdSeleccionarCaminoArchivos 
         Caption         =   "..."
         Height          =   255
         Left            =   6600
         TabIndex        =   1
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivos"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   315
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   450
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   4830
      TabIndex        =   2
      Top             =   765
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6000
      TabIndex        =   3
      Top             =   765
      Width           =   1155
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const gNombreBaseDeDatos = "Iga.mdb"
Dim gNombreBaseDeDatosDGC As String

Private LoadOk As Boolean

Private Sub cmdAceptar_Click()
    
    GuardarConfiguracion
    
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSeleccionarCaminoArchivos_Click()
    
    Dim ArchivoOk As Boolean
    
    On Error GoTo Error
    
    CommonDialog1.DialogTitle = "Camino de los Archivos"
    CommonDialog1.Flags = cdlOFNExplorer Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    'CommonDialog1.DefaultExt = "*.log"
    'CommonDialog1.Filter = "(*.log)"
    
    Do
        
        CommonDialog1.ShowOpen
        
        ArchivoOk = UCase(Right(CommonDialog1.FileTitle, 3)) = "LOG" Or UCase(Right(CommonDialog1.FileTitle, 3)) = "TXT" Or UCase(Right(CommonDialog1.FileTitle, 3)) = "ESD"
        
        If Not ArchivoOk Then
            
            MsgBox "El archivo seleccionado no es valido.", vbOKOnly + vbInformation
            
        End If
        
    Loop Until ArchivoOk
    
    If ArchivoOk Then
        
        txtCaminoArchivos.Text = Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
        
    End If
    
Error:
    
    If Err.Number <> 0 And Err.Number <> 32755 Then
        
        MsgBox "Módulo: cmdSeleccionarCaminoArchivos_Click." & Chr(10) & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical
        
        Err.Clear
        
    End If
    
End Sub

Private Sub Form_Activate()
    
    If Not LoadOk Then
        
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Load()
    
    LoadOk = False
    
    txtCaminoArchivos.Text = GetSetting("Iga", "Caminos", "Archivos")
    LoadOk = True
    
End Sub

Private Function GuardarConfiguracion() As Boolean
    
    On Error GoTo Error
    
    GuardarConfiguracion = True
    
    SaveSetting "Iga", "Caminos", "NombreBaseDeDatosDGC", gNombreBaseDeDatosDGC
    SaveSetting "Iga", "Caminos", "Archivos", txtCaminoArchivos.Text
    
Error:
    
    If Err.Number <> 0 Then
        
        GuardarConfiguracion = False
        
        MsgBox "Ocurrió el error: " & Err.Number & ": " & Err.Description, vbOKOnly + vbInformation
        
        Err.Clear
        
    End If
    
End Function
