VERSION 5.00
Begin VB.Form FrmDatosComponente 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del componente"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDatosComponente.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCmdCancelar 
      Height          =   495
      Left            =   4080
      TabIndex        =   19
      Top             =   2320
      Width           =   1455
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -240
         Picture         =   "FrmDatosComponente.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Frame fraCmdAceptar 
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   2320
      Width           =   1215
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -360
         Picture         =   "FrmDatosComponente.frx":0B85
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.TextBox TxtFactor 
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
      Left            =   3495
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4965
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E8E8E8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   30
      TabIndex        =   6
      Top             =   -75
      Width           =   5415
      Begin VB.ComboBox CmbComponentes 
         Height          =   330
         Left            =   3735
         TabIndex        =   1
         Top             =   315
         Width           =   1560
      End
      Begin VB.TextBox TxtNormArea 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2505
         TabIndex        =   7
         Top             =   1950
         Width           =   720
      End
      Begin VB.TextBox TxtUnits 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Text            =   "PPM"
         Top             =   1410
         Width           =   975
      End
      Begin VB.TextBox TxtExternal 
         Height          =   285
         Left            =   885
         TabIndex        =   4
         Top             =   1410
         Width           =   570
      End
      Begin VB.TextBox TxtArea 
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Top             =   870
         Width           =   750
      End
      Begin VB.TextBox TxtRetention 
         Height          =   285
         Left            =   885
         TabIndex        =   2
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox TxtNumeroComponente 
         Height          =   285
         Left            =   885
         TabIndex        =   0
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NormArea"
         Height          =   210
         Left            =   1725
         TabIndex        =   14
         Top             =   2010
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         Height          =   210
         Left            =   3120
         TabIndex        =   13
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "External"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         Height          =   210
         Left            =   3120
         TabIndex        =   11
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retention"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   930
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   210
         Left            =   3120
         TabIndex        =   9
         Top             =   390
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   555
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Factor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2250
      TabIndex        =   15
      Top             =   4980
      Width           =   450
   End
End
Attribute VB_Name = "FrmDatosComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tabla As String
Public gObjPozoAnalis                       As clsPozoAnalis
Public gObjPozoAnalisComponente             As ClsPozoAnalisComponente

Private pAccion                             As euAccion


Public Property Get Accion() As euAccion
    Accion = pAccion
End Property

Public Property Let Accion(ByVal vNewValue As euAccion)
    pAccion = vNewValue
End Property


Private Sub cmdAceptar_Click()
    
    Dim NumeroAnalisis As Long
    Dim NumeroComponente As Integer
    Dim ok As Boolean
    
    If ValidarDatos Then
        If Me.Accion = euAccionAgregar Then
            gDatos.BeginTrans
            If ComponenteAgregar() Then
                gDatos.CommitTrans
                Me.Hide
            Else
                gDatos.RollBackTrans
            End If
            
        ElseIf Me.Accion = euAccionModificar Then
            gDatos.BeginTrans
            If ComponenteModificar Then
                CalcularRelacionesCromatograficas gObjPozoAnalis, "MODIFICAR", ok
                ActualizarGasTotalCromatografico gObjPozoAnalis, ok
                If ok Then
                    gDatos.CommitTrans
                    Me.Hide
                Else
                    gDatos.RollBackTrans
                End If
            Else
                gDatos.RollBackTrans
            End If
        End If
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    
    Unload Me
    
End Sub

Private Function ValidarDatos() As Boolean
    
    On Error GoTo Error
    
    ValidarDatos = True
    
    If Trim(TxtRetention.Text) = "" Then
    
        TxtRetention.Text = "0"

'        ValidarDatos = False
'
'        frmMsg.MostrarMsg "Los datos del componente están incompletos.", "Error"
'
'        TxtRetention.SetFocus
'
'        Exit Function

    End If
'
    If Trim(TxtArea.Text) = "" Then
    
        TxtArea.Text = "0"

'        ValidarDatos = False
'
'        frmMsg.MostrarMsg "Los datos del componente están incompletos.", "Error"
'
'        TxtArea.SetFocus
'
'        Exit Function

    End If
    
    If Trim(TxtExternal.Text) = "" Then
    
        TxtExternal.Text = "0"
                
'        ValidarDatos = False
'
'        frmMsg.MostrarMsg "Los datos del componente están incompletos.", "Error"
'
'        TxtExternal.SetFocus
'
        Exit Function
        
    End If
    

    
Error:
    
    If Err.Number <> 0 Then
        
        ValidarDatos = False
        
        frmMsg.MostrarMsg "Módulo: ValidarDatos." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Sub Form_Activate()
    
    SetearFormularioSegunAccion

End Sub

Private Sub SetearFormularioSegunAccion()
    
    If Accion = euAccionAgregar Then
        Me.caption = "Agregando Componente"
        Set gObjPozoAnalisComponente = New ClsPozoAnalisComponente
        
    ElseIf Accion = euAccionModificar Then
        Me.caption = "Modificando Componente"
        PozoAnalisComponenteDatosMostrar
    End If

End Sub

Private Sub Form_Load()
    ComponentesBuscar
End Sub

Private Function PozoAnalisComponenteDatosMostrar() As Boolean
    
    On Error GoTo Error
    
    PozoAnalisComponenteDatosMostrar = True
    
    PosicionarSimpleCombo CmbComponentes, gObjPozoAnalisComponente.NumeroComponente
    
    If gObjPozoAnalisComponente.Component = "" Then
        CmbComponentes.Enabled = True
    Else
        CmbComponentes.Enabled = False
    End If
    
    TxtRetention.Text = gObjPozoAnalisComponente.Retention
    TxtArea.Text = gObjPozoAnalisComponente.Area
    TxtExternal.Text = gObjPozoAnalisComponente.Externo
    TxtUnits.Text = gObjPozoAnalisComponente.Units
    TxtNormArea.Text = gObjPozoAnalisComponente.NormArea
    
    If gObjPozoAnalisComponente.Externo <> 0 Then
        TxtFactor = gObjPozoAnalisComponente.Area / gObjPozoAnalisComponente.Externo
        Else: TxtFactor = 0
    End If
    
Error:
    
    If Err.Number <> 0 Then
        PozoAnalisComponenteDatosMostrar = False
        frmMsg.MostrarMsg "Módulo: PozoAnalisComponenteDatosMostrar." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function ComponenteModificar() As Boolean
    
    gObjPozoAnalisComponente.Retention = TxtRetention.Text
    gObjPozoAnalisComponente.Area = TxtArea.Text
    gObjPozoAnalisComponente.Externo = TxtExternal.Text
    gObjPozoAnalisComponente.Units = TxtUnits.Text
    gObjPozoAnalisComponente.NormArea = TxtNormArea.Text
    
    Set gObjPozoAnalis.PozoAnalisComponentes.Datos = gDatos
    
    ComponenteModificar = gObjPozoAnalis.PozoAnalisComponentes.dbModificar(gObjPozoAnalisComponente)
    
End Function
Private Sub TxtArea_GotFocus()
    TextSelected
End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 46 Then
        
        If InStr(1, TxtArea.Text, ".") <> 0 Then
            
            KeyAscii = 0
            
        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtArea_LostFocus()
    
    Dim Factor As Double
    Dim Area As Double
    
    If Me.Accion = euAccionModificar Then
        
        If TxtArea.Text <> "" Then
            
            Area = CDbl(TxtArea.Text)
            Factor = CDbl(TxtFactor.Text)
            
            'TxtExternal.Text = Format(Area / Factor, "0")
            
        End If
        
    End If
    
End Sub

Private Sub TxtExternal_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtExternal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtExternal_LostFocus()
    
    Dim Factor As Double
    Dim Externo As Long
    
    If Me.Accion = euAccionModificar Then
        
        If TxtExternal.Text <> "" Then
            
            Externo = CLng(TxtExternal.Text)
            Factor = CDbl(TxtFactor.Text)
            
            TxtArea.Text = Format(Factor * Externo, "0.0")
            
        End If
        
    End If
    
End Sub

Private Sub TxtNormArea_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtNormArea_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 46 Then
        
        If InStr(1, TxtNormArea.Text, ".") <> 0 Then
            
            KeyAscii = 0
            
        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtNumeroComponente_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtNumeroComponente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtNumeroComponente_Validate(Cancel As Boolean)
    
    Dim NumeroAnalisis As Long
    Dim NumeroComponente As Integer
    
    If TxtNumeroComponente.Text <> "" Then
        NumeroComponente = TxtNumeroComponente.Text
        If Me.Accion = euAccionAgregar Then
            If ExisteComponenteEnAnalisis(NumeroComponente) Then
                frmMsg.MostrarMsg "El componente que desea ingresar ya existe.", "Error", MdiPrincipal
                Cancel = True
            Else
                If ExisteComponente(NumeroComponente) Then
                    PosicionarSimpleCombo CmbComponentes, NumeroComponente
                Else
                    frmMsg.MostrarMsg "El número de componente que desea ingresar no es válido.", "Error", MdiPrincipal
                    Cancel = True
                End If
                
            End If
            
        End If
        
    End If
    
End Sub

Private Sub TxtRetention_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtRetention_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 46 Then
        
        If InStr(1, TxtRetention.Text, ".") <> 0 Then
            
            KeyAscii = 0
            
        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtUnits_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtUnits_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 39 Then
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Function ComponentesBuscar() As Boolean
        
    Dim objComponente As clsComponente
    
    Dim StrSql As String
    Dim Componentes As Recordset
    
    On Error GoTo Error
    
    ComponentesBuscar = True
    For Each objComponente In gObjComponentes
        CmbComponentes.AddItem objComponente.NombreComponente
        CmbComponentes.ItemData(CmbComponentes.NewIndex) = objComponente.CodigoComponente
    Next
    
Error:
    
    If Err.Number <> 0 Then
        ComponentesBuscar = False
        frmMsg.MostrarMsg "Módulo: ComponentesBuscar." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function ComponenteAgregar() As Boolean
    
    On Error GoTo Error
    
    ComponenteAgregar = False
    
    gObjPozoAnalisComponente.Area = TxtArea.Text
    
    gObjPozoAnalisComponente.NumeroAnalisis = gObjPozoAnalis.NumeroAnalisis
    gObjPozoAnalisComponente.NumeroComponente = CmbComponentes.ItemData(CmbComponentes.ListIndex)
    gObjPozoAnalisComponente.Component = CmbComponentes.Text
    gObjPozoAnalisComponente.Retention = TxtRetention.Text
    gObjPozoAnalisComponente.Externo = TxtExternal.Text
    gObjPozoAnalisComponente.Units = TxtUnits.Text
    gObjPozoAnalisComponente.IdPozo = gObjPozoAnalis.IdPozo
    
    Set gObjPozoAnalis.PozoAnalisComponentes.Datos = gDatos
    
    If gObjPozoAnalis.PozoAnalisComponentes.dbAgregar(gObjPozoAnalisComponente) Then
        ComponenteAgregar = True
        gObjPozoAnalis.PozoAnalisComponentes.colAgregar gObjPozoAnalisComponente
    End If
    
Error:
    If Err.Number <> 0 Then
        ComponenteAgregar = False
        frmMsg.MostrarMsg "Módulo: ComponenteAgregar." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error"
        Err.Clear
    End If
    
End Function

Private Function ExisteComponenteEnAnalisis(NumeroComponente As Integer) As Boolean
   
    Dim StrSql                                  As String
    Dim rDatos                                  As Recordset
    Dim objPozoAnalisComponente                 As ClsPozoAnalisComponente
    
    On Error GoTo Error
    
    ExisteComponenteEnAnalisis = Not gObjPozoAnalis.PozoAnalisComponentes.colComponente(gObjPozoAnalis.IdPozo, gObjPozoAnalis.NumeroAnalisis, NumeroComponente) Is Nothing
    
Error:
    
    If Err.Number <> 0 Then
        
        ExisteComponenteEnAnalisis = False
        
        frmMsg.MostrarMsg "Módulo: ExisteComponenteEnAnalisis. " & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function ExisteComponente(NumeroComponente As Integer) As Boolean
   
    Dim StrSql As String
    Dim rDatos As Recordset
    
    On Error GoTo Error
    
    ExisteComponente = Not gObjComponentes.colComponente(NumeroComponente) Is Nothing
    
Error:
    
    If Err.Number <> 0 Then
        ExisteComponente = False
        frmMsg.MostrarMsg "Módulo: ExisteComponente. " & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function
