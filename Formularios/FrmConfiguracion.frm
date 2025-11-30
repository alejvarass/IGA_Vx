VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConfiguracion 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5805
   Begin VB.Frame fraCmdAceptar 
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      Top             =   4275
      Width           =   1215
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -360
         Picture         =   "FrmConfiguracion.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Frame fraCmdCancelar 
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   4275
      Width           =   1455
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -240
         Picture         =   "FrmConfiguracion.frx":0B36
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   -120
         Width           =   1935
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4170
      Left            =   75
      TabIndex        =   2
      Top             =   45
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7355
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15263976
      TabCaption(0)   =   "Factores"
      TabPicture(0)   =   "FrmConfiguracion.frx":0F51
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LvwComponentes"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Variables generales"
      TabPicture(1)   =   "FrmConfiguracion.frx":0F6D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TxtArchivoDeResultados"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtArchivoDeResultadosGenerales"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ChkAnalisisAutomaticos"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "CheckCambioEscalas"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "CheckTomaProfundidad"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "CheckDisparoCiclico"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TxtNivelDisparoMetroMetro"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "TxtUno"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "TxtDiez"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TxtCien"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtIpRabit"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmbTipoEquipoGas"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmbTipoEquipoCroma"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtIpDisparoCroma"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      Begin VB.TextBox txtIpDisparoCroma 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2835
         TabIndex        =   27
         Top             =   2850
         Width           =   2730
      End
      Begin MSComctlLib.ImageCombo cmbTipoEquipoCroma 
         Height          =   345
         Left            =   2835
         TabIndex        =   26
         Top             =   2130
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo cmbTipoEquipoGas 
         Height          =   345
         Left            =   2835
         TabIndex        =   24
         Top             =   1725
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtIpRabit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2835
         TabIndex        =   21
         Top             =   2505
         Width           =   2730
      End
      Begin VB.TextBox TxtCien 
         DataField       =   "Cien"
         DataSource      =   "AdodcConfiguracion"
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
         Left            =   4320
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   5010
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtDiez 
         DataField       =   "Diez"
         DataSource      =   "AdodcConfiguracion"
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
         Left            =   2730
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   5010
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtUno 
         DataField       =   "Uno"
         DataSource      =   "AdodcConfiguracion"
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
         Left            =   1350
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   5010
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtNivelDisparoMetroMetro 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   3720
         Width           =   765
      End
      Begin VB.CheckBox CheckDisparoCiclico 
         Caption         =   "Disparo cíclico"
         Height          =   405
         Left            =   2850
         TabIndex        =   8
         Top             =   1545
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   3375
         Width           =   765
      End
      Begin VB.CheckBox CheckTomaProfundidad 
         Caption         =   "Tomar profundidad DGC"
         Height          =   210
         Left            =   2880
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox CheckCambioEscalas 
         Caption         =   "Cambio de escalas Automáticos"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox ChkAnalisisAutomaticos 
         Caption         =   "Disparo de análisis automáticos"
         Height          =   210
         Left            =   90
         TabIndex        =   5
         Top             =   1335
         Width           =   2655
      End
      Begin VB.TextBox TxtArchivoDeResultadosGenerales 
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Top             =   750
         Width           =   2730
      End
      Begin VB.TextBox TxtArchivoDeResultados 
         Height          =   315
         Left            =   2850
         TabIndex        =   3
         Top             =   375
         Width           =   2730
      End
      Begin MSComctlLib.ListView LvwComponentes 
         Height          =   3360
         Left            =   -74775
         TabIndex        =   0
         Top             =   585
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   5927
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "IP disparo croma"
         Height          =   210
         Left            =   75
         TabIndex        =   28
         Top             =   2895
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Croma desde"
         Height          =   210
         Left            =   75
         TabIndex        =   25
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Gas total desde"
         Height          =   210
         Left            =   75
         TabIndex        =   23
         Top             =   1785
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "IP de Rabbit"
         Height          =   210
         Left            =   75
         TabIndex        =   22
         Top             =   2550
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Nivel disparo croma metro a metro (ppm)"
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   3780
         Width           =   3045
      End
      Begin VB.Label Label3 
         Caption         =   "Duración estimada del cromatograma (seg)"
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   3435
         Width           =   3195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de resultados generales"
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   795
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de resultados"
         Height          =   210
         Left            =   90
         TabIndex        =   1
         Top             =   495
         Width           =   1605
      End
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    
    If ValidarDatos Then
        gDatos.BeginTrans
        If GuardarFactores Then
            If GuardarVariablesGenerales Then
                gDatos.CommitTrans
                FrmAnalisis.configVistaEscalas
                MdiPrincipal.IniGralComm
                Unload Me
            Else
                gDatos.RollBackTrans
            End If
        Else
            gDatos.RollBackTrans
        End If
    End If

End Sub

Private Function ValidarDatos() As Boolean

    On Error GoTo Error
    ValidarDatos = True
    
    If cmbTipoEquipoCroma.SelectedItem Is Nothing Then
        ValidarDatos = False
        frmMsg.MostrarMsg "Debe seleccionar el tipo de equipo de Cromatografía", "Error", MdiPrincipal
        cmbTipoEquipoCroma.SetFocus
        Exit Function
    ElseIf cmbTipoEquipoCroma.SelectedItem.Key = "2ID" Or cmbTipoEquipoCroma.SelectedItem.Key = "3ID" Then
        If txtIpRabit.Text = "" Then
            ValidarDatos = False
            frmMsg.MostrarMsg "Debe ingresar la dirección IP de Rabbit", "Error", MdiPrincipal
            txtIpRabit.SetFocus
            Exit Function
        End If
    End If
    
    If cmbTipoEquipoGas.SelectedItem Is Nothing Then
        ValidarDatos = False
        frmMsg.MostrarMsg "Debe seleccionar el tipo de equipo de Gas", "Error", MdiPrincipal
        cmbTipoEquipoGas.SetFocus
        Exit Function
    End If
    
Error:
    If Err.Number <> 0 Then
        ValidarDatos = False
        frmMsg.MostrarMsg "Módulo:ValidarDatos." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If

End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentrarFormulario Me
    SetearList
    ConsultarTiposEquipoGas
    ConsultarTiposEquipoCroma
    BuscarComponentes
    BuscarVariablesGenerales
End Sub

Private Sub ConsultarTiposEquipoGas()
    
    Dim objTipoEquipoGas            As clsTipoEquipoGas
    
    
    For Each objTipoEquipoGas In gObjTiposEquipoGas
        cmbTipoEquipoGas.ComboItems.Add , objTipoEquipoGas.Key, objTipoEquipoGas.TipoEquipoGas
    Next
    
End Sub


Private Sub ConsultarTiposEquipoCroma()
    
    Dim objTipoEquipoCroma            As clsTipoEquipoCroma
    
    
    For Each objTipoEquipoCroma In gObjTiposEquipoCroma
        cmbTipoEquipoCroma.ComboItems.Add , objTipoEquipoCroma.Key, objTipoEquipoCroma.TipoEquipoCroma
    Next
    
End Sub

Private Sub SetearList()
    
    With LvwComponentes.ColumnHeaders
        
        .Add , "Componente", "Componente", 2000
        .Add , "Factor", "Factor", 2000
        
    End With
    
    LvwComponentes.Sorted = False
    
End Sub

Private Function BuscarComponentes() As Boolean
    
    Dim objComponente                   As clsComponente
    Dim Item                            As ListItem
    
    On Error GoTo Error
    BuscarComponentes = True
    For Each objComponente In gObjComponentes

        Set Item = LvwComponentes.ListItems.Add(, objComponente.CodigoComponente & "ID")
        Item.Text = objComponente.NombreComponente
        Item.Tag = objComponente.CodigoFactor
        Item.SubItems(1) = gObjFactores.colFactor(objComponente.CodigoFactor).NombreFactor
    
    Next
    If LvwComponentes.ListItems.Count <> 0 Then
        LvwComponentes.SelectedItem = LvwComponentes.ListItems(1)
        LvwComponentes.SelectedItem.Selected = True
    End If
    
Error:
    
    If Err.Number <> 0 Then
        BuscarComponentes = False
        frmMsg.MostrarMsg "Módulo: BuscarComponentes." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set FrmConfiguracion = Nothing
End Sub

Private Sub LvwComponentes_DblClick()
    
    Select Case LvwComponentes.SelectedItem.SubItems(1)
        
        Case Is = ""
            
            LvwComponentes.SelectedItem.SubItems(1) = "EFast"
            LvwComponentes.SelectedItem.Tag = 1
            
        Case Is = "EFast"
            
            LvwComponentes.SelectedItem.SubItems(1) = "ENormal"
            LvwComponentes.SelectedItem.Tag = 2
            
        Case Is = "ENormal"
            
            LvwComponentes.SelectedItem.SubItems(1) = ""
            LvwComponentes.SelectedItem.Tag = 0
            
    End Select
    
End Sub

Private Function GuardarFactores() As Boolean
    
    Dim CodigoComponente            As Integer
    Dim CodigoFactor                As Integer
    Dim CodigoFactorAnt             As Integer
    Dim objComponente               As clsComponente
    
    Dim Item As ListItem
    
    On Error GoTo Error
    
    GuardarFactores = True
    
    Set gObjComponentes.Datos = gDatos
    
    For Each Item In LvwComponentes.ListItems
        
        CodigoComponente = Val(Item.Key)
        CodigoFactor = Item.Tag
        Set objComponente = gObjComponentes.colComponente(CodigoComponente)
        If Not objComponente Is Nothing Then
            CodigoFactorAnt = objComponente.CodigoFactor
            objComponente.CodigoFactor = CodigoFactor
            
            If Not gObjComponentes.dbModificar(objComponente) Then
                objComponente.CodigoFactor = CodigoFactorAnt
                GuardarFactores = False
                Exit For
            End If
        End If
        
    Next
    
Error:
    If Err.Number <> 0 Then
        GuardarFactores = False
        frmMsg.MostrarMsg "Módulo:GuardarFactores." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function GuardarVariablesGenerales() As Boolean
    
    Dim StrSql                          As String
    Dim ArchivoDeResultados             As String
    Dim ArchivoDeResultadosGenerales    As String
    
    On Error GoTo Error
    
    GuardarVariablesGenerales = False
    
    ArchivoDeResultados = TxtArchivoDeResultados.Text
    ArchivoDeResultadosGenerales = TxtArchivoDeResultadosGenerales.Text
    
    
    StrSql = ""
    StrSql = StrSql & "UPDATE iga_Configuracion SET "
    StrSql = StrSql & "ArchivoDeResultados = '" & ArchivoDeResultados & "' , "
    StrSql = StrSql & "ArchivoDeResultadosGenerales = '" & ArchivoDeResultadosGenerales & "', "
    StrSql = StrSql & "AnalisisAutomaticos = " & IIf(ChkAnalisisAutomaticos.Value = 1, "TRUE", "FALSE") & ", "
    StrSql = StrSql & "CambioEscalas = " & IIf(CheckCambioEscalas.Value = 1, "TRUE", "FALSE") & ", "
    StrSql = StrSql & "TomaProfundidad = " & IIf(CheckTomaProfundidad.Value = 1, "TRUE", "FALSE") & ", "
    StrSql = StrSql & "DisparoCiclico = " & IIf(CheckDisparoCiclico.Value = 1, "TRUE", "FALSE") & ", "
    StrSql = StrSql & "TiempoAnalisis = " & Val(Text1.Text) & ", "
    StrSql = StrSql & "IpRabit = '" & txtIpRabit.Text & "', "
    StrSql = StrSql & "IpDisparoCroma = '" & txtIpDisparoCroma.Text & "', "
    StrSql = StrSql & "IdTipoEquipoCroma = " & gObjTiposEquipoCroma.colTipoEquipoCroma(, cmbTipoEquipoCroma.SelectedItem.Key).IdTipoEquipoCroma & " , "
    StrSql = StrSql & "IdTipoEquipoGas = " & gObjTiposEquipoGas.colTipoEquipoGas(, cmbTipoEquipoGas.SelectedItem.Key).IdTipoEquipoGas & " , "
    StrSql = StrSql & "NivelDisparoMetroMetro = " & Val(TxtNivelDisparoMetroMetro.Text) & ";"
    
    If gDatos.DbEjecutar(StrSql) Then
        
        GuardarVariablesGenerales = True
        
        gObjConfiguracion.ArchivoDeResultados = ArchivoDeResultados
        gObjConfiguracion.ArchivoDeResultadosGenerales = ArchivoDeResultadosGenerales
        gObjConfiguracion.AnalisisAutomaticos = ChkAnalisisAutomaticos.Value = 1
        gObjConfiguracion.CambioEscalas = CheckCambioEscalas.Value = 1
        gObjConfiguracion.TomaProfundidad = CheckTomaProfundidad.Value = 1
        gObjConfiguracion.IpRabit = txtIpRabit.Text
        gObjConfiguracion.IpDisparoCroma = txtIpDisparoCroma.Text
        gObjConfiguracion.TiempoAnalisis = Val(Text1.Text)
        gObjConfiguracion.NivelDisparoMetroMetro = Val(TxtNivelDisparoMetroMetro.Text)
        gObjConfiguracion.DisparoCiclico = CheckDisparoCiclico.Value = 1
        gObjConfiguracion.IdTipoEquipoCroma = gObjTiposEquipoCroma.colTipoEquipoCroma(, cmbTipoEquipoCroma.SelectedItem.Key).IdTipoEquipoCroma
        gObjConfiguracion.IdTipoEquipoGas = gObjTiposEquipoGas.colTipoEquipoGas(, cmbTipoEquipoGas.SelectedItem.Key).IdTipoEquipoGas
        
    End If
Error:
    
    If Err.Number <> 0 Then
        GuardarVariablesGenerales = False
        frmMsg.MostrarMsg "Módulo:GuardarVariablesGenerales." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function BuscarVariablesGenerales() As Boolean
    
    Dim StrSql As String
    Dim Variables As Recordset
    
    On Error GoTo Error
    
    BuscarVariablesGenerales = True
    
    TxtArchivoDeResultados.Text = gObjConfiguracion.ArchivoDeResultados
    TxtArchivoDeResultadosGenerales.Text = gObjConfiguracion.ArchivoDeResultadosGenerales
    ChkAnalisisAutomaticos.Value = IIf(gObjConfiguracion.AnalisisAutomaticos, 1, 0)
    CheckCambioEscalas.Value = IIf(gObjConfiguracion.CambioEscalas, 1, 0)
    CheckTomaProfundidad.Value = IIf(gObjConfiguracion.TomaProfundidad, 1, 0)
    CheckDisparoCiclico.Value = IIf(gObjConfiguracion.DisparoCiclico, 1, 0)
    Text1.Text = gObjConfiguracion.TiempoAnalisis
    TxtNivelDisparoMetroMetro = gObjConfiguracion.NivelDisparoMetroMetro
    txtIpRabit.Text = gObjConfiguracion.IpRabit
    txtIpDisparoCroma.Text = gObjConfiguracion.IpDisparoCroma
    
    Set cmbTipoEquipoCroma.SelectedItem = cmbTipoEquipoCroma.ComboItems(gObjTiposEquipoCroma.colTipoEquipoCroma(gObjConfiguracion.IdTipoEquipoCroma).Key)
    cmbTipoEquipoCroma.Refresh
    Set cmbTipoEquipoGas.SelectedItem = cmbTipoEquipoGas.ComboItems(gObjTiposEquipoGas.colTipoEquipoGas(gObjConfiguracion.IdTipoEquipoGas).Key)
    cmbTipoEquipoGas.Refresh
Error:
    
    If Err.Number <> 0 Then
        BuscarVariablesGenerales = False
        frmMsg.MostrarMsg "Módulo: BuscarVariablesGenerales." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Sub TDBGrid1_AfterUpdate()
    Uno = TxtUno.Text
    Diez = TxtDiez.Text
    Cien = TxtCien.Text
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
            
            SendKeys "{TAB}"
            
        ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 45 Or KeyAscii = 8 Then
            
            'Tecla permitida
            
        Else
            
            KeyAscii = 0
            
        End If

End Sub

Private Sub TxtArchivoDeResultados_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtArchivoDeResultados_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 39 Then
        
        KeyAscii = 0
        
    Else
        
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
    End If
    
End Sub

Private Sub TxtArchivoDeResultadosGenerales_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtArchivoDeResultadosGenerales_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 39 Then
        
        KeyAscii = 0
        
    Else
        
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
    End If
    
End Sub



Private Sub TxtNivelDisparoMetroMetro_KeyPress(KeyAscii As Integer)

    SoloNumeros KeyAscii
    
End Sub
