VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAbmPozos 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pozos"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbmPozos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6000
   Begin VB.Frame fraCmdSalir 
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   4560
      Width           =   975
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -360
         Picture         =   "FrmAbmPozos.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   -120
         Width           =   1815
      End
   End
   Begin VB.TextBox txtPopupElegido 
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
      Left            =   2310
      TabIndex        =   0
      Top             =   5850
      Width           =   1170
   End
   Begin MSComctlLib.ListView LvwPozos 
      Height          =   4440
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   7832
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Popup Elegido"
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
      Left            =   1170
      TabIndex        =   2
      Top             =   5895
      Width           =   1035
   End
End
Attribute VB_Name = "FrmAbmPozos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BotonDerecho As Boolean
Private LoadOk As Boolean

Private Sub CmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    If Not LoadOk Then
        
        Unload Me
        
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 45 Then
        
        PozoAgregar
        
    ElseIf KeyCode = 46 Then
        
        PozoBorrar
        
    End If
    
End Sub

Private Sub Form_Load()
    If Not gBaseDeDatos.Conectada Then
        Unload Me
        Exit Sub
    End If
    
    LoadOk = False
    CentrarFormulario Me
    SetearList
    
    If BuscarPozos Then
        
        LoadOk = True
        
    End If
    
End Sub

Private Sub LvwPozos_Click()
    If Not gBaseDeDatos.Conectada Then
        Exit Sub
    End If
    
    If BotonDerecho Then
        
        MdiPrincipal.mnuPopupPozosBorrar.Enabled = LvwPozos.ListItems.Count <> 0
        MdiPrincipal.mnuPopupPozosModificar.Enabled = LvwPozos.ListItems.Count <> 0
        
        Me.PopupMenu MdiPrincipal.mnuPopupPozos
        
        If txtPopupElegido.Text = "AGREGAR" Then
            
            PozoAgregar
            
        ElseIf txtPopupElegido.Text = "MODIFICAR" Then
            
            PozoModificar
            
        ElseIf txtPopupElegido.Text = "BORRAR" Then
            
            PozoBorrar
            
        End If
        
        txtPopupElegido.Text = ""
        
    End If
    
End Sub

Private Sub LvwPozos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If LvwPozos.SortKey = ColumnHeader.Index - 1 Then
        
        If LvwPozos.SortOrder = lvwAscending Then
            
            LvwPozos.SortOrder = lvwDescending
            
        Else
            
            LvwPozos.SortOrder = lvwAscending
            
        End If
        
    Else
        
        LvwPozos.SortOrder = lvwAscending
        LvwPozos.SortKey = ColumnHeader.Index - 1
        
    End If
    
End Sub

Private Sub LvwPozos_DblClick()
    If Not gBaseDeDatos.Conectada Then
        Exit Sub
    End If
    
    If LvwPozos.ListItems.Count <> 0 Then
        
        PozoModificar
        
    End If
    
End Sub

Private Sub LvwPozos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        LvwPozos_DblClick
        
    End If
    
End Sub

Private Sub LvwPozos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BotonDerecho = Button = vbRightButton
    
End Sub

Public Sub PozoAgregar()
    
    Dim IdPozo As Integer
    Dim Pozo As String
    
    FrmDatosPozos.chkEstoyAgregando.Value = 1
    FrmDatosPozos.chkGuardarDatosEnAceptar.Value = 1
    FrmDatosPozos.Show vbModal
    
    If FormularioCargado("FrmDatosPozos") Then
        
        IdPozo = FrmDatosPozos.TxtIDPozo.Text
        Pozo = FrmDatosPozos.TxtPozo.Text
        
        Unload FrmDatosPozos
        
        If PozoAgregarList(IdPozo, Pozo) Then
            
            LvwPozos.SetFocus
            
        Else
            
            Unload Me
            
        End If
        
    Else
        
        LvwPozos.SetFocus
        
    End If
    
End Sub

Private Function PozoAgregarList(IdPozo As Integer, Pozo As String) As Boolean
    
    Dim Pozos As ListItem
    
    On Error GoTo Error
    
    PozoAgregarList = True
    
    Set Pozos = LvwPozos.ListItems.Add(, IdPozo & "ID")
    
    With Pozos
        
        .Text = Pozo
        
    End With
    
    Set LvwPozos.SelectedItem = Pozos
    LvwPozos.SelectedItem.Selected = True
    
Error:
    
    If Err.Number <> 0 Then
        
        PozoAgregarList = False
        
        frmMsg.MostrarMsg "Módulo: PozoAgregarList." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function BuscarPozos() As Boolean
    
    Dim StrSql As String
    Dim Pozos As Recordset
    Dim Item As ListItem
    
    On Error GoTo Error
    
    BuscarPozos = True
    
    StrSql = ""
    StrSql = StrSql & "SELECT "
    StrSql = StrSql & "   IDPozo, "
    StrSql = StrSql & "   Pozo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "   Pozos "
    StrSql = StrSql & "ORDER BY "
    StrSql = StrSql & "   Pozo "
    
    Set Pozos = gBaseDeDatos.EjecutarSeleccion(StrSql)
    
    With Pozos
        
        Do Until .EOF
            
            Set Item = LvwPozos.ListItems.Add(, !IdPozo & "ID")
            
            Item.Text = !Pozo
            
            .MoveNext
            
        Loop
        
        .Close
        
    End With
    
    If LvwPozos.ListItems.Count <> 0 Then
        
        LvwPozos.SelectedItem = LvwPozos.ListItems(1)
        LvwPozos.SelectedItem.Selected = True
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        BuscarPozos = False
        
        frmMsg.MostrarMsg "Módulo: BuscarPozos." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Public Sub PozoBorrar()
    
    Dim IdPozo As Integer
    Dim Pozo As String
    Dim Respuesta As Integer
    
    IdPozo = Val(LvwPozos.SelectedItem.Key)
    Pozo = LvwPozos.SelectedItem.Text
    
    Respuesta = MsgBox("Está seguro que desea borrar el pozo '" & Pozo & "'.", vbYesNo + vbQuestion)
    
    If Respuesta = vbYes Then
        
        gBaseDeDatos.BeginTrans
        
        If PozoBorrarBase(IdPozo, Pozo) Then
            
            gBaseDeDatos.CommitTrans
            
            If PozoBorrarList Then
                
                LvwPozos.SetFocus
                
            Else
                
                Unload Me
                
            End If
            
        Else
            
            gBaseDeDatos.RollBackTrans
            LvwPozos.SetFocus
            
        End If
        
    Else
        
        LvwPozos.SetFocus
        
    End If
    
End Sub

Private Function PozoBorrarBase(IdPozo As Integer, Pozo As String) As Boolean
    
    Dim StrSql As String
    
    On Error GoTo Error
    
    PozoBorrarBase = True
    
    StrSql = ""
    StrSql = StrSql & "DELETE FROM Pozos "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "   IDPozo = " & IdPozo
    
    gBaseDeDatos.EjecutarSeleccion StrSql
    
Error:
    
    If Err.Number <> 0 Then
        
        PozoBorrarBase = False
        
        If Err.Number = 3200 Then
            
            frmMsg.MostrarMsg "No se puede borrar el pozo '" & Pozo & "' porque esto afectaría a la integridad de los datos del sistema.", "Error", MdiPrincipal
            
        Else
            
            frmMsg.MostrarMsg "Módulo:PozoBorrarBase." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
            
        End If
        
        Err.Clear
        
    End If
    
End Function

Private Function PozoBorrarList() As Boolean
    
    On Error GoTo Error
    
    PozoBorrarList = True
    
    LvwPozos.ListItems.Remove LvwPozos.SelectedItem.Index
    
    If LvwPozos.ListItems.Count <> 0 Then
        
        LvwPozos.SelectedItem.Selected = True
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        PozoBorrarList = False
        
        frmMsg.MostrarMsg "Módulo: PozoBorrarList." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Public Sub PozoModificar()
    
    Dim Pozo As String
    
    FrmDatosPozos.TxtIDPozo.Text = Val(LvwPozos.SelectedItem.Key)
    FrmDatosPozos.chkPrepararFormParaModificar.Value = 1
    FrmDatosPozos.chkEstoyModificando.Value = 1
    FrmDatosPozos.chkGuardarDatosEnAceptar.Value = 1
    
    FrmDatosPozos.Show vbModal
    
    If FormularioCargado("FrmDatosPozos") Then
        
        Pozo = FrmDatosPozos.TxtPozo.Text
        
        Unload FrmDatosPozos
        
        If PozoModificarList(Pozo) Then
            
            LvwPozos.SetFocus
            
        Else
            
            Unload Me
            
        End If
        
    Else
        
        LvwPozos.SetFocus
        
    End If
    
End Sub

Private Function PozoModificarList(Pozo As String) As Boolean
    
    Dim Pozos As ListItem
    
    On Error GoTo Error
    
    PozoModificarList = True
    
    Set Pozos = LvwPozos.SelectedItem
    
    With Pozos
        
        .Text = Pozo
        
    End With
    
    Set LvwPozos.SelectedItem = Pozos
    LvwPozos.SelectedItem.Selected = True
    
Error:
    
    If Err.Number <> 0 Then
        
        PozoModificarList = False
        
        frmMsg.MostrarMsg "Módulo: PozoModificarList." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Sub SetearList()
    
    With LvwPozos.ColumnHeaders
        
        .Add , "Nombre", "Nombre", 5610
        
    End With
    
    LvwPozos.Sorted = True
    
End Sub


