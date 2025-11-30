VERSION 5.00
Begin VB.Form FrmDatosPozos 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del pozo"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDatosPozos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCmdAceptar 
      Height          =   495
      Left            =   4440
      TabIndex        =   35
      Top             =   3120
      Width           =   1215
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -360
         Picture         =   "FrmDatosPozos.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Frame fraCmdCancelar 
      Height          =   495
      Left            =   5640
      TabIndex        =   33
      Top             =   3120
      Width           =   1455
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -240
         Picture         =   "FrmDatosPozos.frx":0B36
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Frame Frmme1 
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
      Height          =   3165
      Left            =   45
      TabIndex        =   17
      Top             =   -45
      Width           =   7005
      Begin VB.CheckBox ChkPozoActivo 
         BackColor       =   &H00E8E8E8&
         Caption         =   "Pozo activo"
         Height          =   195
         Left            =   2850
         TabIndex        =   8
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E8E8E8&
         Caption         =   "Coordenadas"
         Height          =   945
         Left            =   135
         TabIndex        =   29
         Top             =   2115
         Width           =   6765
         Begin VB.TextBox TxtZ 
            Height          =   285
            Left            =   5400
            TabIndex        =   12
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox TxtY 
            Height          =   285
            Left            =   3255
            TabIndex        =   10
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox TxtX 
            Height          =   315
            Left            =   855
            TabIndex        =   9
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Z"
            Height          =   210
            Left            =   5205
            TabIndex        =   32
            Top             =   420
            Width           =   105
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y"
            Height          =   210
            Left            =   3075
            TabIndex        =   31
            Top             =   420
            Width           =   120
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            Height          =   210
            Left            =   660
            TabIndex        =   30
            Top             =   420
            Width           =   105
         End
      End
      Begin VB.TextBox TxtPais 
         Height          =   285
         Left            =   4380
         TabIndex        =   7
         Top             =   1470
         Width           =   2505
      End
      Begin VB.TextBox TxtProvincia 
         Height          =   285
         Left            =   990
         TabIndex        =   6
         Top             =   1470
         Width           =   2505
      End
      Begin VB.TextBox TxtArea 
         Height          =   285
         Left            =   4380
         TabIndex        =   5
         Top             =   1065
         Width           =   2505
      End
      Begin VB.TextBox TxtCuenca 
         Height          =   285
         Left            =   990
         TabIndex        =   4
         Top             =   1065
         Width           =   2505
      End
      Begin VB.TextBox TxtCategoria 
         Height          =   285
         Left            =   4380
         TabIndex        =   3
         Top             =   660
         Width           =   2505
      End
      Begin VB.TextBox TxtYacimiento 
         Height          =   285
         Left            =   990
         TabIndex        =   2
         Top             =   660
         Width           =   2505
      End
      Begin VB.TextBox TxtCompañia 
         Height          =   285
         Left            =   4380
         TabIndex        =   1
         Top             =   255
         Width           =   2505
      End
      Begin VB.TextBox TxtPozo 
         Height          =   285
         Left            =   990
         TabIndex        =   0
         Top             =   255
         Width           =   2505
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Pais"
         Height          =   210
         Left            =   3615
         TabIndex        =   28
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Provincia"
         Height          =   210
         Left            =   135
         TabIndex        =   27
         Top             =   1530
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Area"
         Height          =   210
         Left            =   3615
         TabIndex        =   26
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Cuenca"
         Height          =   210
         Left            =   135
         TabIndex        =   25
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Categoria"
         Height          =   210
         Left            =   3615
         TabIndex        =   24
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Yacimiento"
         Height          =   210
         Left            =   135
         TabIndex        =   23
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Compañia"
         Height          =   210
         Left            =   3615
         TabIndex        =   22
         Top             =   315
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8E8E8&
         Caption         =   "Nombre"
         Height          =   210
         Left            =   135
         TabIndex        =   18
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.TextBox TxtIDPozo 
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
      Left            =   5640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5250
      WhatsThisHelpID =   10031
      Width           =   765
   End
   Begin VB.CheckBox chkEstoyAgregando 
      Caption         =   "Estoy Agregando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4425
      Width           =   1515
   End
   Begin VB.CheckBox chkEstoyModificando 
      Caption         =   "Estoy Modificando"
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
      Left            =   1320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4770
      Width           =   1605
   End
   Begin VB.CheckBox chkPrepararFormParaModificar 
      Caption         =   "Preparar el Formulario para Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4035
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4485
      Width           =   2925
   End
   Begin VB.CheckBox chkGuardarDatosEnAceptar 
      Caption         =   "Guardar los datos cuando se selecciona aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5040
      WhatsThisHelpID =   10030
      Width           =   2175
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "ID Pozo"
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
      Left            =   4980
      TabIndex        =   21
      Top             =   5295
      WhatsThisHelpID =   10032
      Width           =   570
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Propiedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   1125
      TabIndex        =   20
      Top             =   4020
      Width           =   1605
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Métodos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   3825
      TabIndex        =   19
      Top             =   4050
      Width           =   1095
   End
End
Attribute VB_Name = "FrmDatosPozos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoadOk As Boolean

Private Sub chkPrepararFormParaModificar_Click()
    
    Dim idPozo As Integer
    If Not gBaseDeDatos.Conectada Then
        Exit Sub
    End If
    
    If chkPrepararFormParaModificar.Value = 1 Then
        
        idPozo = TxtIDPozo.Text
        
        If Not PrepararFormParaModificar(idPozo) Then
            
            Unload Me
            
        End If
        
    End If
    
End Sub

Private Sub cmdAceptar_Click()
    
    Dim idPozo As Integer
    If Not gBaseDeDatos.Conectada Then
        Unload Me
        Exit Sub
    End If
    
    If ValidarDatos Then
        
        If chkGuardarDatosEnAceptar.Value = 1 Then
            
            If chkEstoyAgregando.Value = 1 Then
                
                gBaseDeDatos.BeginTrans
                
                If PozoAgregar(idPozo) Then
                    
                    gBaseDeDatos.CommitTrans
                    
                    TxtIDPozo.Text = idPozo
                    
                    Me.Hide
                    
                Else
                    
                    gBaseDeDatos.RollBackTrans
                    
                End If
                
            ElseIf chkEstoyModificando.Value = 1 Then
                
                idPozo = TxtIDPozo.Text
                
                gBaseDeDatos.BeginTrans
                
                If PozoModificar(idPozo) Then
                    
                    gBaseDeDatos.CommitTrans
                    
                    Me.Hide
                    
                Else
                    
                    gBaseDeDatos.RollBackTrans
                    
                End If
                
            End If
            
        Else
            
            Me.Hide
            
        End If
        
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    
    Unload Me
    
End Sub

Private Function ValidarDatos() As Boolean
    
    On Error GoTo Error
    
    ValidarDatos = True
    
    If Trim(TxtPozo.Text) = "" Then
        
        ValidarDatos = False
        
        frmMsg.MostrarMsg "El nombre del pozo no puede estar en blanco.", "Error", MdiPrincipal
        
        TxtPozo.SetFocus
        
        Exit Function
        
    End If
    
    If gObjPozoActivo.idPozo <> 0 Then
        
        If ChkPozoActivo.Value = 1 Then
            
            If chkEstoyAgregando.Value = 1 Then
                
                ValidarDatos = False
                
                frmMsg.MostrarMsg "Solamente puede haber un pozo activo.", "Error", MdiPrincipal
                
                ChkPozoActivo.SetFocus
                
                Exit Function
                
            Else
                
                If gObjPozoActivo.idPozo <> CInt(TxtIDPozo.Text) Then
                    
                    ValidarDatos = False
                    
                    frmMsg.MostrarMsg "Solamente puede haber un pozo activo.", "Error", MdiPrincipal
                    
                    ChkPozoActivo.SetFocus
                    
                    Exit Function
                    
                End If
                
            End If
            
        End If
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        ValidarDatos = False
        
        frmMsg.MostrarMsg "Módulo: ValidarDatos." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Sub Form_Activate()
    
    If Not LoadOk Then
        
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Load()
    
    LoadOk = True
    
End Sub

Private Sub TxtPozo_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtPozo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 39 Then
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtPozo_LostFocus()
    
    TxtPozo.Text = UCase(TxtPozo.Text)
    
End Sub

Private Function PozoAgregarBase(idPozo As Integer) As Boolean
    
    Dim StrSql As String
    Dim MayorCodigo As Recordset
    Dim Pozo As String
    Dim Compañia As String
    Dim Yacimiento As String
    Dim Categoria As String
    Dim Cuenca As String
    Dim Area As String
    Dim Provincia As String
    Dim Pais As String
    Dim X As String
    Dim Y As String
    Dim Z As String
    Dim Activo As Boolean
    
    On Error GoTo Error
    
    PozoAgregarBase = False
    
    Pozo = TxtPozo.Text
    Compañia = TxtCompañia.Text
    Yacimiento = TxtYacimiento.Text
    Categoria = TxtCategoria.Text
    Cuenca = TxtCuenca.Text
    Area = TxtArea.Text
    Provincia = TxtArea.Text
    Pais = TxtPais.Text
    X = TxtX.Text
    Y = TxtY.Text
    Z = TxtZ.Text
    
    If ChkPozoActivo.Value = 1 Then
        
        Activo = True
        
    Else
        
        Activo = False
        
    End If
    
    StrSql = ""
    StrSql = StrSql & "SELECT "
    StrSql = StrSql & "   MAX(IDPozo) AS MayorCodigo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "   Pozos "
    
    Set MayorCodigo = gBaseDeDatos.EjecutarSeleccion(StrSql)
    
    With MayorCodigo
        
        If IsNull(!MayorCodigo) Then
            
            idPozo = 1
            
        Else
            
            idPozo = !MayorCodigo + 1
            
        End If
        
    End With
    
    If idPozo <> 0 Then
        
        StrSql = ""
        StrSql = StrSql & "INSERT INTO Pozos ( "
        StrSql = StrSql & "   IDPozo , "
        StrSql = StrSql & "   Pozo , "
        StrSql = StrSql & "   Compañia , "
        StrSql = StrSql & "   Yacimiento , "
        StrSql = StrSql & "   Categoria , "
        StrSql = StrSql & "   Cuenca , "
        StrSql = StrSql & "   Area , "
        StrSql = StrSql & "   Provincia , "
        StrSql = StrSql & "   Pais , "
        StrSql = StrSql & "   X , "
        StrSql = StrSql & "   Y , "
        StrSql = StrSql & "   Z , "
        StrSql = StrSql & "   Activo ) "
        StrSql = StrSql & "VALUES ( "
        StrSql = StrSql & "   " & idPozo & " , "
        StrSql = StrSql & "   '" & Pozo & "' , "
        StrSql = StrSql & "   '" & Compañia & "' , "
        StrSql = StrSql & "   '" & Yacimiento & "' , "
        StrSql = StrSql & "   '" & Categoria & "' , "
        StrSql = StrSql & "   '" & Cuenca & "' , "
        StrSql = StrSql & "   '" & Area & "' , "
        StrSql = StrSql & "   '" & Provincia & "' , "
        StrSql = StrSql & "   '" & Pais & "' , "
        StrSql = StrSql & "   '" & X & "' , "
        StrSql = StrSql & "   '" & Y & "' , "
        StrSql = StrSql & "   '" & Z & "' , "
        
        If Activo Then
        
            StrSql = StrSql & "   TRUE ) "
            
        Else
            
            StrSql = StrSql & "   FALSE ) "
            
        End If
        
        gBaseDeDatos.EjecutarSeleccion StrSql
        
        If gBaseDeDatos.RegistrosAfectados = 1 Then
            
            ''MdiPrincipal.StatusBar1.Panels(1).Text = "Archivo de resultados: "
            
            If gObjConfiguracion.ArchivoDeResultados = "" Then
                
                ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & "NO CONFIGURADO "
                
            Else
                
                ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & gObjConfiguracion.ArchivoDeResultados
                
            End If
            
            If gObjConfiguracion.ArchivoDeResultadosGenerales = "" Then
                
                ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " - Archivo general: NO CONFIGURADO - Modo de trabajo: "
                
            Else
                
                ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " - Archivo general: " & gObjConfiguracion.ArchivoDeResultadosGenerales & " - Modo de trabajo: "
                
            End If
            
            If gObjConfiguracion.AnalisisAutomaticos Then
                
                ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " AUTOMATICO"
                
            Else
                
                ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " MANUAL"
                
            End If
            
            If Activo Then
                
                gObjPozoActivo.idPozo = idPozo
                
            End If
            
            ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & "  - ID Pozo: " & gobjpozoactivo.idpozo
            
            PozoAgregarBase = True
            
            
            If FormularioCargado("FrmAnalisis") Then
                
                FrmAnalisis.BuscarAnalisis
                FrmAnalisis.LimpiatDatosDefinitivos
                FrmAnalisis.LimpiatDatosTemporales
                
            End If
            
        Else
            
            PozoAgregarBase = False
            
        End If
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        PozoAgregarBase = False
        
        frmMsg.MostrarMsg "Módulo: PozoAgregarBase." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function PrepararFormParaModificar(idPozo As Integer) As Boolean
    
    Dim StrSql As String
    Dim rDatosPozo As Recordset
    
    On Error GoTo Error
    
    PrepararFormParaModificar = True
    
    StrSql = ""
    StrSql = StrSql & "SELECT "
    StrSql = StrSql & "    Pozos.Pozo, "
    StrSql = StrSql & "    Pozos.Compañia, "
    StrSql = StrSql & "    Pozos.Yacimiento, "
    StrSql = StrSql & "    Pozos.Categoria, "
    StrSql = StrSql & "    Pozos.Cuenca, "
    StrSql = StrSql & "    Pozos.Area, "
    StrSql = StrSql & "    Pozos.Provincia, "
    StrSql = StrSql & "    Pozos.Pais, "
    StrSql = StrSql & "    Pozos.X, "
    StrSql = StrSql & "    Pozos.Y, "
    StrSql = StrSql & "    Pozos.Z, "
    StrSql = StrSql & "    Pozos.Activo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "    Pozos "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "   IDPozo = " & idPozo
    
    Set rDatosPozo = gBaseDeDatos.EjecutarSeleccion(StrSql)
    
    If Not rDatosPozo.BOF And Not rDatosPozo.EOF Then
        
        TxtPozo.Text = rDatosPozo!Pozo
        TxtCompañia.Text = rDatosPozo!Compañia
        TxtYacimiento.Text = rDatosPozo!Yacimiento
        TxtCategoria.Text = rDatosPozo!Categoria
        TxtCuenca.Text = rDatosPozo!Cuenca
        TxtArea.Text = rDatosPozo!Area
        TxtProvincia.Text = rDatosPozo!Provincia
        TxtPais.Text = rDatosPozo!Pais
        TxtX.Text = rDatosPozo!X
        TxtY.Text = rDatosPozo!Y
        TxtZ.Text = rDatosPozo!Z
        
        If rDatosPozo!Activo Then
            
            ChkPozoActivo.Value = 1
            
        Else
            
            ChkPozoActivo.Value = 0
            
        End If
        
        rDatosPozo.Close
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        PrepararFormParaModificar = False
        
        frmMsg.MostrarMsg "Módulo: PrepararFormParaModificar." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function PozoModificarBase(idPozo As Integer) As Boolean
    
    Dim StrSql  As String
    Dim Pozo As String
    Dim Compañia As String
    Dim Yacimiento As String
    Dim Cuenca As String
    Dim Area As String
    Dim Provincia As String
    Dim Pais As String
    Dim X As String
    Dim Y As String
    Dim Z As String
    Dim Activo As Boolean
    
    On Error GoTo Error
    
    Pozo = TxtPozo.Text
    Compañia = TxtCompañia.Text
    Yacimiento = TxtYacimiento.Text
    Cuenca = TxtCuenca.Text
    Area = TxtArea.Text
    Provincia = TxtArea.Text
    Pais = TxtPais.Text
    X = TxtX.Text
    Y = TxtY.Text
    Z = TxtZ.Text
    
    If ChkPozoActivo.Value = 1 Then
        
        Activo = True
        
    Else
        
        Activo = False
        
    End If
    
    StrSql = ""
    StrSql = StrSql & "UPDATE "
    StrSql = StrSql & "   Pozos "
    StrSql = StrSql & "SET "
    StrSql = StrSql & "   Pozo = '" & Pozo & "', "
    StrSql = StrSql & "   Compañia = '" & Compañia & "', "
    StrSql = StrSql & "   Yacimiento = '" & Yacimiento & "', "
    StrSql = StrSql & "   Cuenca = '" & Cuenca & "', "
    StrSql = StrSql & "   Area = '" & Area & "', "
    StrSql = StrSql & "   Provincia = '" & Provincia & "', "
    StrSql = StrSql & "   Pais = '" & Pais & "', "
    StrSql = StrSql & "   X = '" & X & "', "
    StrSql = StrSql & "   Y = '" & Y & "', "
    StrSql = StrSql & "   Z = '" & Z & "', "
    
    If Activo Then
        
        StrSql = StrSql & "   Activo = TRUE "
        
    Else
        
        StrSql = StrSql & "   Activo = FALSE "
        
    End If
    
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "   IDPozo = " & idPozo
    
    gBaseDeDatos.EjecutarSeleccion StrSql
    
    If gBaseDeDatos.RegistrosAfectados = 1 Then
        
        ''MdiPrincipal.StatusBar1.Panels(1).Text = "Archivo de resultados: "
        
        If gObjConfiguracion.ArchivoDeResultados = "" Then
            
            ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & "NO CONFIGURADO "
            
        Else
            
            ''MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & gObjConfiguracion.ArchivoDeResultados
            
        End If
        
        If gObjConfiguracion.ArchivoDeResultadosGenerales = "" Then
            
            'MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " - Archivo general: NO CONFIGURADO - Modo de trabajo: "
            
        Else
            
            'MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " - Archivo general: " & gObjConfiguracion.ArchivoDeResultadosGenerales & " - Modo de trabajo: "
            
        End If
        
        If gObjConfiguracion.AnalisisAutomaticos Then
            
            'MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " AUTOMATICO"
            
        Else
            
            'MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & " MANUAL"
            
        End If
        
        If Activo Then
            
            gObjPozoActivo.idPozo = idPozo
            
            'MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & "  - ID Pozo: " & gobjpozoactivo.idpozo
            
        Else
            
            gObjPozoActivo.idPozo = 0
            
            'MdiPrincipal.StatusBar1.Panels(1).Text = 'MdiPrincipal.StatusBar1.Panels(1).Text & "  - ID Pozo: NO HAY POZO ACTIVO"
            
        End If
        
        PozoModificarBase = True
        
        If FormularioCargado("FrmAnalisis") Then
            
            FrmAnalisis.BuscarAnalisis
            FrmAnalisis.LimpiatDatosDefinitivos
            FrmAnalisis.LimpiatDatosTemporales
            
        End If
        
    Else
        
        PozoModificarBase = False
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        PozoModificarBase = False
        
        frmMsg.MostrarMsg "Módulo: PozoModificarBase." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function PozoModificar(idPozo As Integer) As Boolean
    
    PozoModificar = False
    
    If PozoModificarBase(idPozo) Then
        
        PozoModificar = True
        
    End If
    
End Function

Private Function PozoAgregar(idPozo As Integer) As Boolean
    
    PozoAgregar = False
    
    If PozoAgregarBase(idPozo) Then
        
        PozoAgregar = True
        
    End If
    
End Function
