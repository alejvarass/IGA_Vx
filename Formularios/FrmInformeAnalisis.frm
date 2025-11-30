VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInformeAnalisis 
   BackColor       =   &H00E8E8E8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInformeAnalisis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   11280
   Begin VB.Frame fraCmdConsultar 
      Height          =   495
      Left            =   8640
      TabIndex        =   65
      Top             =   3960
      Width           =   1455
      Begin VB.CommandButton cmdConsultar 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -120
         Picture         =   "FrmInformeAnalisis.frx":076A
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   -120
         Width           =   1815
      End
   End
   Begin VB.Frame fraCmdSalir 
      Height          =   495
      Left            =   10200
      TabIndex        =   63
      Top             =   3960
      Width           =   975
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00E8E8E8&
         Height          =   810
         Left            =   -360
         Picture         =   "FrmInformeAnalisis.frx":0BB9
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   -120
         Width           =   1815
      End
   End
   Begin VB.ComboBox CmbPozos 
      Height          =   330
      Left            =   510
      TabIndex        =   0
      Top             =   135
      Width           =   2130
   End
   Begin VB.ComboBox cmbTipoDeAnalisis 
      Height          =   330
      Left            =   7515
      TabIndex        =   1
      Top             =   135
      Width           =   3750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E8E8E8&
      Height          =   3465
      Left            =   45
      TabIndex        =   5
      Top             =   480
      Width           =   11235
      Begin TabDlg.SSTab SSTab1 
         Height          =   3210
         Left            =   3600
         TabIndex        =   4
         Top             =   195
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5662
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BackColor       =   15263976
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Elementos"
         TabPicture(0)   =   "FrmInformeAnalisis.frx":0EE8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label15"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label13"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label12"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label10"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label9"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label8"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LvwComponentes"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "TxtObservaciones"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "TxtGasTotalCromatografico"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "TxtGasTotal"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "TxtSH2"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "TxtCO2"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "TxtProfundidad"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "Relaciones cromatográficas principales"
         TabPicture(1)   =   "FrmInformeAnalisis.frx":0F04
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame9"
         Tab(1).Control(1)=   "Frame11"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Relaciones cromatográficas secundarias"
         TabPicture(2)   =   "FrmInformeAnalisis.frx":0F20
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame10"
         Tab(2).Control(1)=   "Frame8"
         Tab(2).Control(2)=   "Frame7"
         Tab(2).ControlCount=   3
         Begin VB.TextBox TxtProfundidad 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1140
            TabIndex        =   54
            Top             =   2265
            Width           =   945
         End
         Begin VB.TextBox TxtCO2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4590
            TabIndex        =   53
            Top             =   2265
            Width           =   945
         End
         Begin VB.TextBox TxtSH2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6540
            TabIndex        =   52
            Top             =   2265
            Width           =   945
         End
         Begin VB.TextBox TxtGasTotal 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2820
            TabIndex        =   51
            Top             =   2265
            Width           =   945
         End
         Begin VB.TextBox TxtGasTotalCromatografico 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1140
            TabIndex        =   50
            Top             =   2700
            Width           =   945
         End
         Begin VB.TextBox TxtObservaciones 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2820
            TabIndex        =   49
            Top             =   2700
            Width           =   4665
         End
         Begin VB.Frame Frame9 
            Caption         =   "Wtness/Balance/Character"
            Height          =   1035
            Left            =   -74955
            TabIndex        =   42
            Top             =   465
            Width           =   7470
            Begin VB.TextBox TxtCH 
               Enabled         =   0   'False
               Height          =   285
               Left            =   5490
               TabIndex        =   45
               Top             =   435
               Width           =   945
            End
            Begin VB.TextBox TxtBH 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3570
               TabIndex        =   44
               Top             =   435
               Width           =   945
            End
            Begin VB.TextBox TxtWH 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1575
               TabIndex        =   43
               Top             =   435
               Width           =   945
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "CH"
               Height          =   210
               Left            =   5070
               TabIndex        =   48
               Top             =   495
               Width           =   210
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "BH"
               Height          =   210
               Left            =   3180
               TabIndex        =   47
               Top             =   495
               Width           =   210
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "WH"
               Height          =   210
               Left            =   1185
               TabIndex        =   46
               Top             =   495
               Width           =   255
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Porcentajes"
            Height          =   1590
            Left            =   -74955
            TabIndex        =   31
            Top             =   1530
            Width           =   7470
            Begin VB.TextBox TxtC1 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1575
               TabIndex        =   36
               Top             =   450
               Width           =   945
            End
            Begin VB.TextBox TxtC2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3570
               TabIndex        =   35
               Top             =   450
               Width           =   945
            End
            Begin VB.TextBox TxtC3 
               Enabled         =   0   'False
               Height          =   285
               Left            =   5490
               TabIndex        =   34
               Top             =   420
               Width           =   945
            End
            Begin VB.TextBox TxtC4 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2550
               TabIndex        =   33
               Top             =   1050
               Width           =   945
            End
            Begin VB.TextBox TxtC5 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4665
               TabIndex        =   32
               Top             =   1050
               Width           =   945
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "C1"
               Height          =   210
               Left            =   1185
               TabIndex        =   41
               Top             =   510
               Width           =   195
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "C2"
               Height          =   210
               Left            =   3180
               TabIndex        =   40
               Top             =   510
               Width           =   195
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "C3"
               Height          =   210
               Left            =   5070
               TabIndex        =   39
               Top             =   510
               Width           =   195
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "C4"
               Height          =   210
               Left            =   2250
               TabIndex        =   38
               Top             =   1110
               Width           =   195
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "C5"
               Height          =   210
               Left            =   4350
               TabIndex        =   37
               Top             =   1110
               Width           =   195
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Others"
            Height          =   1590
            Left            =   -73020
            TabIndex        =   20
            Top             =   1530
            Width           =   3720
            Begin VB.TextBox TxtSnGeo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1605
               TabIndex        =   25
               Top             =   210
               Width           =   945
            End
            Begin VB.TextBox TxtGeo1 
               Enabled         =   0   'False
               Height          =   285
               Left            =   765
               TabIndex        =   24
               Top             =   660
               Width           =   945
            End
            Begin VB.TextBox TxtGeo2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   23
               Top             =   660
               Width           =   945
            End
            Begin VB.TextBox TxtGeo3 
               Enabled         =   0   'False
               Height          =   285
               Left            =   765
               TabIndex        =   22
               Top             =   1170
               Width           =   945
            End
            Begin VB.TextBox TxtGeo4 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   21
               Top             =   1170
               Width           =   945
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "Sn Geo"
               Height          =   210
               Left            =   975
               TabIndex        =   30
               Top             =   270
               Width           =   540
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               Caption         =   "Geo1"
               Height          =   210
               Left            =   255
               TabIndex        =   29
               Top             =   720
               Width           =   390
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               Caption         =   "Geo2"
               Height          =   210
               Left            =   2115
               TabIndex        =   28
               Top             =   720
               Width           =   390
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Geo3"
               Height          =   210
               Left            =   255
               TabIndex        =   27
               Top             =   1230
               Width           =   390
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Geo4"
               Height          =   210
               Left            =   2115
               TabIndex        =   26
               Top             =   1230
               Width           =   390
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Coustau"
            Height          =   1035
            Left            =   -71220
            TabIndex        =   15
            Top             =   465
            Width           =   3720
            Begin VB.TextBox TxtCous1 
               Enabled         =   0   'False
               Height          =   285
               Left            =   765
               TabIndex        =   17
               Top             =   435
               Width           =   945
            End
            Begin VB.TextBox TxtCous2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   16
               Top             =   435
               Width           =   945
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Cous1"
               Height          =   210
               Left            =   255
               TabIndex        =   19
               Top             =   495
               Width           =   465
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Cous2"
               Height          =   210
               Left            =   2115
               TabIndex        =   18
               Top             =   495
               Width           =   465
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Baroid (Pixler)"
            Height          =   1035
            Left            =   -74955
            TabIndex        =   6
            Top             =   465
            Width           =   3720
            Begin VB.TextBox TxtBar2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   645
               TabIndex        =   10
               Top             =   255
               Width           =   945
            End
            Begin VB.TextBox TxtBar3 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2565
               TabIndex        =   9
               Top             =   255
               Width           =   945
            End
            Begin VB.TextBox TxtBar4 
               Enabled         =   0   'False
               Height          =   285
               Left            =   645
               TabIndex        =   8
               Top             =   630
               Width           =   930
            End
            Begin VB.TextBox TxtBar5 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2565
               TabIndex        =   7
               Top             =   630
               Width           =   945
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Bar2"
               Height          =   210
               Left            =   255
               TabIndex        =   14
               Top             =   315
               Width           =   345
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Bar3"
               Height          =   210
               Left            =   2145
               TabIndex        =   13
               Top             =   315
               Width           =   345
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Bar4"
               Height          =   210
               Left            =   255
               TabIndex        =   12
               Top             =   705
               Width           =   345
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Bar5"
               Height          =   210
               Left            =   2145
               TabIndex        =   11
               Top             =   705
               Width           =   345
            End
         End
         Begin MSComctlLib.ListView LvwComponentes 
            Height          =   1830
            Left            =   45
            TabIndex        =   2
            Top             =   330
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   3228
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Profundidad"
            Height          =   210
            Left            =   135
            TabIndex        =   60
            Top             =   2310
            Width           =   870
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CO2"
            Height          =   210
            Left            =   4215
            TabIndex        =   59
            Top             =   2310
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "SH2"
            Height          =   210
            Left            =   6150
            TabIndex        =   58
            Top             =   2310
            Width           =   300
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Gas total"
            Height          =   210
            Left            =   2145
            TabIndex        =   57
            Top             =   2310
            Width           =   645
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Gas tot. crom."
            Height          =   210
            Left            =   135
            TabIndex        =   56
            Top             =   2745
            Width           =   1020
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Observ."
            Height          =   210
            Left            =   2145
            TabIndex        =   55
            Top             =   2745
            Width           =   585
         End
      End
      Begin MSComctlLib.ListView LvwAnalisis 
         Height          =   3195
         Left            =   30
         TabIndex        =   3
         Top             =   195
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   675
      Top             =   6135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInformeAnalisis.frx":0F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInformeAnalisis.frx":1390
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInformeAnalisis.frx":17E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInformeAnalisis.frx":1C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInformeAnalisis.frx":208C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pozo"
      Height          =   210
      Left            =   75
      TabIndex        =   62
      Top             =   195
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de analisis"
      Height          =   210
      Left            =   6375
      TabIndex        =   61
      Top             =   195
      Width           =   1110
   End
End
Attribute VB_Name = "FrmInformeAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LoadOk                              As Boolean
Private BotonDerecho                        As Boolean
Private pObjPozos                           As New clsPozos

Private Sub CmbPozos_Click()
    
    LvwAnalisis.ListItems.Clear
    LimpiatDatos
    
End Sub

Private Sub CmbPozos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub cmbTipoDeAnalisis_Click()
    
    LvwAnalisis.ListItems.Clear
    LimpiatDatos
    
End Sub

Private Sub cmbTipoDeAnalisis_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub CmdConsultar_Click()
    
    If ValidarDatos Then
        BuscarAnalisis
    End If
    
End Sub

Private Sub CmdSalir_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    If Not LoadOk Then
        
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Load()
    
    LoadOk = True
    
    CentrarFormulario Me
    SetearList
    
    If TiposDeAnalisisBuscar() Then
        
        If PozosBuscar() Then
            
            LoadOk = True
            
        End If
        
    End If
    
End Sub

Private Sub SetearList()
    
    With LvwAnalisis.ColumnHeaders
        
        .Add , "Nº", "Nº", 700
        .Add , "Fecha", "Fecha", 1730
        .Add , "Prof.", "Prof.", 800
        
    End With
    
    LvwAnalisis.Sorted = True
    
    With LvwComponentes.ColumnHeaders
        
        .Add , "Numero", "Numero", 800
        .Add , "Component", "Component", 1500
        .Add , "Retention", "Retention", 1000
        .Add , "Area", "Area", 1000
        .Add , "Externo", "Externo", 1050
        .Add , "Units", "Units", 1000
        .Add , "NormArea", "NormArea", 1000
        
    End With
    
    LvwComponentes.Sorted = True
    LvwComponentes.SortKey = 0
    
End Sub

Private Sub LvwAnalisis_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim NumeroAnalisis As Long
    
    If LvwAnalisis.ListItems.Count <> 0 Then
        NumeroAnalisis = Val(LvwAnalisis.SelectedItem.Key)
        BuscarDatosAnalisis NumeroAnalisis
    End If
    
End Sub

Private Function BuscarAnalisis() As Boolean
    
    Dim Item                            As ListItem
    
    Dim CodigoTipoDeAnalisis            As Integer
    Dim IdPozo                          As Integer
    
    Dim objPozo                         As clsPozo
    Dim objPozoAnalis                   As clsPozoAnalis
    
    On Error GoTo Error
    
    LvwAnalisis.ListItems.Clear
        
    IdPozo = CmbPozos.ItemData(CmbPozos.ListIndex)
    CodigoTipoDeAnalisis = cmbTipoDeAnalisis.ItemData(cmbTipoDeAnalisis.ListIndex)
    
    Set objPozo = pObjPozos.colPozo(IdPozo)
    If Not objPozo Is Nothing Then
        Set objPozo.PozoAnalisis.Datos = gDatos
        objPozo.PozoAnalisis.ConsultarPozoAnalisis objPozo.IdPozo, , , , CodigoTipoDeAnalisis
    
        For Each objPozoAnalis In objPozo.PozoAnalisis
            Set Item = LvwAnalisis.ListItems.Add(, objPozoAnalis.NumeroAnalisis & "ID")
            
            Item.Tag = objPozoAnalis.CodigoTipoDeAnalisis
            
            Item.Text = Format(objPozoAnalis.NumeroAnalisis, "00000")
            Item.SubItems(1) = objPozoAnalis.NumeroAnalisis
            Item.SubItems(2) = Format(objPozoAnalis.fecha, "DD/MM/YYYY HH:MM:SS")
            Item.SubItems(3) = objPozoAnalis.Profundidad
            
            Select Case objPozoAnalis.CodigoTipoDeAnalisis
                Case Is = 1
                    Item.SmallIcon = 1
                Case Is = 2
                    Item.SmallIcon = 2
                Case Is = 3
                    Item.SmallIcon = 3
                Case Is = 4
                    Item.SmallIcon = 4
                Case Is = 5
                    Item.SmallIcon = 5
            End Select
        Next
        
        If LvwAnalisis.ListItems.Count <> 0 Then
            LvwAnalisis.SelectedItem = LvwAnalisis.ListItems(1)
            LvwAnalisis.SelectedItem.Selected = True
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        BuscarAnalisis = False
        frmMsg.MostrarMsg "Módulo: BuscarAnalisis." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function BuscarDatosAnalisis(NumeroAnalisis As Long) As Boolean
    
    Dim IdPozo                      As Long
    Dim objPozo                     As clsPozo
    Dim objPozoAnalis               As clsPozoAnalis
    Dim objPozoAnalisComponente     As ClsPozoAnalisComponente
    Dim Item                        As ListItem
    
    On Error GoTo Error
    
    BuscarDatosAnalisis = True
    
    IdPozo = CmbPozos.ItemData(CmbPozos.ListIndex)
    
    Set objPozo = pObjPozos.colPozo(IdPozo)
    If Not objPozo Is Nothing Then
        Set objPozoAnalis = objPozo.PozoAnalisis.ColPozoAnalis(objPozo.IdPozo, NumeroAnalisis)
    
        If Not objPozoAnalis Is Nothing Then
                
            TxtProfundidad.Text = objPozoAnalis.Profundidad
            TxtCO2.Text = objPozoAnalis.CO2
            TxtSH2.Text = objPozoAnalis.SH2
            TxtGasTotal.Text = objPozoAnalis.GasTotal
            TxtGasTotalCromatografico.Text = objPozoAnalis.GasTotalCromatografico
            If Not IsNull(objPozoAnalis.Observaciones) Then
                TxtObservaciones.Text = objPozoAnalis.Observaciones
            End If
            '************RelacionesCromatograficas
            TxtBar2.Text = objPozoAnalis.Bar2
            TxtBar3.Text = objPozoAnalis.Bar3
            TxtBar4.Text = objPozoAnalis.Bar4
            TxtBar5.Text = objPozoAnalis.Bar5
            TxtCous1.Text = objPozoAnalis.Cous1
            TxtCous2.Text = objPozoAnalis.Cous2
            TxtWH.Text = objPozoAnalis.WH
            TxtBH.Text = objPozoAnalis.BH
            TxtCH.Text = objPozoAnalis.CH
            TxtGeo1.Text = objPozoAnalis.Geo1
            TxtGeo2.Text = objPozoAnalis.Geo2
            TxtGeo3.Text = objPozoAnalis.Geo3
            TxtGeo4.Text = objPozoAnalis.Geo4
            TxtSnGeo.Text = objPozoAnalis.SnGeo
            TxtC1.Text = objPozoAnalis.C1
            TxtC2.Text = objPozoAnalis.C2
            TxtC3.Text = objPozoAnalis.C3
            TxtC4.Text = objPozoAnalis.C4
            TxtC5.Text = objPozoAnalis.C5
            
            '***************Componentes
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
                Set Item = LvwComponentes.ListItems.Add(, objPozoAnalisComponente.NumeroComponente & "ID")
                
                Item.Text = objPozoAnalisComponente.NumeroComponente
                Item.SubItems(1) = objPozoAnalisComponente.Component
                Item.SubItems(2) = objPozoAnalisComponente.Retention
                Item.SubItems(3) = objPozoAnalisComponente.Area
                Item.SubItems(4) = objPozoAnalisComponente.Externo
                Item.SubItems(5) = objPozoAnalisComponente.Units
                Item.SubItems(6) = objPozoAnalisComponente.NormArea
                
            Next

            If LvwComponentes.ListItems.Count <> 0 Then
                LvwComponentes.SelectedItem = LvwComponentes.ListItems(1)
                LvwComponentes.SelectedItem.Selected = True
            End If

        Else
            TxtProfundidad.Text = ""
            TxtCO2.Text = ""
            TxtSH2.Text = ""
            TxtGasTotal.Text = ""
            TxtGasTotalCromatografico.Text = ""
            '************RelacionesCromatograficas
            TxtBar2.Text = ""
            TxtBar3.Text = ""
            TxtBar4.Text = ""
            TxtBar5.Text = ""
            TxtCous1.Text = ""
            TxtCous2.Text = ""
            TxtWH.Text = ""
            TxtBH.Text = ""
            TxtCH.Text = ""
            TxtGeo1.Text = ""
            TxtGeo2.Text = ""
            TxtGeo3.Text = ""
            TxtGeo4.Text = ""
            TxtSnGeo.Text = ""
            TxtC1.Text = ""
            TxtC2.Text = ""
            TxtC3.Text = ""
            TxtC4.Text = ""
            TxtC5.Text = ""
        
        End If
            
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        BuscarDatosAnalisis = False
        
        frmMsg.MostrarMsg "Módulo: BuscarDatosAnalisis." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function ValidarDatos() As Boolean
    
    On Error GoTo Error
    
    ValidarDatos = True
    If CmbPozos.ListIndex = -1 Then
        ValidarDatos = False
        frmMsg.MostrarMsg "Debe seleccionar el pozo.", "Error", MdiPrincipal
        CmbPozos.SetFocus
        Exit Function
    End If
    
    If cmbTipoDeAnalisis.ListIndex = -1 Then
        ValidarDatos = False
        frmMsg.MostrarMsg "Debe seleccionar el tipo de análisis.", "Error", MdiPrincipal
        cmbTipoDeAnalisis.SetFocus
        Exit Function
    End If
    
Error:
    
    If Err.Number <> 0 Then
        ValidarDatos = False
        frmMsg.MostrarMsg "Módulo: ValidarDatos." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function TiposDeAnalisisBuscar() As Boolean
    
    Dim objTipoDeAnalisis                     As clsTipoDeAnalisis
    
    On Error GoTo Error
    
    TiposDeAnalisisBuscar = True
    
    For Each objTipoDeAnalisis In gObjTiposDeAnalisis
        cmbTipoDeAnalisis.AddItem objTipoDeAnalisis.NombreTipoDeAnalisis
        cmbTipoDeAnalisis.ItemData(cmbTipoDeAnalisis.NewIndex) = objTipoDeAnalisis.CodigoTipoDeAnalisis
    Next
Error:
    
    If Err.Number <> 0 Then
        TiposDeAnalisisBuscar = False
        frmMsg.MostrarMsg "Módulo: TiposDeAnalisisBuscar." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Sub LimpiatDatos()
    
    LvwComponentes.ListItems.Clear
    
    TxtBar2.Text = ""
    TxtBar3.Text = ""
    TxtBar4.Text = ""
    TxtBar5.Text = ""
    TxtCous1.Text = ""
    TxtCous2.Text = ""
    TxtWH.Text = ""
    TxtBH.Text = ""
    TxtCH.Text = ""
    TxtGeo1.Text = ""
    TxtGeo2.Text = ""
    TxtGeo3.Text = ""
    TxtGeo4.Text = ""
    TxtSnGeo.Text = ""
    TxtC1.Text = ""
    TxtC2.Text = ""
    TxtC3.Text = ""
    TxtC4.Text = ""
    TxtC5.Text = ""
    
    TxtProfundidad.Text = ""
    TxtCO2.Text = ""
    TxtSH2.Text = ""
    TxtGasTotal.Text = ""
    TxtGasTotalCromatografico.Text = ""
    
End Sub

Private Function PozosBuscar() As Boolean
    
    Dim objPozo         As clsPozo
    
    On Error GoTo Error
    
    PozosBuscar = True
    
    Set pObjPozos.Datos = gDatos
    pObjPozos.ConsultarPozos
    
    For Each objPozo In pObjPozos
        CmbPozos.AddItem objPozo.NombrePozo
        CmbPozos.ItemData(CmbPozos.NewIndex) = objPozo.IdPozo
    Next
    
Error:
    
    If Err.Number <> 0 Then
        PozosBuscar = False
        frmMsg.MostrarMsg "Módulo: PozosBuscar." & Chr(10) & "Ocurrió el error " & Err.Number & ": " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function
