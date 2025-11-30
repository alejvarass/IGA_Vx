VERSION 5.00
Object = "{A8B345A0-74B5-11D3-85C2-00105AC8B715}#1.0#0"; "iProfessionalLibrary.ocx"
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E1C1B26-94B1-11D0-B18F-0000E8CA3ED9}#1.0#0"; "TAS.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{575C54CD-73C5-437B-B179-A0B0F35AC1A7}#1.0#0"; "cstcpctl.ocx"
Begin VB.MDIForm MdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "IGA Server"
   ClientHeight    =   8970
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   10455
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin SocketWrenchCtl.SocketWrench Sw 
      Left            =   630
      Top             =   720
      _cx             =   741
      _cy             =   741
   End
   Begin iProfessionalLibrary.iTimersX TimerCroma 
      Left            =   4725
      Top             =   3375
      Enabled1        =   0   'False
      Enabled2        =   0   'False
      Enabled3        =   0   'False
      Enabled4        =   0   'False
      Enabled5        =   0   'False
      Enabled6        =   0   'False
      Enabled7        =   0   'False
      Enabled8        =   0   'False
      Enabled9        =   0   'False
      Interval1       =   90000
      Interval2       =   7500
      Interval3       =   5000
      Interval4       =   1000
      Interval5       =   1000
      Interval6       =   1000
      Interval7       =   1000
      Interval8       =   1000
      Interval9       =   1000
   End
   Begin InetCtlsObjects.Inet InetDisparoCroma 
      Left            =   60
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
      RequestTimeout  =   10
   End
   Begin VB.Timer TimerSimulador 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   420
      Top             =   3525
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1275
      Top             =   3510
   End
   Begin VB.Timer TimerSalidas 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4170
      Top             =   4200
   End
   Begin VB.Timer TimerEntradas 
      Interval        =   1500
      Left            =   5130
      Top             =   4200
   End
   Begin VB.Timer TimerResultados 
      Interval        =   5000
      Left            =   4650
      Top             =   4200
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.Frame fraCmdIniciarAnalisis 
         Height          =   510
         Left            =   6945
         TabIndex        =   12
         Top             =   45
         Width           =   1950
         Begin VB.CommandButton cmdIniciarAnalisis 
            DisabledPicture =   "FrmPrincipal.frx":0000
            Height          =   735
            Left            =   -15
            Picture         =   "FrmPrincipal.frx":0524
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   -105
            Width           =   2175
         End
      End
      Begin isDigitalLibrary.iLedRoundX iLedStatusCroma 
         Height          =   285
         Left            =   6600
         TabIndex        =   11
         Top             =   180
         Width           =   285
         BackGroundColor =   -2147483633
         Active          =   0   'False
         ActiveColor     =   4259584
         BevelStyle      =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         ShowReflection  =   -1  'True
         AutoInactiveColor=   -1  'True
         InactiveColor   =   12632256
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   19
         Object.Height          =   19
         OPCItemCount    =   0
      End
      Begin TASLib.TasSerial TasSerial2 
         Height          =   2340
         Left            =   14010
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   4128
         _StockProps     =   6
         Caption         =   "TAS(1)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScanAutoTrigger =   0   'False
         DriverDataType  =   2
         CommPort        =   2
         CommDataBits    =   4
         CommParity      =   1
         CommBaudRate    =   15
         DriverErrorLimit=   3
         ImgBackColor    =   13160664
         DriverName      =   "XMODBUSA"
         DriverP0        =   "2"
         DriverP1        =   "6"
         DriverP2        =   "0"
      End
      Begin TASLib.TasSerial TasSerial1 
         Height          =   2340
         Left            =   12570
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   4128
         _StockProps     =   6
         Caption         =   "TAS(0)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScanAutoTrigger =   0   'False
         DriverNumPoints =   18
         CommPort        =   5
         CommDataBits    =   4
         CommParity      =   1
         CommBaudRate    =   15
         DriverErrorLimit=   20
         ImgBackColor    =   13160664
         ScanRate        =   1500
         CommTimeout     =   4500
         DriverName      =   "XMODBUSA"
         DriverP0        =   "2"
         DriverP1        =   "4"
         DriverP2        =   "100"
      End
      Begin VB.TextBox TxtProfundidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11625
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   90
         Width           =   780
      End
      Begin isDigitalLibrary.iLedRoundX LedAnalisisAutomaticos 
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   180
         Width           =   285
         BackGroundColor =   -2147483633
         Active          =   0   'False
         ActiveColor     =   4259584
         BevelStyle      =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         ShowReflection  =   -1  'True
         AutoInactiveColor=   -1  'True
         InactiveColor   =   12632256
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   19
         Object.Height          =   19
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX1 
         Height          =   375
         Left            =   3810
         TabIndex        =   6
         Top             =   120
         Width           =   2445
         AutoSize        =   -1  'True
         Alignment       =   1
         BorderStyle     =   0
         Caption         =   "Análisis Automáticos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   0   'False
         ShadowXOffset   =   1
         ShadowYOffset   =   1
         ShadowColor     =   0
         BackGroundColor =   -2147483633
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   0
         Transparent     =   0   'False
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   163
         Object.Height          =   25
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX2 
         Height          =   375
         Left            =   12510
         TabIndex        =   5
         Top             =   90
         Width           =   2460
         AutoSize        =   0   'False
         Alignment       =   1
         BorderStyle     =   0
         Caption         =   "Profundidad Retorno"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   0   'False
         ShadowXOffset   =   1
         ShadowYOffset   =   1
         ShadowColor     =   8421504
         BackGroundColor =   -2147483633
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   0
         Transparent     =   0   'False
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   164
         Object.Height          =   25
         OPCItemCount    =   0
      End
      Begin iProfessionalLibrary.iSevenSegmentClockSMPTEX Reloj 
         Height          =   600
         Left            =   8955
         TabIndex        =   4
         Top             =   -15
         Visible         =   0   'False
         Width           =   2370
         Time            =   0
         FieldNumber     =   0
         FrameNumber     =   0
         ShowFieldNumber =   0   'False
         ShowFrameNumber =   0   'False
         HourStyle       =   1
         AutoSize        =   0   'False
         DigitSpacing    =   6
         SegmentMargin   =   5
         SegmentColor    =   16777215
         SegmentSeperation=   1
         SegmentSize     =   1
         ShowOffSegments =   0   'False
         PowerOff        =   0   'False
         BackGroundColor =   0
         BorderStyle     =   0
         FrameStyle      =   1
         Hours           =   0
         Minutes         =   0
         Seconds         =   0
         Enabled         =   -1  'True
         SegmentOffColor =   16777215
         AutoSegmentOffColor=   0   'False
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Visible         =   -1  'True
         Object.Width           =   158
         Object.Height          =   39
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iLedRoundX iLedVacio1 
         Height          =   285
         Left            =   1380
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   285
         BackGroundColor =   -2147483633
         Active          =   -1  'True
         ActiveColor     =   255
         BevelStyle      =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         ShowReflection  =   -1  'True
         AutoInactiveColor=   -1  'True
         InactiveColor   =   8421440
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   19
         Object.Height          =   19
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelVacio1 
         Height          =   375
         Left            =   1740
         TabIndex        =   2
         Top             =   120
         Width           =   1560
         AutoSize        =   -1  'True
         Alignment       =   1
         BorderStyle     =   0
         Caption         =   "Alarma Vacío"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   0   'False
         ShadowXOffset   =   1
         ShadowYOffset   =   1
         ShadowColor     =   8421504
         BackGroundColor =   -2147483633
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   0
         Transparent     =   0   'False
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   104
         Object.Height          =   25
         OPCItemCount    =   0
      End
      Begin isAnalogLibrary.iLabelX iLabelX4 
         Height          =   0
         Left            =   4740
         TabIndex        =   1
         Top             =   270
         Width           =   0
         AutoSize        =   -1  'True
         Alignment       =   1
         BorderStyle     =   0
         Caption         =   "iLabel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterMarginLeft =   0
         OuterMarginTop  =   0
         OuterMarginRight=   0
         OuterMarginBottom=   0
         ShadowShow      =   0   'False
         ShadowXOffset   =   1
         ShadowYOffset   =   1
         ShadowColor     =   0
         BackGroundColor =   -2147483633
         UpdateFrameRate =   60
         Object.Visible         =   -1  'True
         FontColor       =   0
         Transparent     =   0   'False
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Enabled         =   -1  'True
         Object.Width           =   0
         Object.Height          =   0
         OPCItemCount    =   0
      End
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   75
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuAnalisis 
      Caption         =   "&Analisis"
      Visible         =   0   'False
      Begin VB.Menu mnuAnalisisActual 
         Caption         =   "Analisis ac&tual"
      End
      Begin VB.Menu mnuAnalisisPozos 
         Caption         =   "&Pozos"
      End
   End
   Begin VB.Menu mnuCaminos 
      Caption         =   "Ca&minos"
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "&Informes"
      Visible         =   0   'False
      Begin VB.Menu mnuInformeDeAnalisis 
         Caption         =   "Informe de análisis"
      End
      Begin VB.Menu mnuInformeDeComparacionAnalisis 
         Caption         =   "&Comparación de análisis"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "&Configuración"
      Begin VB.Menu mnuConfiguracionGral 
         Caption         =   "C&onfiguración General"
      End
      Begin VB.Menu mnuZonasMuertas 
         Caption         =   "&Zona adicional"
         Begin VB.Menu mnuZonaMuerta 
            Caption         =   "Sin zona"
            Index           =   0
         End
         Begin VB.Menu mnuConfZonaAdic 
            Caption         =   "Configurar"
            Index           =   5
         End
      End
      Begin VB.Menu mnuSimulacion 
         Caption         =   "Simulación"
      End
   End
   Begin VB.Menu MnuGenerarReporte 
      Caption         =   "Generar Reporte"
   End
   Begin VB.Menu MnuInterpolar 
      Caption         =   "Interpolar"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Form"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
   Begin VB.Menu mnuPopupTemporal 
      Caption         =   "PopupTemporal"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupTemporalSeleccionar 
         Caption         =   "&Seleccionar"
      End
      Begin VB.Menu mnuPopupTemporalAgregar 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu mnuPopupTemporalModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnuPopupTemporalBorrar 
         Caption         =   "&Borrar"
      End
      Begin VB.Menu mnuPopupTemporalSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupTemporalCalibracion 
         Caption         =   "&Calibración"
      End
      Begin VB.Menu mnuPopupTemporalGasDeYacimiento 
         Caption         =   "&Gas de yacimiento"
      End
      Begin VB.Menu mnuPopupTemporalCirculada 
         Caption         =   "&Circulada"
      End
   End
   Begin VB.Menu mnuPopupDefinitivo 
      Caption         =   "PopupDefinitivo"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupDefinitivoModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnuPopupDefinitivoBorrar 
         Caption         =   "&Borrar"
      End
   End
   Begin VB.Menu mnuPopupPozos 
      Caption         =   "PopupPozos"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupPozosAgregar 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu mnuPopupPozosModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnuPopupPozosBorrar 
         Caption         =   "&Borrar"
      End
   End
   Begin VB.Menu mnuGrafico 
      Caption         =   "Menu grafico"
      Visible         =   0   'False
      Begin VB.Menu mnuEscalas 
         Caption         =   "Escala"
         Begin VB.Menu mnuEscala 
            Caption         =   "10.000 ppm"
            Index           =   10
         End
         Begin VB.Menu mnuEscala 
            Caption         =   "20.000 ppm"
            Index           =   20
         End
         Begin VB.Menu mnuEscala 
            Caption         =   "50.000 ppm"
            Index           =   50
         End
         Begin VB.Menu mnuEscala 
            Caption         =   "100.000 ppm"
            Index           =   100
         End
         Begin VB.Menu mnuEscala 
            Caption         =   "200.000 ppm"
            Index           =   200
         End
         Begin VB.Menu mnuEscala 
            Caption         =   "500.000 ppm"
            Index           =   500
         End
         Begin VB.Menu mnuEscala 
            Caption         =   "1.000.000 ppm"
            Index           =   1000
         End
      End
      Begin VB.Menu mnuSpanes 
         Caption         =   "Intervalos temporales"
         Begin VB.Menu mnuSpan 
            Caption         =   "15 minutos"
            Index           =   15
         End
         Begin VB.Menu mnuSpan 
            Caption         =   "30 minutos"
            Index           =   30
         End
         Begin VB.Menu mnuSpan 
            Caption         =   "1 hora"
            Index           =   60
         End
         Begin VB.Menu mnuSpan 
            Caption         =   "2 horas"
            Index           =   120
         End
      End
   End
   Begin VB.Menu mnuEscalasLog 
      Caption         =   "Escala"
      Visible         =   0   'False
      Begin VB.Menu mnuEscalaLog 
         Caption         =   "10.000 ppm"
         Index           =   10
      End
      Begin VB.Menu mnuEscalaLog 
         Caption         =   "500.000 ppm"
         Index           =   500
      End
      Begin VB.Menu mnuEscalaLog 
         Caption         =   "1.000.000 ppm"
         Index           =   1000
      End
   End
End
Attribute VB_Name = "MdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents pObjAdminEventos                 As clsAdminEventos
Attribute pObjAdminEventos.VB_VarHelpID = -1
Private Intervalo                                   As Long

Public FechaHoraUltimaGrabacion                     As Date

Private pFechaHoraUltimoMensajeRecivido             As Date
Public ServerName                                   As String
Const PARAM_LEN = 10
Dim Counter                                         As Double

Dim LineaLeida                                      As String

Dim ErrorLecturaDeGas                               As Boolean
Dim ErrorEscrituraDeGas                             As Boolean
Dim ErrorLeerRetYGases                              As Boolean
Dim ErrorEscrituraRetYGases                         As Boolean
Dim ErrorEscrituraDisparo                           As Boolean
Dim ErrorLecturaDisparo                             As Boolean

Private pCantidadEscaneos                           As Long


Private Sub cmdIniciarAnalisis_Click()

    DispararCroma

End Sub

Public Sub JustTimerCroma()

    TimerCroma_OnTimer2

End Sub



Private Sub MDIForm_Load()
    
    Set pObjAdminEventos = gObjAdminEventos
    Counter = 1
    Intervalo = 1
    Ticks = 0
    TicksAnterior = 0
    DeboTirarCroma = False
    CromaLanzada = False
    EFastLocal = 1
    ENormalLocal = 1
    EContinuoLocal = 1
    ObtenerEscalasRabbit
    LimpiarResultados
    InicializarClienteBaselineGasTotal
    FrmAnalisis.Show

End Sub


Public Sub IniGralComm()
    
    EFastLocal = 1
    ENormalLocal = 1
    EContinuoLocal = 1
    ObtenerEscalasRabbit
    InicializarClienteBaselineGasTotal

End Sub





Public Sub ObtenerEscalasRabbit()
    
    Dim cadenaRecibida      As String
    Dim nHeaderCount        As Integer
    Dim strHeader()         As String
    Dim NombreHost          As String
    Dim cadenaAEnviar       As String
    Dim Valores()            As String
    
    If gObjConfiguracion.IdTipoEquipoCroma = 2 Or gObjConfiguracion.IdTipoEquipoCroma = 3 Then
       NombreHost = gObjConfiguracion.IpRabit
       cadenaAEnviar = "cadena=" & cadenaAEnviar
       nHeaderCount = 5
       
       ReDim strHeader(nHeaderCount)
       strHeader(1) = "POST /getDatosFID.cgi HTTP/1.0" & Chr(13) & Chr(10)
       strHeader(2) = "Content-type: application/x-www-form-urlencoded" & Chr(13) & Chr(10)
       strHeader(3) = "Content-length: " & Len(cadenaAEnviar) & Chr(13) & Chr(10)
       strHeader(4) = Chr(13) & Chr(10)
       strHeader(5) = cadenaAEnviar & Chr(13) & Chr(10)
       
       If comunicar(strHeader, nHeaderCount, cadenaRecibida, Sw, NombreHost) Then
            'cadenaRecibida = "DC01:|24ld|3.43ld|145ld|10000.00ld|2919.00ld|414.53ld|-19ld| 37.45ld|145ld|10000.00ld|267.00ld|6142.32ld|-40ld|476.19ld|145ld|10000.00ld|21.00ld|88095.24ld|40ld|10.00ld|-1ld|1000.00ld|100.00ld|0.00ld|50ld|0.10ld|-1ld|1000.00ld| 10000.00ld|0.00ld|0ld|0.00ld|-1ld|14.50ld|10000.00ld|-0.00ld|0ld|0.00ld|14ld|14.50ld|10000.00ld|0.02ld|10ld|100ld|1ld|OFFld|OFFld|OFFld|RUNNINGld|STAND BYld|6142ld|1ld|1ld|100ld|60ld|0.0ld|-0.2ld|10ld|10ld|0ld|12ld|1ld|1ld|POS?ld|POS?ld|_DC01;"
            Valores = Split(cadenaRecibida, "ld|")
            
            EContinuoLocal = Valores(42)
            ENormalLocal = Valores(43)
            EFastLocal = Valores(44)
            
       End If
       
       
    End If
End Sub

Private Function comunicar(httpHeader() As String, httpHeaderDim As Integer, ByRef cadenaRecibida As String, ByVal Sw As Variant, ByVal NombreHost As String) As Boolean
    Dim nHeader As Integer, bInHeader As Boolean, nPos As Long
    Dim strBuffer As String, nResultCode As Long
    Dim cchHeader As Integer, strHeaderField As String, strHeaderValue As String
    Dim g_nLineCount As Long, g_nTotalBytes As Long
    Dim strHeaderRespuesta() As String
    Dim nHeaderCountRespuesta As Integer
'    Dim strBuffer As String
    Dim cchBuffer As Integer
    Dim strHeaderBuffer As String
    Dim cchHeaderBuffer As Integer
    
    comunicar = False

Restart:
    ReDim g_strLine(0)
    g_nLineCount = 0
    g_nTotalBytes = 0
    
    
    Sw.AutoResolve = False
    Sw.Blocking = True

    '
    ' Attempt the connection to the server
    '
    If Sw.Connect(NombreHost, 80, swProtocolTcp, 2) = 0 _
        Then
            For nHeader = 1 To httpHeaderDim
            '
            ' If the number of bytes written doesn't match the length
            ' of the header string, then something has gone wrong;
            ' for a blocking socket, these values should be the same
            '
                cchHeader = Len(httpHeader(nHeader))
                If Sw.Write(httpHeader(nHeader), cchHeader) <> cchHeader Then
                    Sw.Disconnect
                    Exit Function
                End If
            
            DoEvents
            
            Next
            
            '
            ' The server will reply with a response header block, followed
            ' by the data for the requested resource; we will re-use the
            ' strHeader array to contain the response header values, and
            ' store each line of the resource in the g_strLine array
            '
            ' Note that this sample, as written, expects that only textual
            ' data (such as HTML pages) will be returned by the server
            '
            ReDim strHeaderRespuesta(0)
            
            nHeaderCountRespuesta = 0
            cchHeaderBuffer = 0
            cchBuffer = 0
            bInHeader = True
            Do
                '
                ' Read the data from the socket, and store it in strBuffer;
                ' the actual number of bytes read is stored in cchBuffer
                '
                    DoEvents
                    
                    cchBuffer = Sw.Read(strBuffer, 2048)
                    If cchBuffer = 0 Then
                        '
                        ' The server has closed the connection and we have
                        ' reached the end of the data stream
                        '
                        Exit Do
                    ElseIf cchBuffer = -1 Then
                        '
                        ' An error has occurred while reading data from the
                        ' server; this should be considered a fatal error
                        '
                        Sw.Disconnect
                        Exit Function
                    End If
            
                    If bInHeader Then
                        '
                        ' If we are processing the response header block, then
                        ' store the data into the header buffer
                        '
                        strHeaderBuffer = strHeaderBuffer + strBuffer
                        cchHeaderBuffer = cchHeaderBuffer + cchBuffer
            
                        '
                        ' Look for the end of the header block, which is a
                        ' blank line (a pair of CRLF sequences)
                        '
                        nPos = InStr(strHeaderBuffer, Chr(13) & Chr(10) & Chr(13) & Chr(10))
                        If nPos > 0 Then
                            '
                            ' The end of the header block has been reached; the
                            ' entire response header is stored in strHeaderBuffer
                            ' and the remaining data is left in strBuffer to be
                            ' processed later
                            '
                            cchBuffer = cchBuffer - (nPos + 3)
                            strBuffer = Right(strHeaderBuffer, cchBuffer)
                            strHeaderBuffer = Left(strHeaderBuffer, nPos + 1)
                            bInHeader = False
            
                            '
                            ' Break strHeaderBuffer apart, with each response
                            ' header field being placed into the strHeader array;
                            ' this will make it simple to search for specific
                            ' header values, etc.
                            '
                            Do
                                nPos = InStr(strHeaderBuffer, Chr(10))
                                If nPos = 0 Then
                                    Exit Do
                                Else
                                    nHeaderCountRespuesta = nHeaderCountRespuesta + 1
                                    ReDim Preserve strHeaderRespuesta(nHeaderCountRespuesta)
                                    strHeaderRespuesta(nHeaderCountRespuesta) = Trim(Left(strHeaderBuffer, nPos - 2))
                                    strHeaderBuffer = Right(strHeaderBuffer, Len(strHeaderBuffer) - nPos)
                                End If
                            
                            DoEvents
                            
                            Loop
            
                            '
                            ' Note that strHeader(1) will contain the command
                            ' response from the server, and will typically look
                            ' something like this:
                            '
                            '       HTTP/1.0 200 OK
                            '
                            ' The first part contains the protocol version (in
                            ' this case 1.0), the second is the result code
                            ' and what follows is a textual description of the
                            ' result. A result code in the range of 200-299
                            ' indicates success; for a complete description of
                            ' the result codes, refer to RFC 2616
                            '
                            nPos = InStr(strHeaderRespuesta(1), " ")
                            If nPos > 0 Then
                                nResultCode = Val(Right(strHeaderRespuesta(1), Len(strHeaderRespuesta(1)) - nPos))
                            End If
            
                            If nResultCode >= 300 And nResultCode <= 303 Then
                                '
                                ' A result code in this range indicates that the
                                ' resource has been moved; the new location is
                                ' specified in the Location header field
                                '
                                For nHeader = 2 To nHeaderCountRespuesta
                                    nPos = InStr(strHeaderRespuesta(nHeader), ":")
                                    If nPos > 0 Then
                                        strHeaderField = UCase(Left(strHeaderRespuesta(nHeader), nPos - 1))
                                        If strHeaderField = "LOCATION" Then
                                            Sw.Disconnect
                                            GoTo Restart
                                        End If
                                    End If
                                Next
                            End If
            
                            If nResultCode < 200 Or nResultCode > 299 Then
                                Sw.Disconnect
                                Exit Function
                            End If
            
                            '
                            ' Determine the content type of the data being returned
                            '
                            For nHeader = 2 To nHeaderCountRespuesta
                            
                            DoEvents
                                
                                nPos = InStr(strHeaderRespuesta(nHeader), ":")
                                If nPos > 0 Then
                                    strHeaderField = UCase(Left(strHeaderRespuesta(nHeader), nPos - 1))
                                    
                                    If strHeaderField = "CONTENT-TYPE" Then
                                        '
                                        ' If the content type is not textual, then disconnect
                                        ' and warn the user that we cannot display it
                                        '
                                        strHeaderValue = Trim(Right(strHeaderRespuesta(nHeader), Len(strHeaderRespuesta(nHeader)) - nPos))
                                        If Left(strHeaderValue, 5) <> "text/" Then
                                            Sw.Disconnect
                                            Exit Function
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next
                            '
                            ' Any additional checks for specific header field values
                            ' could be placed here
                            '
                        End If
                    End If
                    '
                    ' If we are not processing the header block, the data into
                    ' individual lines to make it easier to process; this will
                    ' also handle the different end-of-line character sequences
                    ' used by UNIX and Windows servers
                    '
                    If Not bInHeader Then
                        g_nTotalBytes = g_nTotalBytes + cchBuffer
                        If g_nLineCount = 0 Then
                            g_nLineCount = 1
                            ReDim Preserve g_strLine(g_nLineCount)
                        End If
                        '
                        ' If the buffer contains carriage returns, then strip
                        ' them out and use only linefeeds to mark the end-of-line
                        '
                        Do
                            nPos = InStr(strBuffer, Chr(13))
                            If nPos = 0 Then Exit Do
                            strBuffer = Left(strBuffer, nPos - 1) & Right(strBuffer, cchBuffer - nPos)
                            cchBuffer = cchBuffer - 1
                        Loop
                        Do
                            nPos = InStr(strBuffer, Chr(10))
                            If nPos = 0 Then
                                Exit Do
                            Else
                                '
                                ' If the linefeed is at the beginning of the line, then
                                ' simply append CRLF; otherwise append the remaining
                                ' characters and then CRLF.
                                '
                                If nPos = 1 Then
                                    g_strLine(g_nLineCount) = g_strLine(g_nLineCount) & Chr(13) & Chr(10)
                                Else
                                    g_strLine(g_nLineCount) = g_strLine(g_nLineCount) & Left(strBuffer, nPos - 1) & Chr(13) & Chr(10)
                                End If
                                g_nLineCount = g_nLineCount + 1
                                ReDim Preserve g_strLine(g_nLineCount)
                                cchBuffer = cchBuffer - nPos
                                strBuffer = Right(strBuffer, cchBuffer)
                            End If
                        
                        DoEvents
                        
                        Loop
                        If cchBuffer > 0 Then
                            g_strLine(g_nLineCount) = g_strLine(g_nLineCount) + strBuffer
                        End If
                    End If
                    
                    DoEvents
                    
                    cadenaRecibida = strBuffer
                Loop
                Sw.Disconnect
                comunicar = True
            End If
End Function


Private Sub InicializarClienteBaselineGasTotal()
    
    ' El nombre del control Winsock es tcpClient.
    ' Nota: para especificar un host remoto, puede usar
    ' la dirección IP (como "121.111.1.1") o
    ' el nombre "descriptivo" del equipo, como se muestra aquí.
    tcpClient.Close
    tcpClient.Protocol = sckTCPProtocol
    tcpClient.RemoteHost = "192.168.0.230"
    tcpClient.RemotePort = 57344
    ConectarClienteBaselineGasTotal
End Sub

Private Sub ConectarClienteBaselineGasTotal()

' Invoca el método Connect para iniciar
' una conexión.
On Error GoTo Error

    tcpClient.Connect

Error:
        If Err.Number <> 0 Then
            Err.Clear
        End If

End Sub

Private Sub DesconectarClienteBaselineGasTotal()

' Invoca el método Connect para iniciar
' una conexión.
On Error GoTo Error

    tcpClient.Close

Error:
        If Err.Number <> 0 Then
            Err.Clear
        End If

End Sub

Private Sub EnviarDatosClienteBaselineGasTotal(ByVal strEnviar As String)

    tcpClient.SendData strEnviar

End Sub

Public Sub LimpiarResultados()

    
    If UCase(Dir(gCaminoArchivos & gObjConfiguracion.ArchivoDeResultados)) <> "" Then
        Kill gCaminoArchivos & gObjConfiguracion.ArchivoDeResultados
    End If
    
    If UCase(Dir(gCaminoArchivos & "resultados_" & Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & ".txt")) <> "" Then
        Kill gCaminoArchivos & "resultados_" & Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & ".txt"

    End If
    
End Sub

Private Sub MDIForm_Resize()
    Call modAutoResize.ResizeForm(Me)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)


    Set FrmAnalisis = Nothing
    Set frmCaminos = Nothing
    Set FrmConfiguracion = Nothing
    Set frmConfigurarZonaAdic = Nothing
    Set FrmDatosComponente = Nothing
    Set FrmInformeAnalisis = Nothing
    Set frmMsg = Nothing
    Set frmSpanCrono = Nothing
    Set frmSplash = Nothing
    
    TimerEntradas.Enabled = False
    TimerSalidas.Enabled = False
    TimerResultados.Enabled = False
    TimerRefresh.Enabled = False
    TimerSimulador.Enabled = False
    TimerCroma.Enabled1 = False
    TimerCroma.Enabled2 = False
    TimerCroma.Enabled3 = False
    TimerCroma.Enabled4 = False
    TimerCroma.Enabled5 = False
    TimerCroma.Enabled6 = False
    TimerCroma.Enabled7 = False
    TimerCroma.Enabled8 = False
    TimerCroma.Enabled9 = False
    
    Set MdiPrincipal = Nothing


End Sub
Private Sub mnuAnalisisActual_Click()
    FrmAnalisis.Show
End Sub
Private Sub mnuCaminos_Click()
    frmCaminos.Show
End Sub
Private Sub mnuConfiguracionGral_Click()
    FrmConfiguracion.Show
    FrmConfiguracion.SetFocus
End Sub
Private Sub mnuConfZonaAdic_Click(Index As Integer)
    frmConfigurarZonaAdic.Show
    frmConfigurarZonaAdic.SetFocus
End Sub

Private Sub mnuEscala_Click(Index As Integer)
    SetearMenuEscala Index
End Sub

Private Sub mnuEscalaLog_Click(Index As Integer)
    SetearMenuEscalaLog Index
End Sub

Private Sub MnuGenerarReporte_Click()
    On Error GoTo Error

    Screen.MousePointer = 13

    GenerarTxt
    GenerarTxtPorc
    GenerarTxtCronoRop
    GenerarTxtGTC

    Screen.MousePointer = 0

Error:
        If Err.Number <> 0 Then
            frmMsg.MostrarMsg "Error al generar el reporte.", "Error", MdiPrincipal
            Err.Clear
        End If
End Sub
Private Sub mnuInformeDeAnalisis_Click()
    FrmInformeAnalisis.Show
    FrmInformeAnalisis.SetFocus
End Sub
Private Sub MnuInterpolar_Click()
    FrmAnalisis.Interpolar
End Sub
Private Sub mnuPopupDefinitivoBorrar_Click()
    ActiveForm.txtPopupElegido.Text = "BORRAR"
End Sub
Private Sub mnuPopupDefinitivoModificar_Click()
    ActiveForm.txtPopupElegido.Text = "MODIFICAR"
End Sub
Private Sub mnuPopupPozosAgregar_Click()
    ActiveForm.txtPopupElegido.Text = "AGREGAR"
End Sub
Private Sub mnuPopupPozosBorrar_Click()
    ActiveForm.txtPopupElegido.Text = "BORRAR"
End Sub
Private Sub mnuPopupPozosModificar_Click()
    ActiveForm.txtPopupElegido.Text = "MODIFICAR"
End Sub
Private Sub mnuPopupTemporalAgregar_Click()
    ActiveForm.txtPopupElegido.Text = "AGREGAR"
End Sub
Private Sub mnuPopupTemporalBorrar_Click()
    ActiveForm.txtPopupElegido.Text = "BORRAR"
End Sub
Private Sub mnuPopupTemporalCalibracion_Click()
    ActiveForm.txtPopupElegido.Text = "CALIBRACION"
End Sub
Private Sub mnuPopupTemporalCirculada_Click()
    ActiveForm.txtPopupElegido.Text = "CIRCULADA"
End Sub
Private Sub mnuPopupTemporalGasDeYacimiento_Click()
    ActiveForm.txtPopupElegido.Text = "GAS DE YACIMIENTO"
End Sub
Private Sub mnuPopupTemporalModificar_Click()
    ActiveForm.txtPopupElegido.Text = "MODIFICAR"
End Sub
Private Sub mnuPopupTemporalSeleccionar_Click()
    ActiveForm.txtPopupElegido.Text = "SELECCIONAR"
End Sub
Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuSimulacion_Click()

    mnuSimulacion.Checked = Not mnuSimulacion.Checked
    TimerSimulador.Enabled = mnuSimulacion.Checked
    
End Sub

Private Sub mnuSpan_Click(Index As Integer)
    SetearMenuSpanTemporal Index
End Sub



Private Sub mnuZonaMuerta_Click(Index As Integer)
    SetearMenuZonaMuerta Index
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)

Dim strData As String
Dim tArray() As String

    tcpClient.GetData strData, vbString
    tArray = Split(strData, Chr(9))
    'Txtfecha = tArray(0)
    'Txthora = tArray(1)
    If tArray(3) = "ppm" Then
        GasTotal = tArray(2)
    Else
        GasTotal = tArray(2) * 10000
    End If
    
    'TxtUnidad = tArray(3)

End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Err.Number <> 0 Then
        Debug.Print "Error"
    End If
End Sub

Private Sub TimerRefresh_Timer()
    
    LeerComunicacionDC_DT
    If DateDiff("s", HoraInicio, Now()) < gObjConfiguracion.TiempoAnalisis Then Reloj.Time = Now() - HoraInicio

    TxtProfundidad.Text = ProfundidadRetorno

    LedAnalisisAutomaticos.Active = gObjConfiguracion.AnalisisAutomaticos

    If DeboTirarCroma And Not TimerCroma.Enabled1 And Not CromaLanzada And gObjConfiguracion.AnalisisAutomaticos Then
        DispararCroma
    End If

    If DeboAgregarCartel Then
        DeboAgregarCartel = False
        AgregarProfundidadRetorno
        If GasTotal >= gObjConfiguracion.NivelDisparoMetroMetro Then
            If Not CromaLanzada Then
                DeboTirarCroma = True
            End If
        End If
    End If

    If FormularioCargado("FrmAnalisis") Then
        FrmAnalisis.iPlotGases.Limit(0).Line1Position = gObjConfiguracion.NivelDisparoMetroMetro
    End If

    If EstadoAnterior <> Estado And (Estado = "Datos") Then
        ComentarioGas = Estado
    End If

    EstadoAnterior = Estado
    If TimerCroma.Enabled3 = False Then FrmAnalisis.ActualizoDisplayEscalas
    If gObjConfiguracion.CambioEscalas Then AnalizarCambiosEscala
    
    Intervalo = Intervalo + 1
    pObjAdminEventos.ActualizarVistaTiempo Intervalo
    AnalizarRegistroTiempo
    pObjAdminEventos.ActualizarValoresGases

    EscribirGases GasTotal, CO2, SH2
    
End Sub

Private Sub TimerCroma_OnTimer1()

    MdiPrincipal.iLedStatusCroma.Active = False
    'no
    cmdIniciarAnalisis.Enabled = True
    TimerCroma.Enabled1 = False
    Reloj.Visible = False
    CromaLanzada = False

End Sub

'Este timer pretende relanzar el scan cada 5 segundos

Private Sub TimerCroma_OnTimer2()
    
    'Leo Retorno
    LeerComunicacionDC_DT


    TasSerial1.ScanActive = False
    TasSerial2.ScanActive = False

    TasSerial1.AbortCommunication
    TasSerial2.AbortCommunication

    TasSerial1.Wait 1000

    If gObjConfiguracion.PuertoComm = 1 Then
        TasSerial1.CommPort = COM1
        TasSerial2.CommPort = COM1
    End If

    If gObjConfiguracion.PuertoComm = 2 Then
        TasSerial1.CommPort = COM2
        TasSerial2.CommPort = COM2
    End If

    If gObjConfiguracion.PuertoComm = 3 Then
        TasSerial1.CommPort = COM3
        TasSerial2.CommPort = COM3
    End If

    If gObjConfiguracion.PuertoComm = 4 Then
        TasSerial1.CommPort = COM4
        TasSerial2.CommPort = COM4
    End If

    If gObjConfiguracion.PuertoComm = 5 Then
        TasSerial1.CommPort = COM5
        TasSerial2.CommPort = COM5
    End If

    If gObjConfiguracion.PuertoComm = 6 Then
        TasSerial1.CommPort = COM6
        TasSerial2.CommPort = COM6
    End If


    If gObjConfiguracion.PuertoComm = 7 Then
        TasSerial1.CommPort = COM7
        TasSerial2.CommPort = COM7
    End If

    If gObjConfiguracion.PuertoComm = 8 Then
        TasSerial1.CommPort = COM8
        TasSerial2.CommPort = COM8
    End If
    TasSerial1.ScanActive = True
    TasSerial2.ScanActive = True
    TimerEntradas.Enabled = False
    TasSerial1.Trigger
End Sub

Private Sub TimerCroma_OnTimer3()
    TimerCroma.Enabled3 = False
End Sub

Private Sub TimerResultados_Timer()
    
    Dim ok                      As Boolean
    Dim ArchivoSeleccionado     As String
    Dim ArchivoSeleccionado2    As String
    Dim TiempoInicio            As Date
    Dim fecha                   As Date
    Dim fso                     As New FileSystemObject
    
    ArchivoSeleccionado = ""
    ArchivoSeleccionado2 = ""
    
    If Not gObjPozoActivo Is Nothing Then
        If gObjConfiguracion.ArchivoDeResultados <> "" And gObjPozoActivo.IdPozo <> 0 Then
            ArchivoSeleccionado = ""
            
            If fso.FileExists(gCaminoArchivos & gObjConfiguracion.ArchivoDeResultados) Then
                ArchivoSeleccionado = gCaminoArchivos & gObjConfiguracion.ArchivoDeResultados
            ElseIf fso.FileExists(gCaminoArchivos & "resultados_" & Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & ".txt") Then
                ArchivoSeleccionado = gCaminoArchivos & "resultados_" & Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & ".txt"
            Else
                ArchivoSeleccionado = Dir(gCaminoArchivos & "resul*.esd")
                If (ArchivoSeleccionado <> "") Then ArchivoSeleccionado = gCaminoArchivos & ArchivoSeleccionado
                Debug.Print ArchivoSeleccionado
                If ArchivoSeleccionado <> "" Then ArchivoSeleccionado2 = Dir()
                If ArchivoSeleccionado2 <> "" Then ArchivoSeleccionado2 = gCaminoArchivos & ArchivoSeleccionado2
                Debug.Print ArchivoSeleccionado2
            End If
            
            

            
            
            
            If ArchivoSeleccionado <> "" Then
                gDatos.BeginTrans
                TiempoInicio = Now()
    
                'hago una pausa esperando que se complete el reporte de las dos lineas del archivo resultados.log
                Do
                    DoEvents
                Loop Until (DateDiff("s", TiempoInicio, Now()) >= 5)
                
                
            If gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 1 Then
                    leer_log ok, ArchivoSeleccionado
            'IMPROBABLE
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 2 Then
                    leer_log ok, ArchivoSeleccionado
            
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 3 Then
                    leer_txt ok, ArchivoSeleccionado
                
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 4 Then
                    Leer_txt_skycrhome ok, ArchivoSeleccionado
                
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 5 Then
                    Leer_txt_Varian ok, ArchivoSeleccionado, ArchivoSeleccionado2
            End If
                
                
       If gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 1 Then
                    leer_log ok, ArchivoSeleccionado
                    
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 2 Then
                    leer_log ok, ArchivoSeleccionado
            
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 3 Then
                    leer_txt ok, ArchivoSeleccionado
                
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 4 Then
                    Leer_txt_skycrhome ok, ArchivoSeleccionado
                
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 5 Then
                    Leer_txt_Varian ok, ArchivoSeleccionado, ArchivoSeleccionado2
            End If
                
                
       If gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 1 Then
                    leer_log ok, ArchivoSeleccionado
                    
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 2 Then
                    leer_log ok, ArchivoSeleccionado
            
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 3 Then
                    leer_txt ok, ArchivoSeleccionado
                
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 4 Then
                    Leer_txt_skycrhome ok, ArchivoSeleccionado
                
            ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 5 Then
                    Leer_txt_Varian ok, ArchivoSeleccionado, ArchivoSeleccionado2
            End If
                              
                              
                 
                
                
'                If gObjConfiguracion.IdTipoEquipoCroma = 1 Or gObjConfiguracion.IdTipoEquipoCroma = 2 Then
'                    leer_log ok, ArchivoSeleccionado
'                ElseIf gObjConfiguracion.IdTipoEquipoCroma = 3 Then
'                    leer_txt ok, ArchivoSeleccionado
'                ElseIf gObjConfiguracion.IdTipoEquipoCroma = 4 Then
'                    Leer_txt_skycrhome ok, ArchivoSeleccionado
'                End If
                
                
                
                If ok Then
                    gDatos.CommitTrans
                    pObjAdminEventos.ActualizarVistaProf
                    If fso.FileExists(ArchivoSeleccionado) Then fso.DeleteFile ArchivoSeleccionado
                    If fso.FileExists(ArchivoSeleccionado2) Then fso.DeleteFile ArchivoSeleccionado2
                Else
                    gDatos.RollBackTrans
                    frmMsg.MostrarMsg "Lectura de Resultados", "Error", MdiPrincipal
                End If
    
            End If
        End If
    End If
End Sub

Private Sub timerSimulador_timer()
    
    Dim ModuloGrabacion                     As Integer
    On Error GoTo Error
    
    ModuloGrabacion = 5
    
    If Not gObjConfiguracion.GasTotalDesdeDataCenter Then
        GasTotal = Int((6000 * Rnd) + 1)
    End If
    If Not gObjConfiguracion.CO2DesdeDataCenter Then
        CO2 = GasTotal / 4
    End If
    If Not gObjConfiguracion.SH2DesdeDataCenter Then
        SH2 = GasTotal / 100
    End If
Error:
    If Err.Number <> 0 Then
        MsgBox "Por ahora: error: " & Err.Description
        Err.Clear
    End If
    TimerSimulador.Enabled = True
End Sub

Private Sub TimerEntradas_Timer()

    If gObjConfiguracion.IdTipoEquipoGas = 1 Then
        If gObjConfiguracion.PuertoComm = 1 Then
            TasSerial1.CommPort = COM1
            TasSerial2.CommPort = COM1
        End If
    
        If gObjConfiguracion.PuertoComm = 2 Then
            TasSerial1.CommPort = COM2
            TasSerial2.CommPort = COM2
        End If
        If gObjConfiguracion.PuertoComm = 6 Then
            TasSerial1.CommPort = COM6
            TasSerial2.CommPort = COM6
        End If
        
        TasSerial1.Trigger
        TimerEntradas.Enabled = False
        TimerCroma.Enabled2 = True
    End If
    
End Sub
Private Sub TasSerial1_OnSuccessfullyReceived()

    Dim NedTemp                                 As Integer
    Dim ModuloGrabacion                         As Integer
    Dim StrSql                                  As String
    Dim fecha                                   As Double
    Dim Index                                   As Long
    Dim IdPozo                                  As Long
    Dim objPozoGasTiempo                        As clsPozoGasTiempo
    
    ModuloGrabacion = 5

    StartRun = TasSerial1.PointValue(0)

    BackFlushAnalizing = TasSerial1.PointValue(1)

    EContinuoControlador = TasSerial1.PointValue(2)
    ENormalControlador = TasSerial1.PointValue(3)
    EFastControlador = TasSerial1.PointValue(4)

    PosFast = TasSerial1.PointValue(5)
    PosNormal = TasSerial1.PointValue(6)

    TempFID = TasSerial1.PointValue(7)
    TempOVEN = TasSerial1.PointValue(8)

    SamplePressure = TasSerial1.PointValue(9)

    Signal = TasSerial1.PointValue(10)

    If Not gObjConfiguracion.SH2DesdeDataCenter Then
        SH2 = TasSerial1.PointValue(11)
    End If

    If Not gObjConfiguracion.CO2DesdeDataCenter Then
        CO2 = TasSerial1.PointValue(12)
    End If

    If Not gObjConfiguracion.GasTotalDesdeDataCenter Then
        GasTotal = TasSerial1.PointValue(14) * 1000 + TasSerial1.PointValue(13) '* EContinuoLocal
    End If
    
    pObjAdminEventos.ActualizarValoresGases
    
    If TasSerial1.PointValue(17) = 0 _
        Then
            'FrmAnalisis.iLed.Active = False
            AlarmaVacio = 0
            iLedVacio1.Visible = False
            iLabelVacio1.Visible = False
            'iLabelVacio2.Visible = False

        Else
            'FrmAnalisis.iLed.Active = True
            iLedVacio1.Visible = True
            iLabelVacio1.Visible = True
            'iLabelVacio2.Visible = True

            Call sndPlaySound(App.Path & "\Sonidos\" & "vacio.wav", SND_ASYNC)
            Call sndPlaySound(ByVal "", 0)

            AlarmaVacio = 1000
        End If


    TimerSalidas.Enabled = True

    TasSerial2.Trigger

    If Intervalo < 200 Then
        Intervalo = Intervalo + 1
        Else
        Intervalo = 0
    End If

    
    
    Intervalo = Intervalo + 1

    If Not EstoyCargando Then
        If Intervalo Mod ModuloGrabacion = 0 And DateDiff("s", FechaHoraUltimaGrabacion, Now()) > 1 Then
            Set objPozoGasTiempo = New clsPozoGasTiempo
            objPozoGasTiempo.CO2 = CO2
            objPozoGasTiempo.Comentario = ComentarioGas
            objPozoGasTiempo.fecha = Now
            objPozoGasTiempo.GasTotal = GasTotal
            objPozoGasTiempo.IdPozo = gObjPozoActivo.IdPozo
            objPozoGasTiempo.SH2 = SH2
            
            Set gObjPozoActivo.PozoGasTiempos.Datos = gDatos
            gObjPozoActivo.PozoGasTiempos.dbAgregar objPozoGasTiempo
            FechaHoraUltimaGrabacion = objPozoGasTiempo.fecha
            
        End If
        If FrmAnalisis.ChkTiempoReal = 1 Then
            If Intervalo > 200 Then
                pObjAdminEventos.ActualizarVistaTiempo Intervalo
                Intervalo = 0
            End If
        End If
    End If

End Sub

Private Sub TimerSalidas_Timer()
    
    TimerSalidas.Enabled = False
    If gObjConfiguracion.IdTipoEquipoGas = 1 Then
        TasSerial2.PointValue(0) = TasSerial2.PointValue(0) And 64513
        If EContinuoLocal = 1 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 2
        End If
        If EContinuoLocal = 10 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 4
        End If
        If EContinuoLocal = 100 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 8
        End If
        If ENormalLocal = 1 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 16
        End If
        If ENormalLocal = 10 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 32
        End If
        If ENormalLocal = 100 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 64
        End If
        If EFastLocal = 1 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 128
        End If
        If EFastLocal = 10 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 256
        End If
        If EFastLocal = 100 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 512
        End If
        If BackFlushAnalizing = 1 Then
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) And 65534
        Else
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0)
        End If
        TasSerial2.Trigger
    End If
End Sub

Public Sub Salida_rabbit(ByVal escala As Long)
        
    On Error GoTo Error
    
    While (InetDisparoCroma.StillExecuting)
    'this while loop goes forever
        DoEvents
    Wend
    If escala = 1 Then
        If EContinuoLocal = 1 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCETHAXUNO.cgi"
        ElseIf EContinuoLocal = 10 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCETHAXDIEZ.cgi"
        ElseIf EContinuoLocal = 100 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCETHAXCIEN.cgi"
        End If
    End If
    If escala = 2 Then
        If ENormalLocal = 1 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCENormXUNO.cgi"
        ElseIf ENormalLocal = 10 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCENormXDIEZ.cgi"
        ElseIf ENormalLocal = 100 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCENormXCIEN.cgi"
        End If
    End If
        
    If escala = 3 Then
        If EFastLocal = 1 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCEFastXUNO.cgi"
        ElseIf EFastLocal = 10 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCEFastXDIEZ.cgi"
        ElseIf EFastLocal = 100 Then
            InetDisparoCroma.Execute gObjConfiguracion.IpRabit & "/getCEFastXCIEN.cgi"
        End If
    End If

Error:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Err.Clear
    End If

End Sub

Private Sub TasSerial2_OnSuccessfullySent()
    If TimerEntradas.Enabled = False Then
        ResetStart
        TimerEntradas.Enabled = True
    End If
    TimerCroma.Enabled2 = False

End Sub

Public Sub ResetStart()

    If (TasSerial1.PointValue(0) And 1 = 1) Then
        TasSerial2.PointValue(0) = TasSerial2.PointValue(0) And 65534
    End If

End Sub

Public Sub DispararCroma()

    If gObjConfiguracion.IdTipoEquipoCroma = 1 Then
        
        If Not TimerCroma.Enabled1 Then
            cmdIniciarAnalisis.Enabled = False
            TasSerial2.PointValue(0) = TasSerial2.PointValue(0) Or 1
            HoraInicio = Now()
            MdiPrincipal.iLedStatusCroma.Active = True
            TimerCroma.Enabled1 = True
            Reloj.Visible = True
            TimerCroma.Interval1 = gObjConfiguracion.TiempoAnalisis * 1000
            CromaLanzada = True
            ProfundidadAnalisis = ProfundidadRetorno
            DeboTirarCroma = False
            'Agregado el 07/01/2010
            gDatos.BeginTrans
            If Falsear_disparo_analisis Then
                gDatos.CommitTrans
                LlegoMuestra = False
            Else
                gDatos.RollBackTrans
            End If
        End If
        
    Else
        If Not TimerCroma.Enabled1 Then
            'Lanzar desde inet
            cmdIniciarAnalisis.Enabled = False
            InetDisparoCroma.Execute gObjConfiguracion.IpDisparoCroma & "/getStartCroma.cgi", "GET", "/getStartCroma.cgi", vbCrLf
            HoraInicio = Now()
            MdiPrincipal.iLedStatusCroma.Active = True
            Reloj.Visible = True
            TimerCroma.Enabled1 = True
            TimerCroma.Interval1 = gObjConfiguracion.TiempoAnalisis * 1000
            CromaLanzada = True
            ProfundidadAnalisis = ProfundidadRetorno
            If DeboTirarCroma Then
                DeboTirarCroma = False
            End If
            'Agregado el 07/01/2010
            gDatos.BeginTrans
            If Falsear_disparo_analisis Then
                gDatos.CommitTrans
                LlegoMuestra = False
            Else
                gDatos.RollBackTrans
            End If
        End If
    End If
    
End Sub


Private Sub TxtProfundidad_KeyPress(KeyAscii As Integer)
    isDbl KeyAscii, TxtProfundidad.Text
End Sub

Public Sub AgregarProfundidadRetorno()

    Dim Index As Long


    Index = FrmAnalisis.iPlotGases.AddAnnotation
    FrmAnalisis.iPlotGases.Annotation(Index).Font.Size = 12
    FrmAnalisis.iPlotGases.Annotation(Index).Reference = iprtChannel
    FrmAnalisis.iPlotGases.Annotation(Index).ChannelName = FrmAnalisis.iPlotGases.Channel(0).Name


    FrmAnalisis.iPlotGases.Annotation(Index).Y = 13000 'Center X Coordinate
    FrmAnalisis.iPlotGases.Annotation(Index).X = Time + Date 'Center Y Coordinate
    FrmAnalisis.iPlotGases.Annotation(Index).Style = ipasText 'Text Annotation
    FrmAnalisis.iPlotGases.Annotation(Index).FontColor = vbWhite 'White Font
    FrmAnalisis.iPlotGases.Annotation(Index).Text = ProfundidadRetorno
    FrmAnalisis.iPlotGases.Annotation(Index).TextRotation = ira000   'Rotate up-side-down


    ComentarioGas = Format(ProfundidadRetorno, "0.00")
    
End Sub

Public Sub GenerarTxtGTC()

Dim objPozoAnalis                 As clsPozoAnalis
Dim objPozoAnalisComponente As ClsPozoAnalisComponente
Dim Profundidad As String
Dim GTC As String
Dim C1 As String
Dim C2 As String
Dim C3 As String
Dim NC4 As String
Dim IC4 As String
Dim NC5 As String
Dim IC5 As String
Dim LineaTexto As String
Dim Contador As Long
Dim strco2  As String

Me.MousePointer = 13

Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
gObjPozoActivo.PozoAnalisis.ConsultarPozoAnalisis gObjPozoActivo.IdPozo

If UCase(Dir(App.Path & "\Cromatografias_gtc.txt")) <> "" _
    Then
        Kill App.Path & "\Cromatografias_gtc.txt"
    Else
End If

Open App.Path & "\Cromatografias_gtc.txt" For Append As #2   ' Abre el archivo.

        LineaTexto = "Profundidad" & "," & "C1" & "," & "C2" & "," & "C3" & "," & "IC4" & "," & "NC4" & "," & "IC5" & "," & "NC5" & "," & "GTC" & "," & "CO2"
        Print #2, LineaTexto

Contador = 0

    For Each objPozoAnalis In gObjPozoActivo.PozoAnalisis
        If objPozoAnalis.Seleccionado Then
            Contador = Contador + 1
    
            Profundidad = Format(objPozoAnalis.Profundidad, "#0")
            GTC = Format(objPozoAnalis.GasTotalCromatografico, "0")
            
            C1 = "0"
            C2 = "0"
            C3 = "0"
            IC4 = "0"
            NC4 = "0"
            IC5 = "0"
            NC5 = "0"
    
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
    
                Select Case objPozoAnalisComponente.NumeroComponente
    
                    Case 1: C1 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 2: C2 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 3: C3 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 4: IC4 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 5: NC4 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 6: IC5 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 7: NC5 = Format(objPozoAnalisComponente.Externo, "#0")
    
                End Select
    
            Next
    
            strco2 = Format(objPozoAnalis.CO2, "#0")
            LineaTexto = Profundidad & "," & C1 & "," & C2 & "," & C3 & "," & IC4 & "," & NC4 & "," & IC5 & "," & NC5 & "," & GTC & "," & strco2
    
            Print #2, LineaTexto
        End If
    Next

Close #2



Me.MousePointer = 0

End Sub

Public Sub GenerarTxt()

Dim objPozoAnalis                 As clsPozoAnalis
Dim objPozoAnalisComponente As ClsPozoAnalisComponente
Dim Profundidad As String
Dim C1 As String
Dim C2 As String
Dim C3 As String
Dim NC4 As String
Dim IC4 As String
Dim NC5 As String
Dim IC5 As String
Dim LineaTexto As String
Dim Contador As Long
Dim strco2  As String

Me.MousePointer = 13

Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
gObjPozoActivo.PozoAnalisis.ConsultarPozoAnalisis gObjPozoActivo.IdPozo

If UCase(Dir(App.Path & "\Cromatografias.txt")) <> "" _
    Then
        Kill App.Path & "\Cromatografias.txt"
    Else
End If

Open App.Path & "\Cromatografias.txt" For Append As #2   ' Abre el archivo.

        LineaTexto = "Profundidad" & "," & "C1" & "," & "C2" & "," & "C3" & "," & "IC4" & "," & "NC4" & "," & "IC5" & "," & "NC5" & "," & "CO2"
        Print #2, LineaTexto

Contador = 0

    For Each objPozoAnalis In gObjPozoActivo.PozoAnalisis
        If objPozoAnalis.Seleccionado Then
            Contador = Contador + 1
    
            Profundidad = Format(objPozoAnalis.Profundidad, "#0")
            'objPozoAnalis.PozoAnalisComponentes.False
    
            C1 = "0"
            C2 = "0"
            C3 = "0"
            IC4 = "0"
            NC4 = "0"
            IC5 = "0"
            NC5 = "0"
    
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
    
                Select Case objPozoAnalisComponente.NumeroComponente
    
                    Case 1: C1 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 2: C2 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 3: C3 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 4: IC4 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 5: NC4 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 6: IC5 = Format(objPozoAnalisComponente.Externo, "#0")
                    Case 7: NC5 = Format(objPozoAnalisComponente.Externo, "#0")
    
                End Select
    
            Next
    
            strco2 = Format(objPozoAnalis.CO2, "#0")
            LineaTexto = Profundidad & "," & C1 & "," & C2 & "," & C3 & "," & IC4 & "," & NC4 & "," & IC5 & "," & NC5 & "," & strco2
    
            Print #2, LineaTexto
        End If
    Next

Close #2



Me.MousePointer = 0

End Sub

Public Sub GenerarTxtPorc()

Dim objPozoAnalis As New clsPozoAnalis
Dim objPozoAnalisComponente As ClsPozoAnalisComponente
Dim Profundidad As String
Dim C1 As String
Dim C2 As String
Dim C3 As String
Dim NC4 As String
Dim IC4 As String
Dim NC5 As String
Dim IC5 As String
Dim LineaTexto As String
Dim Contador As Long
Dim strco2  As String

Me.MousePointer = 13

Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
gObjPozoActivo.PozoAnalisis.ConsultarPozoAnalisis 1

If UCase(Dir(App.Path & "\CromatografiasPorc.txt")) <> "" Then
    Kill App.Path & "\CromatografiasPorc.txt"
End If

Open App.Path & "\CromatografiasPorc.txt" For Append As #2   ' Abre el archivo.

        LineaTexto = "Profundidad" & "," & "C1" & "," & "C2" & "," & "C3" & "," & "IC4" & "," & "NC4" & "," & "IC5" & "," & "NC5" & "," & "CO2"
        Print #2, LineaTexto

Contador = 0

    For Each objPozoAnalis In gObjPozoActivo.PozoAnalisis
        
        If objPozoAnalis.Seleccionado Then
            Contador = Contador + 1
    
            Profundidad = Format(objPozoAnalis.Profundidad, "#0")
            'objPozoAnalis.CargarComponentesRelaciones False
    
            C1 = "0"
            C2 = "0"
            C3 = "0"
            IC4 = "0"
            NC4 = "0"
            IC5 = "0"
            NC5 = "0"
    
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
    
            Select Case objPozoAnalisComponente.NumeroComponente
    
                Case 1: C1 = Format(objPozoAnalisComponente.NormArea, "#0")
                Case 2: C2 = Format(objPozoAnalisComponente.NormArea, "#0")
                Case 3: C3 = Format(objPozoAnalisComponente.NormArea, "#0")
                Case 4: IC4 = Format(objPozoAnalisComponente.NormArea, "#0")
                Case 5: NC4 = Format(objPozoAnalisComponente.NormArea, "#0")
                Case 6: IC5 = Format(objPozoAnalisComponente.NormArea, "#0")
                Case 7: NC5 = Format(objPozoAnalisComponente.NormArea, "#0")
    
            End Select
    
            Next
    
            strco2 = Format(objPozoAnalis.CO2, "#0")
            LineaTexto = Profundidad & "," & C1 & "," & C2 & "," & C3 & "," & IC4 & "," & NC4 & "," & IC5 & "," & NC5 & "," & strco2
    
            Print #2, LineaTexto
        End If
    Next

Close #2



Me.MousePointer = 0

End Sub



Public Sub GenerarTxtCronoRop()

Dim objPozoEscaneo                          As clsPozoEscaneo
Dim LineaTexto                              As String
Dim Indice                                  As Long

If UCase(Dir(App.Path & "\CronoRop.txt")) <> "" Then
    Kill App.Path & "\CronoRop.txt"
End If

Open App.Path & "\CronoRop.txt" For Append As #2   ' Abre el archivo.

        LineaTexto = "Profundidad" & "," & "Crono" & "," & "ROP" & "," & "GasTotal" & "," & "CO2"
        Print #2, LineaTexto


On Error GoTo Error
    
    Indice = gObjPozoActivo.PozoEscaneos.Count
    
    Do Until Indice = 0
        Set objPozoEscaneo = gObjPozoActivo.PozoEscaneos.ColPozoEscaneoPorIndice(Indice)
        If Not objPozoEscaneo Is Nothing Then
            LineaTexto = objPozoEscaneo.ProfundidadPozo & "," & Format(objPozoEscaneo.Crono, "0.00") & "," & Format(objPozoEscaneo.ROP, "0.00") & "," & objPozoEscaneo.GasTotal & "," & objPozoEscaneo.CO2
            Print #2, LineaTexto
        End If
        Indice = Indice - 1
    Loop

Close #2

Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: ConsultarTodos los cronos ." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Sub

Public Sub AnalizarCambiosEscala()


        If Signal > 5400 Then

            Select Case EContinuoLocal

                Case 1: EContinuoLocal = 10

                Case 10: EContinuoLocal = 100
                         ENormalLocal = 100
                         EFastLocal = 10

                Case 100:
            End Select
        Else

        End If




        If Signal > 600 And ENormalLocal = 1 Then

            ENormalLocal = 10
            EContinuoLocal = 10
            EFastLocal = 1

        End If


        If Signal < 500 Then

            Select Case EContinuoLocal



                Case 10: EContinuoLocal = 1
                         ENormalLocal = 1
                         EFastLocal = 1

                Case 100:   EContinuoLocal = 10
                            ENormalLocal = 10
                            EFastLocal = 1
            End Select

        Else

        End If

End Sub

Sub SetearMenuEscala(ByVal Indice As Long)
    
    On Error Resume Next
    
    If gIndiceEscala <> 0 Then
        mnuEscala(gIndiceEscala).Checked = False
    End If
    mnuEscala(Indice).Checked = True
    FrmAnalisis.iPlotGases.BeginUpdate
    FrmAnalisis.iPlotGases.YAxis(0).Span = Indice * 1000
    FrmAnalisis.iPlotGases.EndUpdate
    FrmAnalisis.iPlotGases.RepaintAll

    gIndiceEscala = Indice
    SaveSetting "Iga", "Config", "Escala", gIndiceEscala
    pObjAdminEventos.ActualizarVistaProf
    
End Sub

Sub SetearMenuEscalaLog(ByVal Indice As Long)
    On Error Resume Next
    If gIndiceEscalaLog <> 0 Then
        mnuEscalaLog(gIndiceEscalaLog).Checked = False
    End If
    mnuEscalaLog(Indice).Checked = True

    FrmAnalisis.iPlotMasterLog.BeginUpdate
    FrmAnalisis.iPlotMasterLog.YAxis(1).Span = Indice * 1000
    FrmAnalisis.iPlotMasterLog.EndUpdate
    FrmAnalisis.iPlotMasterLog.RepaintAll

    gIndiceEscalaLog = Indice
    SaveSetting "Iga", "Config", "EscalaLog", gIndiceEscalaLog
    pObjAdminEventos.ActualizarVistaTiempo 3
    
End Sub

Sub SetearMenuZonaMuerta(ByVal Indice As Long)
    On Error Resume Next
    Dim FechaInicio As Date
    Dim FechaFin As Date

    FechaFin = Month(Now()) * 30 + Day(Now()) + Hour(Now()) / 24 + Minute(Now()) / (24 * 60) + Second(Now()) / (24# * 60# * 60#)
    gRefrescarGrafico = True


    mnuZonaMuerta(gZonaMuerta).Checked = False
    mnuZonaMuerta(Indice).Checked = True
    gZonaMuerta = Indice
    If FrmAnalisis.ChkTiempoReal = 1 Then
        FrmAnalisis.LimpiarYConsultarDatosTiempo
        FechaInicio = DateAdd("n", -1 * gSpan, FechaFin)
        FrmAnalisis.iPlotGases.BeginUpdate
        FrmAnalisis.iPlotGases.XAxis(0).Min = Month(FechaInicio) * 30 + Day(FechaInicio) + Hour(FechaInicio) / 24 + Minute(FechaInicio) / (24 * 60) + Second(FechaInicio) / (24# * 60# * 60#)
        FrmAnalisis.iPlotGases.XAxis(0).Span = ObtenerSpanEnDías(FechaInicio, DateAdd("n", gZonaMuerta, FechaFin))
        FrmAnalisis.iPlotGases.EndUpdate
        FrmAnalisis.iPlotGases.XAxis(0).TrackingEnabled = True
    Else
        FrmAnalisis.ReconsultarDatosTiempo
    End If
    SaveSetting "Iga", "Config", "ZonaMuerta", gZonaMuerta
End Sub

Sub SetearMenuSpanTemporal(ByVal Indice As Long)
    On Error Resume Next
    Dim FechaInicio As Date
    Dim FechaFin As Date

    FechaFin = Month(Now()) * 30 + Day(Now()) + Hour(Now()) / 24 + Minute(Now()) / (24 * 60) + Second(Now()) / (24# * 60# * 60#)


    If gSpan <> 0 Then
        mnuSpan(gSpan).Checked = False
    End If
    mnuSpan(Indice).Checked = True
    gSpan = Indice
    If FrmAnalisis.ChkTiempoReal = 1 Then
        gRefrescarGrafico = True
        FrmAnalisis.LimpiarYConsultarDatosTiempo
        FechaInicio = DateAdd("n", -1 * (gSpan), FechaFin)
        FrmAnalisis.iPlotGases.BeginUpdate
        FrmAnalisis.iPlotGases.XAxis(0).Min = Month(FechaInicio) * 30 + Day(FechaInicio) + Hour(FechaInicio) / 24 + Minute(FechaInicio) / (24 * 60) + Second(FechaInicio) / (24# * 60# * 60#)
        FrmAnalisis.iPlotGases.XAxis(0).Span = ObtenerSpanEnDías(FechaInicio, DateAdd("n", gZonaMuerta, FechaFin))

        FrmAnalisis.iPlotGases.EndUpdate
        FrmAnalisis.iPlotGases.XAxis(0).TrackingEnabled = True
    Else
        FrmAnalisis.ReconsultarDatosTiempo
    End If
    SaveSetting "Iga", "Config", "SpanTemporal", gSpan

End Sub

Private Sub AnalizarRegistroTiempo()

    Dim ModuloGrabacion             As Integer
    Dim objPozoGasTiempo            As clsPozoGasTiempo
    
    If Not gObjPozoActivo Is Nothing Then
        ModuloGrabacion = 5
        If Not EstoyCargando Then
            If Intervalo Mod ModuloGrabacion = 0 And DateDiff("s", FechaHoraUltimaGrabacion, Now()) > 1 Then
                Set objPozoGasTiempo = New clsPozoGasTiempo
                objPozoGasTiempo.CO2 = CO2
                objPozoGasTiempo.Comentario = ComentarioGas
                objPozoGasTiempo.fecha = Now
                objPozoGasTiempo.GasTotal = GasTotal
                objPozoGasTiempo.IdPozo = gObjPozoActivo.IdPozo
                objPozoGasTiempo.SH2 = SH2
                
                Set gObjPozoActivo.PozoGasTiempos.Datos = gDatos
                gDatos.BeginTrans
                If gObjPozoActivo.PozoGasTiempos.dbAgregar(objPozoGasTiempo) Then
                    gDatos.CommitTrans
                    FechaHoraUltimaGrabacion = objPozoGasTiempo.fecha
                Else
                    gDatos.RollBackTrans
                End If
            End If
        End If
    End If
End Sub


Public Function Falsear_disparo_analisis() As Boolean

    Dim rs              As Recordset
    Dim StrSql          As String

    On Error GoTo Error

    StrSql = "UPDATE dc_dt_a_iga SET "
    StrSql = StrSql & "dispararanalisis = false "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idregistro = 1;"
    
    Falsear_disparo_analisis = gDatos.DbEjecutar(StrSql)

Error:

    If Err.Number <> 0 Then
        Falsear_disparo_analisis = False
        gDatos.Error.Description = Err.Description
        gDatos.Error.Metodo = "Falsear_disparo_analisis "
        gDatos.Error.Number = Err.Number
        gDatos.Error.Objeto = "frmPrincipal"
        gDatos.Error.Mostrar
        Err.Clear
        
    End If

End Function

Public Sub LeerComunicacionDC_DT()

    Dim rs              As Recordset
    Dim StrSql          As String
    Dim Desde

    On Error GoTo Error

    gMetroCortadoActual = gObjPozoActivo.PozoEscaneos.dbMaxProfPozoEscaneo(gObjPozoActivo.IdPozo)

    StrSql = "SELECT "
    StrSql = StrSql & "co2, gastotal, sh2, dispararanalisis, profundidadretorno "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "dc_dt_a_iga "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idregistro = 1;"
    
    Set rs = gDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            
            If Not IsNull(rs!ProfundidadRetorno) Then
                ProfundidadRetorno = rs!ProfundidadRetorno
            Else
                ProfundidadRetorno = 0
            End If
            
            Desde = gObjPozoActivo.PozoEscaneos.colMaxProfPozoEscaneo
            
            If gMetroCortadoActual > Desde Then
                'Debo actualizar la colección de escaneos
                gObjPozoActivo.PozoEscaneos.ConsultarPozoEscaneosPorProfundidad gObjPozoActivo.PozoEscaneos.colMaxProfPozoEscaneo, gMetroCortadoActual
                FrmAnalisis.ConsultarPozoEscaneos Desde
                pObjAdminEventos.ActualizarVistaProf
            ElseIf gMetroCortadoActual < gObjPozoActivo.PozoEscaneos.colMaxProfPozoEscaneo Then
                'Levantaron la profundidad de pozo en dc_dt, hay que consultar de nuevo la colección de metros cortados
                gObjPozoActivo.PozoEscaneos.Clear
                gObjPozoActivo.PozoEscaneos.ConsultarPozoEscaneos
                FrmAnalisis.ConsultarPozoEscaneos 0
                pObjAdminEventos.ActualizarVistaProf
            End If
            
            If ProfundidadRetorno <> ProfundidadRetornoAnterior Then
                gObjPozoActivo.PozoEscaneos.ConsultarGasesEscaneos gObjPozoActivo.IdPozo, ProfundidadRetornoAnterior, ProfundidadRetorno
                FrmAnalisis.ActualizarGasesEscaneos ProfundidadRetornoAnterior, ProfundidadRetorno
            End If
            
            If gObjConfiguracion.GasTotalDesdeDataCenter Then GasTotal = rs!GasTotal
            If gObjConfiguracion.CO2DesdeDataCenter Then CO2 = rs!CO2
            If gObjConfiguracion.SH2DesdeDataCenter Then SH2 = rs!SH2
            LlegoMuestra = rs!dispararanalisis
            
            If Not DeboTirarCroma Then
                DeboTirarCroma = rs!dispararanalisis
                If DeboTirarCroma Then
                    ProfundidadAnalisis = ProfundidadRetorno
                End If
            End If
            
            If ProfundidadRetorno <> 0 And ProfundidadRetorno <> ProfundidadRetornoAnterior And ProfundidadRetornoAnterior <> 0 Then
                DeboAgregarCartel = True
            Else
                DeboAgregarCartel = False
            End If
    
            ProfundidadRetornoAnterior = ProfundidadRetorno
 
        End If
    End If

Error:

    If Err.Number <> 0 Then
        Err.Clear
        ErrorLeerRetYGases = True
    End If

End Sub

Private Sub EscribirGases(ValorGasTotal As Variant, ValorCO2 As Variant, ValorSH2 As Variant)

    Dim StrSql            As String

    On Error GoTo Error

    StrSql = "update gasesiga set "
    StrSql = StrSql & "co2 = " & IIf(ValorCO2 > 0, ValorCO2, 0) & ", "
    StrSql = StrSql & "gastotal = " & IIf(ValorGasTotal > 0, ValorGasTotal, 0) & ", "
    StrSql = StrSql & "sh2 = " & IIf(ValorSH2 > 0, ValorSH2, 0) & " "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idmedicion = 1;"

    gDatos.BeginTrans

    gDatos.DbEjecutar StrSql

    If gDatos.RegistrosAfectados >= 1 Then
        gDatos.CommitTrans
    Else
        gDatos.RollBackTrans
    End If

Error:

    If Err.Number <> 0 Then
        Err.Clear
        ErrorEscrituraDeGas = True
    End If

End Sub
