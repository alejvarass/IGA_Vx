VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIdPozo                                 As Long
Private pDiametroPozo                           As Double
Private pNombrePozo                             As String
Private pSupervisor                             As String
Private pCordenadaX                             As Double
Private pCordenadaY                             As Double
Private pCordenadaZ                             As Double
Private pProfundidadInicioControlG              As Double
Private pIdEquipo                               As String
Private pActivo                                 As Boolean
Private pMaxAlarmas                             As Long
Private pNorte                                  As Double
Private pEste                                   As Double
Private pIdFormulaTVD                           As Long
Private pIdPozoWeb                              As Long
Private pIdCuenca                               As Long
Private pProfundidadActual                      As Double
Private pPozoAnalisis                           As New clsPozoAnalisis
Private pPozoGasTiempos                         As New clsPozoGasTiempos
Private pPozoEscaneos                           As New clsPozoEscaneos

Public Property Get Key() As String
    Key = pIdPozo & "ID"
End Property

Public Property Get IdPozo() As Long
    IdPozo = pIdPozo
End Property

Public Property Let IdPozo(ByVal vNewValue As Long)
    pIdPozo = vNewValue
End Property

Public Property Get NombrePozo() As String
    NombrePozo = pNombrePozo
End Property

Public Property Let NombrePozo(ByVal vNewValue As String)
    pNombrePozo = vNewValue
End Property

Public Property Get Supervisor() As String
    Supervisor = pSupervisor
End Property

Public Property Let Supervisor(ByVal vNewValue As String)
    pSupervisor = vNewValue
End Property

Public Property Get CordenadaX() As Double
    CordenadaX = pCordenadaX
End Property

Public Property Let CordenadaX(ByVal vNewValue As Double)
    pCordenadaX = vNewValue
End Property

Public Property Get CordenadaY() As Double
    CordenadaY = pCordenadaY
End Property

Public Property Let CordenadaY(ByVal vNewValue As Double)
    pCordenadaY = vNewValue
End Property

Public Property Get CordenadaZ() As Double
    CordenadaZ = pCordenadaZ
End Property

Public Property Let CordenadaZ(ByVal vNewValue As Double)
    pCordenadaZ = vNewValue
End Property

Public Property Get ProfundidadInicioControlG() As Double
    ProfundidadInicioControlG = pProfundidadInicioControlG
End Property

Public Property Let ProfundidadInicioControlG(ByVal vNewValue As Double)
    pProfundidadInicioControlG = vNewValue
End Property

Public Property Get DiametroPozo() As Double
    DiametroPozo = pDiametroPozo
End Property

Public Property Let DiametroPozo(ByVal vNewValue As Double)
    pDiametroPozo = vNewValue
End Property

Public Property Get IdEquipo() As String
    IdEquipo = pIdEquipo
End Property

Public Property Let IdEquipo(ByVal vNewValue As String)
    pIdEquipo = vNewValue
End Property

Public Property Get Activo() As Boolean
    Activo = pActivo
End Property

Public Property Let Activo(ByVal vNewValue As Boolean)
    pActivo = vNewValue
End Property

Public Property Get ProfundidadActual() As Double
    ProfundidadActual = pProfundidadActual
End Property

Public Property Let ProfundidadActual(ByVal vNewValue As Double)
    pProfundidadActual = vNewValue
End Property

Public Property Get PozoAnalisis() As clsPozoAnalisis
    Set PozoAnalisis = pPozoAnalisis
End Property

Public Property Get PozoGasTiempos() As clsPozoGasTiempos
    Set PozoGasTiempos = pPozoGasTiempos
End Property

Public Property Get PozoEscaneos() As clsPozoEscaneos
    Set PozoEscaneos = pPozoEscaneos
End Property

Public Property Get Duplicado() As clsPozo
    
    Dim objPozo                     As New clsPozo
    Dim objPozoAnalis               As clsPozoAnalis
    Dim objPozoGasTiempo            As clsPozoGasTiempo
    Dim objPozoEscaneo              As clsPozoEscaneo
    
    objPozo.Activo = Me.Activo
    objPozo.CordenadaX = Me.CordenadaX
    objPozo.CordenadaY = Me.CordenadaY
    objPozo.CordenadaZ = Me.CordenadaZ
    objPozo.DiametroPozo = Me.DiametroPozo
    objPozo.IdCuenca = Me.IdCuenca
    objPozo.IdEquipo = Me.IdEquipo
    objPozo.IdFormulaTVD = Me.IdFormulaTVD
    objPozo.IdPozo = Me.IdPozo
    objPozo.IdPozoWeb = Me.IdPozoWeb
    objPozo.NombrePozo = Me.NombrePozo
    objPozo.ProfundidadActual = Me.ProfundidadActual
    objPozo.ProfundidadInicioControlG = Me.ProfundidadInicioControlG
    objPozo.Supervisor = Me.Supervisor
    
    For Each objPozoAnalis In Me.PozoAnalisis
        objPozo.PozoAnalisis.colAgregar objPozoAnalis.Duplicado
    Next
    
    For Each objPozoEscaneo In Me.PozoEscaneos
        objPozo.PozoEscaneos.colAgregar objPozoEscaneo.Duplicado
    Next
    
    Set Duplicado = objPozo
    
    
End Property

Public Sub Clonar(ByRef objPozo As clsPozo)
    
    objPozo.Activo = Me.Activo
    objPozo.CordenadaX = Me.CordenadaX
    objPozo.CordenadaY = Me.CordenadaY
    objPozo.CordenadaZ = Me.CordenadaZ
    objPozo.DiametroPozo = Me.DiametroPozo
    objPozo.IdCuenca = Me.IdCuenca
    objPozo.IdEquipo = Me.IdEquipo
    objPozo.IdFormulaTVD = Me.IdFormulaTVD
    objPozo.IdPozo = Me.IdPozo
    objPozo.IdPozoWeb = Me.IdPozoWeb
    objPozo.NombrePozo = Me.NombrePozo
    objPozo.ProfundidadActual = Me.ProfundidadActual
    objPozo.ProfundidadInicioControlG = Me.ProfundidadInicioControlG
    objPozo.Supervisor = Me.Supervisor
    
End Sub

Public Property Get Norte() As Double
    Norte = pNorte
End Property

Public Property Let Norte(ByVal vNewValue As Double)
    pNorte = vNewValue
End Property

Public Property Get Este() As Double
    Este = pEste
End Property

Public Property Let Este(ByVal vNewValue As Double)
    pEste = vNewValue
End Property

Public Property Get IdFormulaTVD() As Long
    IdFormulaTVD = pIdFormulaTVD
End Property

Public Property Let IdFormulaTVD(ByVal vNewValue As Long)
    pIdFormulaTVD = vNewValue
End Property

Public Property Get IdPozoWeb() As Long
    IdPozoWeb = pIdPozoWeb
End Property

Public Property Let IdPozoWeb(ByVal vNewValue As Long)
    pIdPozoWeb = vNewValue
End Property

Public Property Get IdCuenca() As Long
    IdCuenca = pIdCuenca
End Property

Public Property Let IdCuenca(ByVal vNewValue As Long)
    pIdCuenca = vNewValue
End Property
