VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pClave                              As Long
Private pArchivoDeResultados                As String
Private pArchivoDeResultadosGenerales       As String
Private pAnalisisAutomaticos                As Boolean
Private pCambioEscalas                      As Boolean
Private pTomaProfundidad                    As Boolean
Private pDisparoCiclico                     As Boolean
Private pPuertoComm                         As Long
Private pTiempoAnalisis                     As Long
Private pNivelDisparoMetroMetro             As Long
Private pCO2DesdeDataCenter                 As Boolean
Private pSH2DesdeDataCenter                 As Boolean
Private pGasTotalDesdeDataCenter            As Boolean
Private pIpRabit                            As String
Private pIpDisparoCroma                     As String
Private pIdTipoEquipoGas                    As Long
Private pIdTipoEquipoCroma                  As Long

Public Property Get Key() As String
    Key = pClave & "ID"
End Property

Public Property Get Clave() As Long
    Clave = pClave
End Property

Public Property Let Clave(ByVal vNewValue As Long)
    pClave = vNewValue
End Property

Public Property Get ArchivoDeResultados() As String
    ArchivoDeResultados = pArchivoDeResultados
End Property

Public Property Let ArchivoDeResultados(ByVal vNewValue As String)
    pArchivoDeResultados = vNewValue
End Property

Public Property Get ArchivoDeResultadosGenerales() As String
    ArchivoDeResultadosGenerales = pArchivoDeResultadosGenerales
End Property

Public Property Let ArchivoDeResultadosGenerales(ByVal vNewValue As String)
    pArchivoDeResultadosGenerales = vNewValue
End Property

Public Property Get AnalisisAutomaticos() As Boolean
    AnalisisAutomaticos = pAnalisisAutomaticos
End Property

Public Property Let AnalisisAutomaticos(ByVal vNewValue As Boolean)
    pAnalisisAutomaticos = vNewValue
End Property

Public Property Get CambioEscalas() As Boolean
    CambioEscalas = pCambioEscalas
End Property

Public Property Let CambioEscalas(ByVal vNewValue As Boolean)
    pCambioEscalas = vNewValue
End Property

Public Property Get TomaProfundidad() As Boolean
    TomaProfundidad = pTomaProfundidad
End Property

Public Property Let TomaProfundidad(ByVal vNewValue As Boolean)
    pTomaProfundidad = vNewValue
End Property

Public Property Get DisparoCiclico() As Boolean
    DisparoCiclico = pDisparoCiclico
End Property

Public Property Let DisparoCiclico(ByVal vNewValue As Boolean)
    pDisparoCiclico = vNewValue
End Property

Public Property Get PuertoComm() As Long
    PuertoComm = pPuertoComm
End Property

Public Property Let PuertoComm(ByVal vNewValue As Long)
    pPuertoComm = vNewValue
End Property

Public Property Get TiempoAnalisis() As Long
    TiempoAnalisis = pTiempoAnalisis
End Property

Public Property Let TiempoAnalisis(ByVal vNewValue As Long)
    pTiempoAnalisis = vNewValue
End Property

Public Property Get NivelDisparoMetroMetro() As Long
    NivelDisparoMetroMetro = pNivelDisparoMetroMetro
End Property

Public Property Let NivelDisparoMetroMetro(ByVal vNewValue As Long)
    pNivelDisparoMetroMetro = vNewValue
End Property

Public Property Get CO2DesdeDataCenter() As Boolean
    CO2DesdeDataCenter = pCO2DesdeDataCenter
End Property

Public Property Let CO2DesdeDataCenter(ByVal vNewValue As Boolean)
    pCO2DesdeDataCenter = vNewValue
End Property

Public Property Get SH2DesdeDataCenter() As Boolean
    SH2DesdeDataCenter = pSH2DesdeDataCenter
End Property

Public Property Let SH2DesdeDataCenter(ByVal vNewValue As Boolean)
    pSH2DesdeDataCenter = vNewValue
End Property

Public Property Get GasTotalDesdeDataCenter() As Boolean
    GasTotalDesdeDataCenter = pGasTotalDesdeDataCenter
End Property

Public Property Let GasTotalDesdeDataCenter(ByVal vNewValue As Boolean)
    pGasTotalDesdeDataCenter = vNewValue
End Property

Public Property Get IpRabit() As String
    IpRabit = pIpRabit
End Property

Public Property Let IpRabit(ByVal vNewValue As String)
    pIpRabit = vNewValue
End Property

Public Property Get IdTipoEquipoGas() As Long
    IdTipoEquipoGas = pIdTipoEquipoGas
End Property

Public Property Let IdTipoEquipoGas(ByVal vNewValue As Long)
    pIdTipoEquipoGas = vNewValue
End Property

Public Property Get IdTipoEquipoCroma() As Long
    IdTipoEquipoCroma = pIdTipoEquipoCroma
End Property

Public Property Let IdTipoEquipoCroma(ByVal vNewValue As Long)
    pIdTipoEquipoCroma = vNewValue
End Property

Public Property Get IpDisparoCroma() As String
    IpDisparoCroma = pIpDisparoCroma
End Property

Public Property Let IpDisparoCroma(ByVal vNewValue As String)
    pIpDisparoCroma = vNewValue
End Property
