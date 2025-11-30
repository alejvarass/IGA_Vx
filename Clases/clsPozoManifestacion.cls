VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoManifestacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIDPozo                                 As Long
Private pIdManifestacion                        As Long
Private pIntervaloDesde                         As Long
Private pIntervaloHasta                         As Long
Private pCodigoComposicion                      As Long
Private pCodigoPorosidadVisual                  As Long
Private pGasTotal                               As Long
Private pBackG                                  As Long
Private pC1                                     As Long
Private pC2                                     As Long
Private pC3                                     As Long
Private pIC4                                    As Long
Private pNC4                                    As Long
Private pIC5                                    As Long
Private pNC5                                    As Long
Private pCodigoRastrosCantidad                  As Long
Private pCodigoColor                            As Long
Private pCodigoTipoDeRastro                     As Long
Private pCodigoFD                               As Long
Private pCodigoFE                               As Long

Public Property Get Key() As String
    Key = pIDPozo & pIdManifestacion & "ID"
End Property

Public Property Get IdPozo() As Long
    IdPozo = pIDPozo
End Property

Public Property Let IdPozo(ByVal vNewValue As Long)
    pIDPozo = vNewValue
End Property

Public Property Get IdManifestacion() As Long
    IdManifestacion = pIdManifestacion
End Property

Public Property Let IdManifestacion(ByVal vNewValue As Long)
    pIdManifestacion = vNewValue
End Property

Public Property Get IntervaloDesde() As Long
    IntervaloDesde = pIntervaloDesde
End Property

Public Property Let IntervaloDesde(ByVal vNewValue As Long)
    pIntervaloDesde = vNewValue
End Property

Public Property Get IntervaloHasta() As Long
    IntervaloHasta = pIntervaloHasta
End Property

Public Property Let IntervaloHasta(ByVal vNewValue As Long)
    pIntervaloHasta = vNewValue
End Property

Public Property Get CodigoComposicion() As Long
    CodigoComposicion = pCodigoComposicion
End Property

Public Property Let CodigoComposicion(ByVal vNewValue As Long)
    pCodigoComposicion = vNewValue
End Property

Public Property Get CodigoPorosidadVisual() As Long
    CodigoPorosidadVisual = pCodigoPorosidadVisual
End Property

Public Property Let CodigoPorosidadVisual(ByVal vNewValue As Long)
    pCodigoPorosidadVisual = vNewValue
End Property

Public Property Get GasTotal() As Long
    GasTotal = pGasTotal
End Property

Public Property Let GasTotal(ByVal vNewValue As Long)
    pGasTotal = vNewValue
End Property

Public Property Get BackG() As Long
    BackG = pBackG
End Property

Public Property Let BackG(ByVal vNewValue As Long)
    pBackG = vNewValue
End Property

Public Property Get C1() As Long
    C1 = pC1
End Property

Public Property Let C1(ByVal vNewValue As Long)
    pC1 = vNewValue
End Property

Public Property Get C2() As Long
    C2 = pC2
End Property

Public Property Let C2(ByVal vNewValue As Long)
    pC2 = vNewValue
End Property

Public Property Get C3() As Long
    C3 = pC3
End Property

Public Property Let C3(ByVal vNewValue As Long)
    pC3 = vNewValue
End Property

Public Property Get IC4() As Long
    IC4 = pIC4
End Property

Public Property Let IC4(ByVal vNewValue As Long)
    pIC4 = vNewValue
End Property

Public Property Get NC4() As Long
    NC4 = pNC4
End Property

Public Property Let NC4(ByVal vNewValue As Long)
    pNC4 = vNewValue
End Property

Public Property Get IC5() As Long
    IC5 = pIC5
End Property

Public Property Let IC5(ByVal vNewValue As Long)
    pIC5 = vNewValue
End Property

Public Property Get NC5() As Long
    NC5 = pNC5
End Property

Public Property Let NC5(ByVal vNewValue As Long)
    pNC5 = vNewValue
End Property

Public Property Get CodigoRastrosCantidad() As Long
    CodigoRastrosCantidad = pCodigoRastrosCantidad
End Property

Public Property Let CodigoRastrosCantidad(ByVal vNewValue As Long)
    pCodigoRastrosCantidad = vNewValue
End Property

Public Property Get CodigoColor() As Long
    CodigoColor = pCodigoColor
End Property

Public Property Let CodigoColor(ByVal vNewValue As Long)
    pCodigoColor = vNewValue
End Property

Public Property Get CodigoTipoDeRastro() As Long
    CodigoTipoDeRastro = pCodigoTipoDeRastro
End Property

Public Property Let CodigoTipoDeRastro(ByVal vNewValue As Long)
    pCodigoTipoDeRastro = vNewValue
End Property

Public Property Get CodigoFD() As Long
    CodigoFD = pCodigoFD
End Property

Public Property Let CodigoFD(ByVal vNewValue As Long)
    pCodigoFD = vNewValue
End Property

Public Property Get CodigoFE() As Long
    CodigoFE = pCodigoFE
End Property

Public Property Let CodigoFE(ByVal vNewValue As Long)
    pCodigoFE = vNewValue
End Property

Public Property Get GasTotalSobreBackG() As Double
    GasTotalSobreBackG = Format(GasTotal / BackG, "0.00")
End Property

Public Property Get Pronostico() As Double
    Pronostico = GasTotalSobreBackG * gObjIndicesPonderacion.colIndicePonderacion("POROSVISUAL", Me.CodigoPorosidadVisual).Valor * gObjIndicesPonderacion.colIndicePonderacion("COLOR", Me.CodigoColor).Valor * gObjIndicesPonderacion.colIndicePonderacion("RASTROSCANTIDAD", Me.CodigoRastrosCantidad).Valor * gObjIndicesPonderacion.colIndicePonderacion("TIPOSDERASTROS", Me.CodigoTipoDeRastro).Valor * gObjIndicesPonderacion.colIndicePonderacion("FD", Me.CodigoFD).Valor * gObjIndicesPonderacion.colIndicePonderacion("FE", Me.CodigoFE).Valor
End Property
