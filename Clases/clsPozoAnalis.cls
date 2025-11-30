VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoAnalis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNumeroAnalisis                         As Long
Private pIdPozo                                 As Long
Private pArchivo                                As String
Private pFecha                                  As Date
Private pProfundidad                            As Double
Private pCO2                                    As Long
Private pSH2                                    As Long
Private pGasTotal                               As Long
Private pGasTotalCromatografico                 As Long
Private pCodigoTipoDeAnalisis                   As Long
Private pEFast                                  As Long
Private pENormal                                As Long
Private pObservaciones                          As String
Private pSeleccionado                           As Boolean
Private pBar2                                   As Double
Private pBar3                                   As Double
Private pBar4                                   As Double
Private pBar5                                   As Double
Private pCous1                                  As Double
Private pCous2                                  As Double
Private pWH                                     As Double
Private pBH                                     As Double
Private pCH                                     As Double
Private pGeo1                                   As Double
Private pGeo2                                   As Double
Private pGeo3                                   As Double
Private pGeo4                                   As Double
Private pSnGeo                                  As Double
Private pC1                                     As Double
Private pC2                                     As Double
Private pC3                                     As Double
Private pC4                                     As Double
Private pC5                                     As Double
Private pIC4                                    As Double
Private pNC4                                    As Double
Private pIC5                                    As Double
Private pNC5                                    As Double
Private pIdPozoEscaneo                          As Long
Private pPozoAnalisComponentes                  As New clsPozoAnalisComponentes

Public Property Get Key() As String
    Key = pIdPozo & pNumeroAnalisis & "ID"
End Property

Public Property Get NumeroAnalisis() As Long
    NumeroAnalisis = pNumeroAnalisis
End Property

Public Property Let NumeroAnalisis(ByVal Valor As Long)
    pNumeroAnalisis = Valor
End Property

Public Property Get IdPozo() As Long
    IdPozo = pIdPozo
End Property

Public Property Let IdPozo(ByVal Valor As Long)
    pIdPozo = Valor
End Property

Public Property Get archivo() As String
    archivo = pArchivo
End Property

Public Property Let archivo(ByVal Valor As String)
    pArchivo = Valor
End Property

Public Property Get Fecha() As Date
    Fecha = pFecha
End Property

Public Property Let Fecha(ByVal Valor As Date)
    pFecha = Valor
End Property

Public Property Get Profundidad() As Double
    Profundidad = pProfundidad
End Property

Public Property Let Profundidad(ByVal Valor As Double)
    pProfundidad = Valor
End Property

Public Property Get CO2() As Long
    CO2 = pCO2
End Property

Public Property Let CO2(ByVal Valor As Long)
    pCO2 = Valor
End Property

Public Property Get SH2() As Long
    SH2 = pSH2
End Property

Public Property Let SH2(ByVal Valor As Long)
    pSH2 = Valor
End Property

Public Property Get GasTotal() As Long
    GasTotal = pGasTotal
End Property

Public Property Let GasTotal(ByVal Valor As Long)
    pGasTotal = Valor
End Property

Public Property Get GasTotalCromatografico() As Long
    GasTotalCromatografico = pGasTotalCromatografico
End Property

Public Property Let GasTotalCromatografico(ByVal Valor As Long)
    pGasTotalCromatografico = Valor
End Property

Public Property Get CodigoTipoDeAnalisis() As Long
    CodigoTipoDeAnalisis = pCodigoTipoDeAnalisis
End Property

Public Property Let CodigoTipoDeAnalisis(ByVal Valor As Long)
    pCodigoTipoDeAnalisis = Valor
End Property

Public Property Get EFast() As Long
    EFast = pEFast
End Property

Public Property Let EFast(ByVal Valor As Long)
    pEFast = Valor
End Property

Public Property Get ENormal() As Long
    ENormal = pENormal
End Property

Public Property Let ENormal(ByVal Valor As Long)
    pENormal = Valor
End Property

Public Property Get Observaciones() As String
    Observaciones = pObservaciones
End Property

Public Property Let Observaciones(ByVal Valor As String)
    pObservaciones = Valor
End Property

Public Property Get Seleccionado() As Boolean
    Seleccionado = pSeleccionado
End Property

Public Property Let Seleccionado(ByVal Valor As Boolean)
    pSeleccionado = Valor
End Property

Public Property Get PozoAnalisComponentes() As clsPozoAnalisComponentes
    Set PozoAnalisComponentes = pPozoAnalisComponentes
End Property

Public Property Get Bar2() As Double
    Bar2 = pBar2
End Property

Public Property Let Bar2(ByVal vNewValue As Double)
    pBar2 = vNewValue
End Property

Public Property Get Bar3() As Double
    Bar3 = pBar3
End Property

Public Property Let Bar3(ByVal vNewValue As Double)
    pBar3 = vNewValue
End Property

Public Property Get Bar4() As Double
    Bar4 = pBar4
End Property

Public Property Let Bar4(ByVal vNewValue As Double)
    pBar4 = vNewValue
End Property

Public Property Get Bar5() As Double
    Bar5 = pBar5
End Property

Public Property Let Bar5(ByVal vNewValue As Double)
    pBar5 = vNewValue
End Property

Public Property Get Cous1() As Double
    Cous1 = pCous1
End Property

Public Property Let Cous1(ByVal vNewValue As Double)
    pCous1 = vNewValue
End Property

Public Property Get Cous2() As Double
    Cous2 = pCous2
End Property

Public Property Let Cous2(ByVal vNewValue As Double)
    pCous2 = vNewValue
End Property

Public Property Get WH() As Double
    WH = pWH
End Property

Public Property Let WH(ByVal vNewValue As Double)
    pWH = vNewValue
End Property
Public Property Get BH() As Double
    BH = pBH
End Property

Public Property Let BH(ByVal vNewValue As Double)
    pBH = vNewValue
End Property

Public Property Get CH() As Double
    CH = pCH
End Property

Public Property Let CH(ByVal vNewValue As Double)
    pCH = vNewValue
End Property

Public Property Get Geo1() As Double
    Geo1 = pGeo1
End Property

Public Property Let Geo1(ByVal vNewValue As Double)
    pGeo1 = vNewValue
End Property

Public Property Get Geo2() As Double
    Geo2 = pGeo2
End Property

Public Property Let Geo2(ByVal vNewValue As Double)
    pGeo2 = vNewValue
End Property

Public Property Get Geo3() As Double
    Geo3 = pGeo3
End Property

Public Property Let Geo3(ByVal vNewValue As Double)
    pGeo3 = vNewValue
End Property

Public Property Get Geo4() As Double
    Geo4 = pGeo4
End Property

Public Property Let Geo4(ByVal vNewValue As Double)
    pGeo4 = vNewValue
End Property

Public Property Get SnGeo() As Double
    SnGeo = pSnGeo
End Property

Public Property Let SnGeo(ByVal vNewValue As Double)
    pSnGeo = vNewValue
End Property

Public Property Get C1() As Double
    C1 = pC1
End Property

Public Property Let C1(ByVal vNewValue As Double)
    pC1 = vNewValue
End Property

Public Property Get C2() As Double
    C2 = pC2
End Property

Public Property Let C2(ByVal vNewValue As Double)
    pC2 = vNewValue
End Property

Public Property Get C3() As Double
    C3 = pC3
End Property

Public Property Let C3(ByVal vNewValue As Double)
    pC3 = vNewValue
End Property

Public Property Get C4() As Double
    C4 = pC4
End Property

Public Property Let C4(ByVal vNewValue As Double)
    pC4 = vNewValue
End Property

Public Property Get C5() As Double
    C5 = pC5
End Property

Public Property Let C5(ByVal vNewValue As Double)
    pC5 = vNewValue
End Property

Public Property Get IC4() As Double
    IC4 = pIC4
End Property

Public Property Let IC4(ByVal vNewValue As Double)
    pIC4 = vNewValue
End Property

Public Property Get NC4() As Double
    NC4 = pNC4
End Property

Public Property Let NC4(ByVal vNewValue As Double)
    pNC4 = vNewValue
End Property

Public Property Get IC5() As Double
    IC5 = pIC5
End Property

Public Property Let IC5(ByVal vNewValue As Double)
    pIC5 = vNewValue
End Property

Public Property Get NC5() As Double
    NC5 = pNC5
End Property

Public Property Let NC5(ByVal vNewValue As Double)
    pNC5 = vNewValue
End Property

Public Property Get IdPozoEscaneo() As Double
    IdPozoEscaneo = pIdPozoEscaneo
End Property

Public Property Let IdPozoEscaneo(ByVal vNewValue As Double)
    pIdPozoEscaneo = vNewValue
End Property



Public Property Get Duplicado() As clsPozoAnalis
    
    Dim objPozoAnalis                       As New clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente
    
    objPozoAnalis.archivo = Me.archivo
    objPozoAnalis.Bar2 = Me.Bar2
    objPozoAnalis.Bar3 = Me.Bar3
    objPozoAnalis.Bar4 = Me.Bar4
    objPozoAnalis.Bar5 = Me.Bar5
    objPozoAnalis.BH = Me.BH
    objPozoAnalis.C1 = Me.C1
    objPozoAnalis.C2 = Me.C2
    objPozoAnalis.C3 = Me.C3
    objPozoAnalis.C4 = Me.C4
    objPozoAnalis.C5 = Me.C5
    objPozoAnalis.CH = Me.CH
    objPozoAnalis.CO2 = Me.CO2
    objPozoAnalis.CodigoTipoDeAnalisis = Me.CodigoTipoDeAnalisis
    objPozoAnalis.Cous1 = Me.Cous1
    objPozoAnalis.Cous2 = Me.Cous2
    objPozoAnalis.EFast = Me.EFast
    objPozoAnalis.ENormal = Me.ENormal
    objPozoAnalis.Fecha = Me.Fecha
    objPozoAnalis.GasTotal = Me.GasTotal
    objPozoAnalis.GasTotalCromatografico = Me.GasTotalCromatografico
    objPozoAnalis.Geo1 = Me.Geo1
    objPozoAnalis.Geo2 = Me.Geo2
    objPozoAnalis.Geo3 = Me.Geo3
    objPozoAnalis.Geo4 = Me.Geo4
    objPozoAnalis.IC4 = Me.IC4
    objPozoAnalis.IC5 = Me.IC5
    objPozoAnalis.IdPozo = Me.IdPozo
    objPozoAnalis.NC4 = Me.NC4
    objPozoAnalis.NC5 = Me.NC5
    objPozoAnalis.NumeroAnalisis = Me.NumeroAnalisis
    objPozoAnalis.Observaciones = Me.Observaciones
    objPozoAnalis.Profundidad = Me.Profundidad
    objPozoAnalis.Seleccionado = Me.Seleccionado
    objPozoAnalis.SH2 = Me.SH2
    objPozoAnalis.SnGeo = Me.SnGeo
    objPozoAnalis.WH = Me.WH
    objPozoAnalis.IdPozoEscaneo = Me.IdPozoEscaneo
    
    For Each objPozoAnalisComponente In Me.PozoAnalisComponentes
        objPozoAnalis.PozoAnalisComponentes.colAgregar objPozoAnalisComponente.Duplicado
    Next
    
    Set Duplicado = objPozoAnalis
    
    
End Property

