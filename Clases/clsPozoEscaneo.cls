VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoEscaneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIdPozo                                 As Long
Private pIdPozoEscaneo                          As Long
Private pIdReporte                              As Long
Private pCO2                                    As Double
Private pCrono                                  As Double
Private pGasTotal                               As Double
Private pROP                                    As Double
Private pSH2                                    As Double
Private pProfundidadPozo                        As Double

Public Property Get Key() As String
    Key = pIdPozo & pIdPozoEscaneo & "ID"
End Property

Public Property Get IdPozo() As Long
    IdPozo = pIdPozo
End Property

Public Property Let IdPozo(ByVal vNewValue As Long)
    pIdPozo = vNewValue
End Property

Public Property Get IdPozoEscaneo() As Long
    IdPozoEscaneo = pIdPozoEscaneo
End Property

Public Property Let IdPozoEscaneo(ByVal vNewValue As Long)
    pIdPozoEscaneo = vNewValue
End Property

Public Property Get IdReporte() As Long
    IdReporte = pIdReporte
End Property

Public Property Let IdReporte(ByVal vNewValue As Long)
    pIdReporte = vNewValue
End Property

Public Property Get CO2() As Double
    CO2 = pCO2
End Property

Public Property Let CO2(ByVal vNewValue As Double)
    pCO2 = vNewValue
End Property

Public Property Get Crono() As Double
    Crono = pCrono
End Property

Public Property Let Crono(ByVal vNewValue As Double)
    pCrono = vNewValue
End Property

Public Property Get GasTotal() As Double
    GasTotal = pGasTotal
End Property

Public Property Let GasTotal(ByVal vNewValue As Double)
    pGasTotal = vNewValue
End Property

Public Property Get ROP() As Double
    ROP = pROP
End Property

Public Property Let ROP(ByVal vNewValue As Double)
    pROP = vNewValue
End Property

Public Property Get SH2() As Double
    SH2 = pSH2
End Property

Public Property Let SH2(ByVal vNewValue As Double)
    pSH2 = vNewValue
End Property

Public Property Get ProfundidadPozo() As Double
    ProfundidadPozo = pProfundidadPozo
End Property

Public Property Let ProfundidadPozo(ByVal vNewValue As Double)
    pProfundidadPozo = vNewValue
End Property

Public Property Get Duplicado() As clsPozoEscaneo

    Dim objPozoEscaneo              As New clsPozoEscaneo
    
    objPozoEscaneo.CO2 = Me.CO2
    objPozoEscaneo.Crono = Me.Crono
    objPozoEscaneo.GasTotal = Me.GasTotal
    objPozoEscaneo.IdPozo = Me.IdPozo
    objPozoEscaneo.IdPozoEscaneo = Me.IdPozoEscaneo
    objPozoEscaneo.IdReporte = Me.IdReporte
    objPozoEscaneo.ProfundidadPozo = Me.ProfundidadPozo
    objPozoEscaneo.ROP = Me.ROP
    objPozoEscaneo.SH2 = Me.SH2
    
    Set Duplicado = objPozoEscaneo
    
End Property
