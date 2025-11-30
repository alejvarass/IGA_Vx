VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoGasTiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIDPozo                     As Long
Private pFecha                      As Date
Private pGasTotal                   As Double
Private pCO2                        As Double
Private pSH2                        As Double
Private pComentario                 As String

Public Property Get Key() As String
    Key = pIDPozo & pFecha & "ID"
End Property

Public Property Get IDPozo() As Long
    IDPozo = pIDPozo
End Property

Public Property Let IDPozo(ByVal vNewValue As Long)
    pIDPozo = vNewValue
End Property

Public Property Get Fecha() As Date
    Fecha = pFecha
End Property

Public Property Let Fecha(ByVal vNewValue As Date)
    pFecha = vNewValue
End Property

Public Property Get GasTotal() As Double
    GasTotal = pGasTotal
End Property

Public Property Let GasTotal(ByVal vNewValue As Double)
    pGasTotal = vNewValue
End Property

Public Property Get CO2() As Double
    CO2 = pCO2
End Property

Public Property Let CO2(ByVal vNewValue As Double)
    pCO2 = vNewValue
End Property

Public Property Get SH2() As Double
    SH2 = pSH2
End Property

Public Property Let SH2(ByVal vNewValue As Double)
    pSH2 = vNewValue
End Property

Public Property Get Comentario() As String
    Comentario = pComentario
End Property

Public Property Let Comentario(ByVal vNewValue As String)
    pComentario = vNewValue
End Property
