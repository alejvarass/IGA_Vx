VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPozoAnalisComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pNumeroAnalisis As Long
Private pIdPozo As Long
Private pNumeroComponente As Long
Private pComponent As String
Private pRetention As Double
Private pArea As Double
Private pExterno As Long
Private pUnits As String
Private pNormArea As Double

Public Property Get Key() As String
    Key = pIdPozo & pNumeroAnalisis & pNumeroComponente & "ID"
End Property

Public Property Get NumeroAnalisis() As Long
    NumeroAnalisis = pNumeroAnalisis
End Property

Public Property Let NumeroAnalisis(ByVal vNewValue As Long)
    pNumeroAnalisis = vNewValue
End Property

Public Property Get idPozo() As Long
    idPozo = pIdPozo
End Property

Public Property Let idPozo(ByVal vNewValue As Long)
    pIdPozo = vNewValue
End Property

Public Property Get NumeroComponente() As Long
    NumeroComponente = pNumeroComponente
End Property

Public Property Let NumeroComponente(ByVal vNewValue As Long)
    pNumeroComponente = vNewValue
End Property

Public Property Get Component() As String
    Component = pComponent
End Property

Public Property Let Component(ByVal vNewValue As String)
    pComponent = vNewValue
End Property

Public Property Get Retention() As Double
    Retention = pRetention
End Property

Public Property Let Retention(ByVal vNewValue As Double)
    pRetention = vNewValue
End Property

Public Property Get Area() As Double
    Area = pArea
End Property

Public Property Let Area(ByVal vNewValue As Double)
    pArea = vNewValue
End Property

Public Property Get Externo() As Long
    Externo = pExterno
End Property

Public Property Let Externo(ByVal vNewValue As Long)
    pExterno = vNewValue
End Property

Public Property Get Units() As String
    Units = pUnits
End Property

Public Property Let Units(ByVal vNewValue As String)
    pUnits = vNewValue
End Property

Public Property Get NormArea() As Double
    NormArea = pNormArea
End Property

Public Property Let NormArea(ByVal vNewValue As Double)
    pNormArea = vNewValue
End Property

Public Property Get Duplicado() As ClsPozoAnalisComponente
    
    Dim objComponente As New ClsPozoAnalisComponente
    
    On Error GoTo Error
    
    objComponente.Area = pArea
    objComponente.Component = pComponent
    objComponente.Externo = pExterno
    objComponente.idPozo = pIdPozo
    objComponente.NormArea = pNormArea
    objComponente.NumeroAnalisis = pNumeroAnalisis
    objComponente.NumeroComponente = pNumeroComponente
    objComponente.Retention = pRetention
    objComponente.Units = pUnits
    
    Set Duplicado = objComponente
    
Error:
        If Err.Number <> 0 Then
            MsgBox "Todo mal"
            Err.Clear
        End If
End Property


