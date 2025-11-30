VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigoComponente                   As Long
Private pNombreComponente                   As String
Private pCodigoFactor                       As Long

Public Property Get Key() As String
    Key = pCodigoComponente & "ID"
End Property

Public Property Get CodigoComponente() As Long
    CodigoComponente = pCodigoComponente
End Property

Public Property Let CodigoComponente(ByVal vNewValue As Long)
    pCodigoComponente = vNewValue
End Property

Public Property Get NombreComponente() As String
    NombreComponente = pNombreComponente
End Property

Public Property Let NombreComponente(ByVal vNewValue As String)
    pNombreComponente = vNewValue
End Property

Public Property Get CodigoFactor() As Long
    CodigoFactor = pCodigoFactor
End Property

Public Property Let CodigoFactor(ByVal vNewValue As Long)
    pCodigoFactor = vNewValue
End Property
