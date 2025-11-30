VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigoFactor                       As Long
Private pNombreFactor                       As String

Public Property Get Key() As String
    Key = pCodigoFactor & "ID"
End Property


Public Property Get CodigoFactor() As Long
    CodigoFactor = pCodigoFactor
End Property

Public Property Let CodigoFactor(ByVal vNewValue As Long)
    pCodigoFactor = vNewValue
End Property

Public Property Get NombreFactor() As String
    NombreFactor = pNombreFactor
End Property

Public Property Let NombreFactor(ByVal vNewValue As String)
    pNombreFactor = vNewValue
End Property
