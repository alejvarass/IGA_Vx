VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoDeAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigoTipoDeAnalisis                       As Long
Private pNombreTipoDeAnalisis                       As String

Public Property Get Key() As String
    Key = pCodigoTipoDeAnalisis & "ID"
End Property


Public Property Get CodigoTipoDeAnalisis() As Long
    CodigoTipoDeAnalisis = pCodigoTipoDeAnalisis
End Property

Public Property Let CodigoTipoDeAnalisis(ByVal vNewValue As Long)
    pCodigoTipoDeAnalisis = vNewValue
End Property

Public Property Get NombreTipoDeAnalisis() As String
    NombreTipoDeAnalisis = pNombreTipoDeAnalisis
End Property

Public Property Let NombreTipoDeAnalisis(ByVal vNewValue As String)
    pNombreTipoDeAnalisis = vNewValue
End Property

