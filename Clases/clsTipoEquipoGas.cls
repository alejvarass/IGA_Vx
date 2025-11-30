VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoEquipoGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIdTipoEquipoGas                       As Long
Private pTipoEquipoGas                       As String

Public Property Get Key() As String
    Key = pIdTipoEquipoGas & "ID"
End Property


Public Property Get IdTipoEquipoGas() As Long
    IdTipoEquipoGas = pIdTipoEquipoGas
End Property

Public Property Let IdTipoEquipoGas(ByVal vNewValue As Long)
    pIdTipoEquipoGas = vNewValue
End Property

Public Property Get TipoEquipoGas() As String
    TipoEquipoGas = pTipoEquipoGas
End Property

Public Property Let TipoEquipoGas(ByVal vNewValue As String)
    pTipoEquipoGas = vNewValue
End Property
