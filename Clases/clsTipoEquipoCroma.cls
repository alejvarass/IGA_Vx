VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoEquipoCroma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIdTipoEquipoCroma                       As Long
Private pTipoEquipoCroma                       As String

Public Property Get Key() As String
    Key = pIdTipoEquipoCroma & "ID"
End Property


Public Property Get IdTipoEquipoCroma() As Long
    IdTipoEquipoCroma = pIdTipoEquipoCroma
End Property

Public Property Let IdTipoEquipoCroma(ByVal vNewValue As Long)
    pIdTipoEquipoCroma = vNewValue
End Property

Public Property Get TipoEquipoCroma() As String
    TipoEquipoCroma = pTipoEquipoCroma
End Property

Public Property Let TipoEquipoCroma(ByVal vNewValue As String)
    pTipoEquipoCroma = vNewValue
End Property


