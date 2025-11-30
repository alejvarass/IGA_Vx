VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLitologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigolitologia                       As Long
Private pNombrelitologia                       As String

Public Property Get Key() As String
    Key = pCodigolitologia & "ID"
End Property


Public Property Get Codigolitologia() As Long
    Codigolitologia = pCodigolitologia
End Property

Public Property Let Codigolitologia(ByVal vNewValue As Long)
    pCodigolitologia = vNewValue
End Property

Public Property Get Nombrelitologia() As String
    Nombrelitologia = pNombrelitologia
End Property

Public Property Let Nombrelitologia(ByVal vNewValue As String)
    pNombrelitologia = vNewValue
End Property


