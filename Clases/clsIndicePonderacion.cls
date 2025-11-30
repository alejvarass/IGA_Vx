VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicePonderacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTabla                          As String
Private pCodigo                         As Long
Private pNombre                         As String
Private pValor                          As Double

Public Property Get Key() As String
    Key = pTabla & pCodigo & "ID"
End Property


Public Property Get Tabla() As String
    Tabla = pTabla
End Property

Public Property Let Tabla(ByVal vNewValue As String)
    pTabla = vNewValue
End Property

Public Property Get Codigo() As Long
    Codigo = pCodigo
End Property

Public Property Let Codigo(ByVal vNewValue As Long)
    pCodigo = vNewValue
End Property

Public Property Get nombre() As String
    nombre = pNombre
End Property

Public Property Let nombre(ByVal vNewValue As String)
    pNombre = vNewValue
End Property

Public Property Get Valor() As Double
    Valor = pValor
End Property

Public Property Let Valor(ByVal vNewValue As Double)
    pValor = vNewValue
End Property
