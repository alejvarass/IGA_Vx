VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlOriginal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Variables privadas para almacenar las dimensiones originales
Private m_LeftOriginal As Single
Private m_TopOriginal As Single
Private m_WidthOriginal As Single
Private m_HeightOriginal As Single

' Propiedades públicas para acceder a las variables
Public Property Let LeftOriginal(ByVal vNewValue As Single)
    m_LeftOriginal = vNewValue
End Property

Public Property Get LeftOriginal() As Single
    LeftOriginal = m_LeftOriginal
End Property

Public Property Let TopOriginal(ByVal vNewValue As Single)
    m_TopOriginal = vNewValue
End Property

Public Property Get TopOriginal() As Single
    TopOriginal = m_TopOriginal
End Property

Public Property Let WidthOriginal(ByVal vNewValue As Single)
    m_WidthOriginal = vNewValue
End Property

Public Property Get WidthOriginal() As Single
    WidthOriginal = m_WidthOriginal
End Property

Public Property Let HeightOriginal(ByVal vNewValue As Single)
    m_HeightOriginal = vNewValue
End Property

Public Property Get HeightOriginal() As Single
    HeightOriginal = m_HeightOriginal
End Property

