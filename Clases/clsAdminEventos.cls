VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdminEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event EvtActualizarValoresGases()
Public Event EvtActualizarVistaTiempo(ByVal Intervalo As Long)
Public Event EvtActualizarVistaProf()

Public Sub ActualizarValoresGases()
    RaiseEvent EvtActualizarValoresGases
End Sub

Public Sub ActualizarVistaTiempo(ByVal Intervalo As Long)
    RaiseEvent EvtActualizarVistaTiempo(Intervalo)
End Sub

Public Sub ActualizarVistaProf()
    RaiseEvent EvtActualizarVistaProf
End Sub
