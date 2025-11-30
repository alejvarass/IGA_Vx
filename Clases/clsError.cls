VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDescription        As String
Private pNumber             As String
Private pDestination        As String  '"MSG" Message box (app VB) - "FILE" Archivo de texto
Private pObjeto             As String
Private pMetodo             As String

Public Event ErrorEnDllEvento(ByVal Mensaje As String)

Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(ByVal vNewValue As String)
    pDescription = vNewValue
End Property

Public Property Get Number() As String
    Number = pNumber
End Property

Public Property Let Number(ByVal vNewValue As String)
    pNumber = vNewValue
End Property

Public Property Get Destination() As String
    Destination = pDestination
End Property

Public Property Let Destination(ByVal vNewValue As String)
    pDestination = vNewValue
End Property

Public Property Get Objeto() As String
    Objeto = pObjeto
End Property

Public Property Let Objeto(ByVal vNewValue As String)
    pObjeto = vNewValue
End Property

Public Property Get Metodo() As String
    Metodo = pMetodo
End Property

Public Property Let Metodo(ByVal vNewValue As String)
    pMetodo = vNewValue
End Property

Public Sub Mostrar()
    
    Dim fs                      As New Scripting.FileSystemObject
    Dim archivo                 As TextStream
    
    Set archivo = fs.OpenTextFile(App.Path & "\" & Format(Now, "YYYY-MM-DD") & "LOG.txt", ForAppending, True)
    archivo.WriteLine "Objeto = " & Me.Objeto & " - Método: " & Me.Metodo & "Error: " & Me.Number & " - " & Me.Description
    archivo.Close
    
    RaiseEvent ErrorEnDllEvento("Objeto = " & Me.Objeto & " - Método: " & Me.Metodo & Chr(10) & "Error: " & Me.Number & " - " & Me.Description)

End Sub


