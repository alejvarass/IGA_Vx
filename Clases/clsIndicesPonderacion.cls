VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicesPonderacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColIndicesPonderacion          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColIndicesPonderacion.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColIndicesPonderacion.Count
End Property

Public Function colAgregar(ByVal objIndicePonderacion As clsIndicePonderacion) As Boolean
    
    On Error GoTo Error
    pColIndicesPonderacion.Add objIndicePonderacion, objIndicePonderacion.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacion"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objIndicePonderacion As clsIndicePonderacion) As Boolean
    
    On Error GoTo Error
    pColIndicesPonderacion.Remove objIndicePonderacion.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacion"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarIndicesPonderacion()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objIndicePonderacion              As clsIndicePonderacion
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "codigo, "
    StrSql = StrSql & "nombre, "
    StrSql = StrSql & "valor, "
    StrSql = StrSql & "tabla "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_Indices_Ponderacion "
    StrSql = StrSql & "order by "
    StrSql = StrSql & "nombre;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                
                Set objIndicePonderacion = Me.colIndicePonderacion(rs!tabla, rs!Codigo)
                If objIndicePonderacion Is Nothing Then
                    
                    Set objIndicePonderacion = New clsIndicePonderacion
                    objIndicePonderacion.Codigo = rs!Codigo
                    objIndicePonderacion.nombre = rs!nombre
                    objIndicePonderacion.tabla = rs!tabla
                    objIndicePonderacion.Valor = rs!Valor
                    
                    Me.colAgregar objIndicePonderacion
                    
                End If
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacion"
        pDatos.Error.Metodo = "ConsultarIndicesPonderacion"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Function colIndicesPonderacion_tabla(ByVal tabla As String) As clsIndicesPonderacion
    
    Dim objIndicePonderacion                As clsIndicePonderacion
    Dim objIndicesPonederacion              As New clsIndicesPonderacion
    On Error GoTo Error
    
    For Each objIndicePonderacion In pColIndicesPonderacion
        If objIndicePonderacion.tabla = tabla Then
            objIndicesPonederacion.colAgregar objIndicePonderacion
        End If
    Next
    
    Set colIndicesPonderacion_tabla = objIndicesPonederacion
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacion"
        pDatos.Error.Metodo = "colIndicesPonderacion_tabla"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Property Get colIndicePonderacion( _
                Optional ByVal tabla As String, _
                Optional ByVal Codigo As Long, _
                Optional ByVal Key As Variant _
                ) As clsIndicePonderacion
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = tabla & Codigo & "ID"
    End If
    
    Set colIndicePonderacion = pColIndicesPonderacion(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsIndicesPonderacion"
            pDatos.Error.Metodo = "colIndicePonderacion"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property
