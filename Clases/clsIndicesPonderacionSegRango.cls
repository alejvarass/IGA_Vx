VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicesPonderacionSegRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pColIndicesPonderacionSegRango          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColIndicesPonderacionSegRango.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColIndicesPonderacionSegRango.Count
End Property

Public Function colAgregar(ByVal objIndicePonderacionSegRango As clsIndicePonderacionSegRango) As Boolean
    
    On Error GoTo Error
    pColIndicesPonderacionSegRango.Add objIndicePonderacionSegRango, objIndicePonderacionSegRango.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacionSegRango"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objIndicePonderacionSegRango As clsIndicePonderacionSegRango) As Boolean
    
    On Error GoTo Error
    pColIndicesPonderacionSegRango.Remove objIndicePonderacionSegRango.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacionSegRango"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarIndicesPonderacionSegRango()
    
    Dim StrSql                                      As String
    Dim rs                                          As Recordset
    Dim objIndicePonderacionSegRango                As clsIndicePonderacionSegRango
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "codigo, "
    StrSql = StrSql & "inicial, "
    StrSql = StrSql & "final, "
    StrSql = StrSql & "valor, "
    StrSql = StrSql & "tabla "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_indices_ponderacion_segun_rango "
    StrSql = StrSql & "order by "
    StrSql = StrSql & "inicial;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                
                Set objIndicePonderacionSegRango = Me.colIndicePonderacionSegRango(rs!tabla, rs!Codigo)
                If objIndicePonderacionSegRango Is Nothing Then
                    
                    Set objIndicePonderacionSegRango = New clsIndicePonderacionSegRango
                    objIndicePonderacionSegRango.Codigo = rs!Codigo
                    objIndicePonderacionSegRango.Inicial = rs!Inicial
                    objIndicePonderacionSegRango.Final = rs!Final
                    objIndicePonderacionSegRango.tabla = rs!tabla
                    objIndicePonderacionSegRango.valor = rs!valor
                    
                    Me.colAgregar objIndicePonderacionSegRango
                    
                End If
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacionSegRango"
        pDatos.Error.Metodo = "ConsultarIndicesPonderacionSegRango"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Function colIndicesPonderacionSegRango_tabla(ByVal tabla As String) As clsIndicesPonderacionSegRango
    
    Dim objIndicePonderacionSegRango                As clsIndicePonderacionSegRango
    Dim objIndicesPonderacionSegRango              As New clsIndicesPonderacionSegRango
    On Error GoTo Error
    
    For Each objIndicePonderacionSegRango In pColIndicesPonderacionSegRango
        If objIndicePonderacionSegRango.tabla = tabla Then
            objIndicesPonderacionSegRango.colAgregar objIndicePonderacionSegRango
        End If
    Next
    
    Set colIndicesPonderacionSegRango_tabla = objIndicesPonderacionSegRango
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsIndicesPonderacionSegRango"
        pDatos.Error.Metodo = "colIndicesPonderacionSegRango_tabla"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Property Get colIndicePonderacionSegRango( _
                Optional ByVal tabla As String, _
                Optional ByVal Codigo As Long, _
                Optional ByVal Key As Variant _
                ) As clsIndicePonderacionSegRango
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = tabla & Codigo & "ID"
    End If
    
    Set colIndicePonderacionSegRango = pColIndicesPonderacionSegRango(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsIndicesPonderacionSegRango"
            pDatos.Error.Metodo = "colIndicePonderacionSegRango"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

Public Property Get colIndicePonderacionSegRango_valor( _
                ByVal tabla As String, _
                ByVal valor As Long _
                ) As clsIndicePonderacionSegRango
    
    On Error GoTo Error
    
    Dim objIndicePonderacionSegRango As clsIndicePonderacionSegRango
    
    For Each objIndicePonderacionSegRango In pColIndicesPonderacionSegRango
        If objIndicePonderacionSegRango.tabla = tabla Then
            If objIndicePonderacionSegRango.Inicial <= valor And objIndicePonderacionSegRango.Final >= valor Then
                Set colIndicePonderacionSegRango_valor = objIndicePonderacionSegRango
                Exit For
            End If
        End If
    Next
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsIndicesPonderacionSegRango"
            pDatos.Error.Metodo = "colIndicePonderacionSegRango"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

