VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiposEquipoGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColTiposEquipoGas          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColTiposEquipoGas.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColTiposEquipoGas.Count
End Property

Public Function colAgregar(ByVal objTipoEquipoGas As clsTipoEquipoGas) As Boolean
    
    On Error GoTo Error
    pColTiposEquipoGas.Add objTipoEquipoGas, objTipoEquipoGas.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objTipoEquipoGas As clsTipoEquipoGas) As Boolean
    
    On Error GoTo Error
    pColTiposEquipoGas.Remove objTipoEquipoGas.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarTiposEquipoGas()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objTipoEquipoGas              As clsTipoEquipoGas
    
    On Error GoTo Error
    StrSql = "SELECT " & _
    "IdTipoEquipoGas, TipoEquipoGas " & _
    "From " & _
    "iga_tipos_equipos_gas " & _
    "order by " & _
    "TipoEquipoGas"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objTipoEquipoGas = New clsTipoEquipoGas
                objTipoEquipoGas.IdTipoEquipoGas = rs!IdTipoEquipoGas
                objTipoEquipoGas.TipoEquipoGas = rs!TipoEquipoGas
                
                pColTiposEquipoGas.Add objTipoEquipoGas, objTipoEquipoGas.Key
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "ConsultarTiposEquipoGas"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colTipoEquipoGas( _
                Optional ByVal IdTipoEquipoGas As Long, _
                Optional ByVal Key As Variant _
                ) As clsTipoEquipoGas
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = IdTipoEquipoGas & "ID"
    End If
    Set colTipoEquipoGas = pColTiposEquipoGas(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsTiposEquipoGas"
            pDatos.Error.Metodo = "colTipoEquipoGas"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

Public Property Get dbMax() As Long
    
    Dim StrSql          As String
    Dim rs              As Recordset
    On Error GoTo Error
    
    StrSql = "Select max(IdTipoEquipoGas) as maximo " & _
    "from " & _
    "iga_tipos_equipos_gas;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            If Not IsNull(rs!maximo) Then
                dbMax = rs!maximo
            Else
                dbMax = 0
            End If
        End If
    End If

Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax( _
                Optional ByVal IdTipoEquipoGas As Long _
                ) As Long
    
    Dim max As Long
    Dim objTipoEquipoGas As clsTipoEquipoGas
    
    On Error GoTo Error
    max = 0
    For Each objTipoEquipoGas In pColTiposEquipoGas
        If objTipoEquipoGas.IdTipoEquipoGas > max Then
            max = objTipoEquipoGas.IdTipoEquipoGas
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objTipoEquipoGas As clsTipoEquipoGas) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    objTipoEquipoGas.IdTipoEquipoGas = dbMax + 1
    
    StrSql = "Insert into iga_tipos_equipos_gas (" & _
    "IdTipoEquipoGas, TipoEquipoGas " & _
    ") Values (" & _
    objTipoEquipoGas.IdTipoEquipoGas & " , '" & _
    objTipoEquipoGas.TipoEquipoGas & "' " & _
    ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objTipoEquipoGas As clsTipoEquipoGas) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_tipos_equipos_gas Set " & _
    "TipoEquipoGas = '" & objTipoEquipoGas.TipoEquipoGas & "' " & _
    "Where " & _
    "IdTipoEquipoGas = " & objTipoEquipoGas.IdTipoEquipoGas & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objTipoEquipoGas As clsTipoEquipoGas) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_tipos_equipos_gas " & _
    "Where " & _
    "IdTipoEquipoGas = " & objTipoEquipoGas.IdTipoEquipoGas & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsTiposEquipoGas"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


