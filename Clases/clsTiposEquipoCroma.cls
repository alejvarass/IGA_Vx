VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiposEquipoCroma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColTiposEquipoCroma          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColTiposEquipoCroma.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColTiposEquipoCroma.Count
End Property

Public Function colAgregar(ByVal objTipoEquipoCroma As clsTipoEquipoCroma) As Boolean
    
    On Error GoTo Error
    pColTiposEquipoCroma.Add objTipoEquipoCroma, objTipoEquipoCroma.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objTipoEquipoCroma As clsTipoEquipoCroma) As Boolean
    
    On Error GoTo Error
    pColTiposEquipoCroma.Remove objTipoEquipoCroma.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarTiposEquipoCroma()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objTipoEquipoCroma              As clsTipoEquipoCroma
    
    On Error GoTo Error
    StrSql = "SELECT " & _
    "IdTipoEquipoCroma, TipoEquipoCroma " & _
    "From " & _
    "iga_tipos_equipos_Croma " & _
    "order by " & _
    "TipoEquipoCroma"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objTipoEquipoCroma = New clsTipoEquipoCroma
                objTipoEquipoCroma.IdTipoEquipoCroma = rs!IdTipoEquipoCroma
                objTipoEquipoCroma.TipoEquipoCroma = rs!TipoEquipoCroma
                
                pColTiposEquipoCroma.Add objTipoEquipoCroma, objTipoEquipoCroma.Key
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "ConsultarTiposEquipoCroma"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colTipoEquipoCroma( _
                Optional ByVal IdTipoEquipoCroma As Long, _
                Optional ByVal Key As Variant _
                ) As clsTipoEquipoCroma
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = IdTipoEquipoCroma & "ID"
    End If
    Set colTipoEquipoCroma = pColTiposEquipoCroma(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsTiposEquipoCroma"
            pDatos.Error.Metodo = "colTipoEquipoCroma"
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
    
    StrSql = "Select max(IdTipoEquipoCroma) as maximo " & _
    "from " & _
    "iga_tipos_equipos_Croma;"

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
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax( _
                Optional ByVal IdTipoEquipoCroma As Long _
                ) As Long
    
    Dim max As Long
    Dim objTipoEquipoCroma As clsTipoEquipoCroma
    
    On Error GoTo Error
    max = 0
    For Each objTipoEquipoCroma In pColTiposEquipoCroma
        If objTipoEquipoCroma.IdTipoEquipoCroma > max Then
            max = objTipoEquipoCroma.IdTipoEquipoCroma
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objTipoEquipoCroma As clsTipoEquipoCroma) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    objTipoEquipoCroma.IdTipoEquipoCroma = dbMax + 1
    
    StrSql = "Insert into iga_tipos_equipos_Croma (" & _
    "IdTipoEquipoCroma, TipoEquipoCroma " & _
    ") Values (" & _
    objTipoEquipoCroma.IdTipoEquipoCroma & " , '" & _
    objTipoEquipoCroma.TipoEquipoCroma & "' " & _
    ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objTipoEquipoCroma As clsTipoEquipoCroma) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_tipos_equipos_Croma Set " & _
    "TipoEquipoCroma = '" & objTipoEquipoCroma.TipoEquipoCroma & "' " & _
    "Where " & _
    "IdTipoEquipoCroma = " & objTipoEquipoCroma.IdTipoEquipoCroma & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objTipoEquipoCroma As clsTipoEquipoCroma) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_tipos_equipos_Croma " & _
    "Where " & _
    "IdTipoEquipoCroma = " & objTipoEquipoCroma.IdTipoEquipoCroma & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsTiposEquipoCroma"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



