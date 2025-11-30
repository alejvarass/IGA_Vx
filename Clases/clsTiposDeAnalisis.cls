VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTiposDeAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColTiposDeAnalisis          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColTiposDeAnalisis.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColTiposDeAnalisis.Count
End Property

Public Function colAgregar(ByVal objTipoDeAnalisis As clsTipoDeAnalisis) As Boolean
    
    On Error GoTo Error
    pColTiposDeAnalisis.Add objTipoDeAnalisis, objTipoDeAnalisis.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objTipoDeAnalisis As clsTipoDeAnalisis) As Boolean
    
    On Error GoTo Error
    pColTiposDeAnalisis.Remove objTipoDeAnalisis.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarTiposDeAnalisiss()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objTipoDeAnalisis              As clsTipoDeAnalisis
    
    On Error GoTo Error
    StrSql = "SELECT " & _
    "CodigoTipoDeAnalisis, NombreTipoDeAnalisis " & _
    "From " & _
    "Iga_Tipos_De_Analisis " & _
    "order by " & _
    "NombreTipoDeAnalisis"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objTipoDeAnalisis = New clsTipoDeAnalisis
                objTipoDeAnalisis.CodigoTipoDeAnalisis = rs!CodigoTipoDeAnalisis
                objTipoDeAnalisis.NombreTipoDeAnalisis = rs!NombreTipoDeAnalisis
                
                pColTiposDeAnalisis.Add objTipoDeAnalisis, objTipoDeAnalisis.Key
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "ConsultarTiposDeAnalisiss"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colTipoDeAnalisis( _
                Optional ByVal CodigoTipoDeAnalisis As Long, _
                Optional ByVal Key As Variant _
                ) As clsTipoDeAnalisis
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = CodigoTipoDeAnalisis & "ID"
    End If
    Set colTipoDeAnalisis = pColTiposDeAnalisis(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsTiposDeAnalisis"
            pDatos.Error.Metodo = "colTipoDeAnalisis"
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
    
    StrSql = "Select max(CodigoTipoDeAnalisis) as maximo " & _
    "from " & _
    "iga_Tipos_De_Analisis;"

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
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax( _
                Optional ByVal CodigoTipoDeAnalisis As Long _
                ) As Long
    
    Dim max As Long
    Dim objTipoDeAnalisis As clsTipoDeAnalisis
    
    On Error GoTo Error
    max = 0
    For Each objTipoDeAnalisis In pColTiposDeAnalisis
        If objTipoDeAnalisis.CodigoTipoDeAnalisis > max Then
            max = objTipoDeAnalisis.CodigoTipoDeAnalisis
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objTipoDeAnalisis As clsTipoDeAnalisis) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    objTipoDeAnalisis.CodigoTipoDeAnalisis = dbMax + 1
    
    StrSql = "Insert into iga_Tipos_De_Analisis (" & _
    "CodigoTipoDeAnalisis, NombreTipoDeAnalisis " & _
    ") Values (" & _
    objTipoDeAnalisis.CodigoTipoDeAnalisis & " , '" & _
    objTipoDeAnalisis.NombreTipoDeAnalisis & "' " & _
    ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objTipoDeAnalisis As clsTipoDeAnalisis) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_Tipos_De_Analisis Set " & _
    "NombreTipoDeAnalisis = '" & objTipoDeAnalisis.NombreTipoDeAnalisis & "' " & _
    "Where " & _
    "CodigoTipoDeAnalisis = " & objTipoDeAnalisis.CodigoTipoDeAnalisis & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objTipoDeAnalisis As clsTipoDeAnalisis) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_Tipos_De_Analisis " & _
    "Where " & _
    "CodigoTipoDeAnalisis = " & objTipoDeAnalisis.CodigoTipoDeAnalisis & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsTiposDeAnalisis"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function
