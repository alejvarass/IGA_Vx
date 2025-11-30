VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComponentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColComponentes          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColComponentes.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColComponentes.Count
End Property

Public Function colAgregar(ByVal objComponente As clsComponente) As Boolean
    
    On Error GoTo Error
    pColComponentes.Add objComponente, objComponente.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objComponente As clsComponente) As Boolean
    
    On Error GoTo Error
    pColComponentes.Remove objComponente.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarComponentes()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objComponente              As clsComponente
    
    On Error GoTo Error
    StrSql = "SELECT " & _
    "CodigoComponente, NombreComponente, codigofactor " & _
    "From " & _
    "iga_componentes " & _
    "order by " & _
    "NombreComponente"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objComponente = New clsComponente
                objComponente.CodigoComponente = rs!CodigoComponente
                objComponente.NombreComponente = rs!NombreComponente
                objComponente.CodigoFactor = rs!CodigoFactor
                
                pColComponentes.Add objComponente, objComponente.Key
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "ConsultarComponentes"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colComponente( _
                Optional ByVal CodigoComponente As Long, _
                Optional ByVal Key As Variant _
                ) As clsComponente
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = CodigoComponente & "ID"
    End If
    Set colComponente = pColComponentes(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsComponentes"
            pDatos.Error.Metodo = "colComponente"
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
    
    StrSql = "Select max(CodigoComponente) as maximo " & _
    "from " & _
    "iga_componentes;"

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
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax( _
                Optional ByVal CodigoComponente As Long _
                ) As Long
    
    Dim max As Long
    Dim objComponente As clsComponente
    
    On Error GoTo Error
    max = 0
    For Each objComponente In pColComponentes
        If objComponente.CodigoComponente > max Then
            max = objComponente.CodigoComponente
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objComponente As clsComponente) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    objComponente.CodigoComponente = dbMax + 1
    
    StrSql = "Insert into iga_componentes (" & _
    "CodigoComponente, NombreComponente " & _
    ") Values (" & _
    objComponente.CodigoComponente & " , '" & _
    objComponente.NombreComponente & "' " & _
    ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objComponente As clsComponente) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_componentes Set " & _
    "NombreComponente = '" & objComponente.NombreComponente & "' " & _
    "Where " & _
    "CodigoComponente = " & objComponente.CodigoComponente & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objComponente As clsComponente) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_componentes " & _
    "Where " & _
    "CodigoComponente = " & objComponente.CodigoComponente & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsComponentes"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


