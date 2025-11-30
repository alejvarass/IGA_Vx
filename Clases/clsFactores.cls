VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFactores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColFactores          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColFactores.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColFactores.Count
End Property

Public Function colAgregar(ByVal objFactor As clsFactor) As Boolean
    
    On Error GoTo Error
    pColFactores.Add objFactor, objFactor.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objFactor As clsFactor) As Boolean
    
    On Error GoTo Error
    pColFactores.Remove objFactor.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarFactores()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objFactor              As clsFactor
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "CodigoFactor, NombreFactor "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_Factores "
    StrSql = StrSql & "order by "
    StrSql = StrSql & "NombreFactor;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objFactor = New clsFactor
                objFactor.CodigoFactor = rs!CodigoFactor
                objFactor.NombreFactor = rs!NombreFactor
                
                pColFactores.Add objFactor, objFactor.Key
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "ConsultarFactores"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colFactor( _
                Optional ByVal CodigoFactor As Long, _
                Optional ByVal Key As Variant _
                ) As clsFactor
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = CodigoFactor & "ID"
    End If
    Set colFactor = pColFactores(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsFactores"
            pDatos.Error.Metodo = "colFactor"
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
    
    StrSql = "Select max(CodigoFactor) as maximo "
    StrSql = StrSql & "from "
    StrSql = StrSql & "iga_Factores;"

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
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax() As Long
    
    Dim max As Long
    Dim objFactor As clsFactor
    
    On Error GoTo Error
    max = 0
    For Each objFactor In pColFactores
        If objFactor.CodigoFactor > max Then
            max = objFactor.CodigoFactor
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objFactor As clsFactor) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "Insert into iga_Factores ("
    StrSql = StrSql & "CodigoFactor, "
    StrSql = StrSql & "NombreFactor "
    StrSql = StrSql & ") Values ("
    StrSql = StrSql & objFactor.CodigoFactor & " , '"
    StrSql = StrSql & objFactor.NombreFactor & "' "
    StrSql = StrSql & ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objFactor As clsFactor) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_Factores Set "
    StrSql = StrSql & "NombreFactor = '" & objFactor.NombreFactor & "' "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "CodigoFactor = " & objFactor.CodigoFactor & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objFactor As clsFactor) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_Factores "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "CodigoFactor = " & objFactor.CodigoFactor & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsFactores"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

